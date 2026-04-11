"""
Little Helper - System Tray Application (main entry point).
Clipboard image paster + screenshot tool + GPU power control + system overlay.
"""

import os
import sys
import ctypes
import ctypes.wintypes

# Enable DPI awareness (must happen before any UI)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

import threading
import logging
import time

import tkinter as tk
import pystray
import win32gui
from PIL import Image

import config as cfg
import clipboard_paste
import screenshot as screenshot_mod
import hotkey
import gpu_power
import system_overlay
import fan_control

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_PATH = cfg.get_log_path()

# Rotate previous log instead of clearing it, so crash logs are preserved
_bak_path = LOG_PATH + ".bak"
if os.path.exists(LOG_PATH):
    try:
        if os.path.exists(_bak_path):
            os.remove(_bak_path)
        os.rename(LOG_PATH, _bak_path)
    except Exception:
        open(LOG_PATH, "w", encoding="utf-8").close()

logging.getLogger().setLevel(logging.WARNING)

log = logging.getLogger("little_helper")
log.setLevel(logging.DEBUG)
log.propagate = False

_fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
_fh.setLevel(logging.DEBUG)
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s %(message)s"))
log.addHandler(_fh)

# Propagate to sub-module loggers
for _mod in ("little_helper.config", "little_helper.clipboard_paste",
             "little_helper.screenshot", "little_helper.hotkey",
             "little_helper.gpu_power", "little_helper.system_overlay",
             "little_helper.fan_control"):
    _ml = logging.getLogger(_mod)
    _ml.setLevel(logging.DEBUG)
    _ml.propagate = False
    _ml.addHandler(_fh)


def _log_unhandled(exc_type, exc_value, exc_tb):
    log.critical("Unhandled exception", exc_info=(exc_type, exc_value, exc_tb))
    _fh.flush()


def _log_thread_exception(args):
    log.critical(
        f"Unhandled exception in thread {args.thread!r}",
        exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
    )
    _fh.flush()


sys.excepthook = _log_unhandled
threading.excepthook = _log_thread_exception

# ---------------------------------------------------------------------------
# Global state
# ---------------------------------------------------------------------------

_config: dict    = {}
_tray_icon       = None
_settings_root   = None

WM_CLOSE         = 0x0010
_WINDOW_TITLE    = "LittleHelper_Hidden"
_WINDOW_CLASS    = "LittleHelper_WndClass"

# ---------------------------------------------------------------------------
# Notification helper
# ---------------------------------------------------------------------------

def _notify(message: str, title: str = "Little Helper") -> None:
    if _tray_icon:
        try:
            _tray_icon.notify(message, title)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Shutdown
# ---------------------------------------------------------------------------

def _shutdown() -> None:
    """Unified cleanup: close overlay, restore GPU, stop hook."""
    log.info("Shutting down...")
    if _settings_root is not None:
        try:
            _settings_root.destroy()
        except Exception:
            pass
    system_overlay.close_overlay()
    fan_control.stop_fan_control()
    fan_control.stop_gpu_fan_control()
    gpu_power.restore_gpu_power_limit()
    hotkey.stop_keyboard_hook()


# ---------------------------------------------------------------------------
# Single-instance hidden window
# ---------------------------------------------------------------------------

def kill_previous_instance() -> None:
    import win32process
    current_pid = os.getpid()
    _u32 = ctypes.windll.user32

    def _enum(hwnd, _):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid != current_pid and win32gui.GetWindowText(hwnd) == _WINDOW_TITLE:
                log.info(f"Found previous instance (pid={pid}), sending WM_CLOSE")
                _u32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_enum, None)
    time.sleep(0.5)


def _on_hidden_wnd_close(hwnd, msg, wp, lp):
    log.info("WM_CLOSE received, shutting down")
    _shutdown()
    if _tray_icon:
        _tray_icon.stop()
    win32gui.DestroyWindow(hwnd)
    return 0


def create_hidden_window() -> int:
    import win32api
    wc = win32gui.WNDCLASS()
    wc.lpfnWndProc  = {WM_CLOSE: _on_hidden_wnd_close}
    wc.lpszClassName = _WINDOW_CLASS
    wc.hInstance    = win32api.GetModuleHandle(None)
    try:
        cls = win32gui.RegisterClass(wc)
    except Exception:
        cls = win32gui.RegisterClass(wc)
    hwnd = win32gui.CreateWindow(
        cls, _WINDOW_TITLE, 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None
    )
    log.info(f"Hidden window created (hwnd={hwnd})")
    return hwnd


# ---------------------------------------------------------------------------
# Settings dialog
# ---------------------------------------------------------------------------

def _force_window_focus(root: tk.Tk) -> None:
    """Reliably bring a tkinter window to the foreground on Windows."""
    root.lift()
    root.focus_force()
    try:
        hwnd = root.winfo_id()
        ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def show_settings_dialog() -> None:
    # If already open, just bring it to front and return immediately so the
    # pystray callback thread is never blocked.
    if _settings_root is not None:
        try:
            # Schedule via tkinter's own thread to stay thread-safe.
            _settings_root.after(0, lambda: _force_window_focus(_settings_root))
        except Exception:
            pass
        return

    # Run the dialog in a dedicated daemon thread so the pystray message pump
    # is never blocked (a blocked pump means the window gets no Win32 messages
    # and is completely unresponsive).
    def _run():
        global _settings_root

        _initing = [True]   # guard: suppress side-effect callbacks during widget setup

        root = tk.Tk()
        _settings_root = root
        root.title("Settings - Little Helper")
        root.resizable(False, False)

        width, height = 420, 780
        root.update_idletasks()
        x = (root.winfo_screenwidth()  // 2) - (width  // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"{width}x{height}+{x}+{y}")

        # Grab focus: set topmost briefly to cut through Windows' focus-steal
        # prevention, then clear it so normal alt-tab behaviour is preserved.
        def _acquire_focus():
            root.attributes("-topmost", True)
            _force_window_focus(root)
            root.after(400, lambda: root.attributes("-topmost", False))

        root.after(50, _acquire_focus)

        outer = tk.Frame(root, padx=16, pady=12)
        outer.pack(fill="both", expand=True)

        def _section(parent, title):
            tk.Label(parent, text=title, anchor="w", font=("Arial", 9, "bold")).pack(
                fill="x", pady=(8, 2)
            )
            tk.Frame(parent, bg="#cccccc", height=1).pack(fill="x")
            f = tk.Frame(parent, pady=4)
            f.pack(fill="x")
            return f

        # ── Section 1: Hotkeys ────────────────────────────────────────────────
        hk_frame = _section(outer, "Hotkeys")

        paste_mod = tk.StringVar(master=root, value=_config["paste_hotkey"]["modifier"].capitalize())
        paste_key = tk.StringVar(master=root, value=_config["paste_hotkey"]["key"].upper())

        def _make_key_entry(parent, var):
            e = tk.Entry(parent, textvariable=var, width=5)
            e.bind("<FocusIn>", lambda _: e.after(0, lambda: (e.select_range(0, "end"), e.icursor("end"))))
            e.bind("<Key>", lambda ev: (var.set(ev.keysym.upper()[:1]) if len(ev.char) == 1 and ev.char.isprintable() else None) or "break")
            return e

        row0 = tk.Frame(hk_frame)
        row0.pack(fill="x", pady=3)
        tk.Label(row0, text="Paste:", width=14, anchor="w").pack(side="left")
        tk.OptionMenu(row0, paste_mod, "Ctrl", "Alt").pack(side="left", padx=4)
        _make_key_entry(row0, paste_key).pack(side="left")

        ss_mod = tk.StringVar(master=root, value=_config["screenshot_hotkey"]["modifier"].capitalize())
        ss_key = tk.StringVar(master=root, value=_config["screenshot_hotkey"]["key"].upper())

        row1 = tk.Frame(hk_frame)
        row1.pack(fill="x", pady=3)
        tk.Label(row1, text="Screenshot:", width=14, anchor="w").pack(side="left")
        tk.OptionMenu(row1, ss_mod, "Ctrl", "Alt").pack(side="left", padx=4)
        _make_key_entry(row1, ss_key).pack(side="left")

        def _apply_hotkey_modifiers(*_):
            if _initing[0]: return
            _config["paste_hotkey"]["modifier"]      = paste_mod.get().lower()
            _config["screenshot_hotkey"]["modifier"] = ss_mod.get().lower()
            cfg.save_config(_config)

        paste_mod.trace_add("write", _apply_hotkey_modifiers)
        ss_mod.trace_add("write", _apply_hotkey_modifiers)

        # ── Section 2: GPU Power Limit ────────────────────────────────────────
        gpu_frame = _section(outer, "GPU Power Limit  (Nvidia only)")

        gpu_enabled = tk.BooleanVar(master=root, value=_config["gpu_power_limit"]["enabled"])
        gpu_watts   = tk.IntVar(   master=root, value=_config["gpu_power_limit"]["watts"])

        # Query GPU limits for Spinbox range
        limits = gpu_power.get_gpu_power_limits()
        w_min, w_max = (50, 500) if limits is None else (int(limits[0]), int(limits[1]))

        row2 = tk.Frame(gpu_frame)
        row2.pack(fill="x", pady=3)
        gpu_cb = tk.Checkbutton(row2, text="Enable on startup", variable=gpu_enabled)
        gpu_cb.pack(side="left")

        row3 = tk.Frame(gpu_frame)
        row3.pack(fill="x", pady=3)
        tk.Label(row3, text="Target watts:", width=14, anchor="w").pack(side="left")
        gpu_spin = tk.Spinbox(row3, from_=w_min, to=w_max, increment=5,
                              textvariable=gpu_watts, width=6)
        gpu_spin.pack(side="left", padx=4)
        tk.Label(row3, text=f"(GPU range: {w_min}–{w_max} W)",
                 font=("Arial", 8), fg="gray").pack(side="left")

        def _toggle_spin(*_):
            gpu_spin.configure(state="normal" if gpu_enabled.get() else "disabled")
        gpu_enabled.trace_add("write", _toggle_spin)
        _toggle_spin()

        def _apply_gpu_power(*_):
            if _initing[0]: return
            _config["gpu_power_limit"]["enabled"] = gpu_enabled.get()
            _config["gpu_power_limit"]["watts"]   = gpu_watts.get()
            cfg.save_config(_config)
            gpu_power.apply_gpu_power_limit(_config, notify_fn=_notify)

        gpu_enabled.trace_add("write", _apply_gpu_power)
        gpu_watts.trace_add("write", _apply_gpu_power)

        # ── Section 3: Overlay ────────────────────────────────────────────────
        ov_frame = _section(outer, "System Monitor Overlay")

        ov_enabled  = tk.BooleanVar(master=root, value=_config["overlay"]["enabled"])
        ov_opacity  = tk.DoubleVar( master=root, value=_config["overlay"]["opacity"])
        ov_refresh  = tk.IntVar(    master=root, value=_config["overlay"]["refresh_ms"])

        row4 = tk.Frame(ov_frame)
        row4.pack(fill="x", pady=3)
        tk.Checkbutton(row4, text="Show overlay on startup",
                       variable=ov_enabled).pack(side="left")

        row5 = tk.Frame(ov_frame)
        row5.pack(fill="x", pady=3)
        tk.Label(row5, text="Opacity:", width=14, anchor="w").pack(side="left")
        tk.Spinbox(row5, from_=0.20, to=1.00, increment=0.05, format="%.2f",
                   textvariable=ov_opacity, width=6).pack(side="left", padx=4)

        row6 = tk.Frame(ov_frame)
        row6.pack(fill="x", pady=3)
        tk.Label(row6, text="Refresh (ms):", width=14, anchor="w").pack(side="left")
        tk.Spinbox(row6, from_=100, to=5000, increment=100,
                   textvariable=ov_refresh, width=6).pack(side="left", padx=4)

        def _apply_overlay(*_):
            if _initing[0]: return
            _config["overlay"]["enabled"]    = ov_enabled.get()
            _config["overlay"]["opacity"]    = round(ov_opacity.get(), 2)
            _config["overlay"]["refresh_ms"] = ov_refresh.get()
            cfg.save_config(_config)
            system_overlay.apply_overlay_opacity(_config["overlay"]["opacity"])

        ov_enabled.trace_add("write", _apply_overlay)
        ov_opacity.trace_add("write", _apply_overlay)
        ov_refresh.trace_add("write", _apply_overlay)

        # ── Section 4: CPU Fan Control ────────────────────────────────────────
        fc_frame = _section(outer, "CPU Fan Control")

        fc_cfg      = _config.get("fan_control", {})
        fc_enabled  = tk.BooleanVar(master=root, value=fc_cfg.get("enabled", False))
        fc_source   = tk.StringVar(master=root, value=fc_cfg.get("source", "gpu_temp"))
        fc_interval = tk.IntVar(   master=root, value=fc_cfg.get("interval_s", 3))
        fc_manual   = tk.DoubleVar(master=root, value=fc_cfg.get("manual_pct", 50))

        row_fc0 = tk.Frame(fc_frame)
        row_fc0.pack(fill="x", pady=3)
        tk.Checkbutton(row_fc0, text="Enable on startup", variable=fc_enabled).pack(side="left")

        row_fc1 = tk.Frame(fc_frame)
        row_fc1.pack(fill="x", pady=3)
        tk.Label(row_fc1, text="Source:", width=14, anchor="w").pack(side="left")
        tk.OptionMenu(row_fc1, fc_source, "cpu_temp", "gpu_temp", "mixed", "manual").pack(
            side="left", padx=4
        )

        row_fc2 = tk.Frame(fc_frame)
        row_fc2.pack(fill="x", pady=3)
        tk.Label(row_fc2, text="Interval (s):", width=14, anchor="w").pack(side="left")
        tk.Spinbox(row_fc2, from_=1, to=60, increment=1,
                   textvariable=fc_interval, width=5).pack(side="left", padx=4)

        # Manual fan speed slider (only visible when source == "manual")
        row_fc3 = tk.Frame(fc_frame)
        manual_pct_label = tk.Label(row_fc3, text="Fan speed:", width=14, anchor="w")
        manual_pct_label.pack(side="left")
        manual_val_label = tk.Label(row_fc3, text=f"{int(fc_manual.get())}%", width=5, anchor="w")

        def _on_manual_slider(val):
            # Only update the display label; fan speed is applied by the poll loop
            manual_val_label.configure(text=f"{int(float(val))}%")

        manual_slider = tk.Scale(
            row_fc3, variable=fc_manual, from_=0, to=100, resolution=1,
            orient="horizontal", length=180, showvalue=False, command=_on_manual_slider,
        )
        manual_slider.pack(side="left", padx=4)
        manual_val_label.pack(side="left")

        # Poll slider value every 400 ms and push to fan control when in manual mode
        _poll_id = [None]

        def _poll_manual():
            if fc_source.get() == "manual":
                fan_control.set_manual_pct(float(fc_manual.get()))
            _poll_id[0] = root.after(400, _poll_manual)

        _poll_manual()

        def _apply_source_immediately(*_):
            source = fc_source.get()
            if source == "manual":
                row_fc3.pack(fill="x", pady=3)
            else:
                row_fc3.pack_forget()
            if _initing[0]: return
            _config["fan_control"]["source"] = source
            cfg.save_config(_config)
            if fan_control.fan_control_is_active():
                fan_control.stop_fan_control()
                lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
                if lhm_computer is not None:
                    fan_control.start_fan_control(_config, lhm_computer, lhm_lock)

        fc_source.trace_add("write", _apply_source_immediately)
        _apply_source_immediately()  # set initial visibility

        def _apply_fc_enabled(*_):
            if _initing[0]: return
            _config["fan_control"]["enabled"] = fc_enabled.get()
            cfg.save_config(_config)
            if fc_enabled.get():
                if not fan_control.fan_control_is_active():
                    lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
                    if lhm_computer is not None:
                        fan_control.start_fan_control(_config, lhm_computer, lhm_lock)
            else:
                fan_control.stop_fan_control()

        fc_enabled.trace_add("write", _apply_fc_enabled)

        def _apply_fc_interval(*_):
            if _initing[0]: return
            _config["fan_control"]["interval_s"] = fc_interval.get()
            cfg.save_config(_config)
            if fan_control.fan_control_is_active():
                fan_control.stop_fan_control()
                lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
                if lhm_computer is not None:
                    fan_control.start_fan_control(_config, lhm_computer, lhm_lock)

        fc_interval.trace_add("write", _apply_fc_interval)

        # ── Section 5: GPU Fan Control ────────────────────────────────────────
        gfc_frame = _section(outer, "GPU Fan Control  (Nvidia only)")

        gfc_cfg      = _config.get("gpu_fan_control", {})
        gfc_enabled  = tk.BooleanVar(master=root, value=gfc_cfg.get("enabled", False))
        gfc_source   = tk.StringVar( master=root, value=gfc_cfg.get("source", "gpu_temp"))
        gfc_interval = tk.IntVar(    master=root, value=gfc_cfg.get("interval_s", 2))
        gfc_manual   = tk.DoubleVar( master=root, value=gfc_cfg.get("manual_pct", 50))

        row_gfc0 = tk.Frame(gfc_frame)
        row_gfc0.pack(fill="x", pady=3)
        tk.Checkbutton(row_gfc0, text="Enable on startup", variable=gfc_enabled).pack(side="left")

        row_gfc1 = tk.Frame(gfc_frame)
        row_gfc1.pack(fill="x", pady=3)
        tk.Label(row_gfc1, text="Source:", width=14, anchor="w").pack(side="left")
        tk.OptionMenu(row_gfc1, gfc_source, "gpu_temp", "manual").pack(side="left", padx=4)

        row_gfc2 = tk.Frame(gfc_frame)
        row_gfc2.pack(fill="x", pady=3)
        tk.Label(row_gfc2, text="Interval (s):", width=14, anchor="w").pack(side="left")
        tk.Spinbox(row_gfc2, from_=1, to=60, increment=1,
                   textvariable=gfc_interval, width=5).pack(side="left", padx=4)

        # Manual slider (only visible when source == "manual")
        row_gfc3 = tk.Frame(gfc_frame)
        gfc_manual_pct_label = tk.Label(row_gfc3, text="Fan speed:", width=14, anchor="w")
        gfc_manual_pct_label.pack(side="left")
        gfc_manual_val_label = tk.Label(row_gfc3, text=f"{int(gfc_manual.get())}%", width=5, anchor="w")

        def _on_gfc_manual_slider(val):
            gfc_manual_val_label.configure(text=f"{int(float(val))}%")

        gfc_manual_slider = tk.Scale(
            row_gfc3, variable=gfc_manual, from_=0, to=100, resolution=1,
            orient="horizontal", length=180, showvalue=False, command=_on_gfc_manual_slider,
        )
        gfc_manual_slider.pack(side="left", padx=4)
        gfc_manual_val_label.pack(side="left")

        _gfc_poll_id = [None]

        def _poll_gfc_manual():
            if gfc_source.get() == "manual":
                fan_control.set_gpu_manual_pct(float(gfc_manual.get()))
            _gfc_poll_id[0] = root.after(400, _poll_gfc_manual)

        _poll_gfc_manual()

        def _apply_gfc_source_immediately(*_):
            source = gfc_source.get()
            if source == "manual":
                row_gfc3.pack(fill="x", pady=3)
            else:
                row_gfc3.pack_forget()
            if _initing[0]: return
            _config["gpu_fan_control"]["source"] = source
            cfg.save_config(_config)
            if fan_control.gpu_fan_control_is_active():
                fan_control.stop_gpu_fan_control()
                fan_control.start_gpu_fan_control(_config)

        gfc_source.trace_add("write", _apply_gfc_source_immediately)
        _apply_gfc_source_immediately()  # set initial visibility

        def _apply_gfc_enabled(*_):
            if _initing[0]: return
            _config["gpu_fan_control"]["enabled"] = gfc_enabled.get()
            cfg.save_config(_config)
            if gfc_enabled.get():
                if not fan_control.gpu_fan_control_is_active():
                    fan_control.start_gpu_fan_control(_config)
            else:
                fan_control.stop_gpu_fan_control()

        gfc_enabled.trace_add("write", _apply_gfc_enabled)

        def _apply_gfc_interval(*_):
            if _initing[0]: return
            _config["gpu_fan_control"]["interval_s"] = gfc_interval.get()
            cfg.save_config(_config)
            if fan_control.gpu_fan_control_is_active():
                fan_control.stop_gpu_fan_control()
                fan_control.start_gpu_fan_control(_config)

        gfc_interval.trace_add("write", _apply_gfc_interval)

        # ── Close: save key fields and slider defaults ────────────────────────
        def _close():
            global _settings_root
            dirty = False
            for path, new_val in [
                (("paste_hotkey",      "key"),        paste_key.get().upper()),
                (("screenshot_hotkey", "key"),        ss_key.get().upper()),
                (("fan_control",       "manual_pct"), int(fc_manual.get())),
                (("gpu_fan_control",   "manual_pct"), int(gfc_manual.get())),
            ]:
                section, key = path
                if _config[section][key] != new_val:
                    _config[section][key] = new_val
                    dirty = True
            if dirty:
                cfg.save_config(_config)
            if _poll_id[0]:
                root.after_cancel(_poll_id[0])
            if _gfc_poll_id[0]:
                root.after_cancel(_gfc_poll_id[0])
            _settings_root = None
            root.destroy()

        root.protocol("WM_DELETE_WINDOW", _close)

        # ── Bottom buttons ────────────────────────────────────────────────────
        btn_frame = tk.Frame(outer, pady=6)
        btn_frame.pack(fill="x")

        def _open_config():
            import subprocess
            subprocess.Popen(["notepad.exe", cfg.get_config_path()])

        tk.Button(btn_frame, text="Open config.json", command=_open_config).pack(side="left")

        # All widgets built — allow trace callbacks to fire from now on
        _initing[0] = False

        root.mainloop()

    threading.Thread(target=_run, daemon=True).start()


# ---------------------------------------------------------------------------
# Tray icon
# ---------------------------------------------------------------------------

def create_tray_icon() -> None:
    global _tray_icon

    icon_path  = cfg.get_resource_path("icon.ico")
    icon_image = Image.open(icon_path) if os.path.exists(icon_path) \
                 else Image.new("RGB", (64, 64), (0, 150, 80))

    def _on_exit(icon, item):
        log.info("Exit requested via tray")
        _shutdown()
        icon.stop()

    def _on_settings(icon, item):
        show_settings_dialog()

    def _on_toggle_overlay(icon, item):
        threading.Thread(
            target=system_overlay.toggle_overlay,
            args=(_config, cfg.save_config),
            kwargs={"on_close_fn": icon.update_menu},
            daemon=True,
        ).start()

    def _on_toggle_fan_control(icon, item):
        if fan_control.fan_control_is_active():
            fan_control.stop_fan_control()
        else:
            lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
            if lhm_computer is not None:
                fan_control.start_fan_control(_config, lhm_computer, lhm_lock)
            else:
                _notify(
                    "LibreHardwareMonitor not available.\n"
                    "CPU fan control requires LHM to be initialised.",
                    "CPU Fan Control",
                )

    def _on_toggle_gpu_fan_control(icon, item):
        if fan_control.gpu_fan_control_is_active():
            fan_control.stop_gpu_fan_control()
        else:
            fan_control.start_gpu_fan_control(_config)

    def _get_hotkey_display():
        p = _config["paste_hotkey"]
        s = _config["screenshot_hotkey"]
        return (f"Paste: {p['modifier'].capitalize()}+{p['key'].upper()}"
                f"  |  Shot: {s['modifier'].capitalize()}+{s['key'].upper()}")

    menu = pystray.Menu(
        pystray.MenuItem("Little Helper", None, enabled=False),
        pystray.MenuItem(lambda text: _get_hotkey_display(), None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "Overlay",
            _on_toggle_overlay,
            checked=lambda item: system_overlay.overlay_is_open(),
        ),
        pystray.MenuItem(
            "CPU Fan Control",
            _on_toggle_fan_control,
            checked=lambda item: fan_control.fan_control_is_enabled(),
        ),
        pystray.MenuItem(
            "GPU Fan Control",
            _on_toggle_gpu_fan_control,
            checked=lambda item: fan_control.gpu_fan_control_is_enabled(),
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Settings", _on_settings, default=True),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Exit", _on_exit),
    )

    icon = pystray.Icon("little_helper", icon_image, "Little Helper", menu)
    _tray_icon = icon

    # Paste callback (closure captures _config and _notify)
    def _paste_cb():
        clipboard_paste.on_paste(_config, notify_fn=_notify)

    def _screenshot_cb():
        screenshot_mod.on_screenshot(_config, notify_fn=_notify)

    # Start keyboard hook in daemon thread
    threading.Thread(
        target=hotkey.start_keyboard_hook,
        args=(_config, _paste_cb, _screenshot_cb),
        daemon=True,
    ).start()

    def on_setup(icon):
        icon.visible = True
        # Apply GPU power limit after tray is visible (so notify works)
        gpu_power.apply_gpu_power_limit(_config, notify_fn=_notify)
        # Sync manual_pct from config into the module
        fan_control.set_manual_pct(_config.get("fan_control", {}).get("manual_pct", 50))
        fan_control.set_gpu_manual_pct(_config.get("gpu_fan_control", {}).get("manual_pct", 50))
        # Auto-start CPU fan control if configured
        if _config.get("fan_control", {}).get("enabled", False):
            lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
            if lhm_computer is not None:
                fan_control.start_fan_control(_config, lhm_computer, lhm_lock)
            else:
                log.warning("CPU fan control enabled in config but LHM is not available")
        # Auto-start GPU fan control if configured
        if _config.get("gpu_fan_control", {}).get("enabled", False):
            fan_control.start_gpu_fan_control(_config)
        # Auto-open overlay if configured
        if _config.get("overlay", {}).get("enabled", False):
            threading.Thread(
                target=system_overlay.toggle_overlay,
                args=(_config, cfg.save_config),
                kwargs={"on_close_fn": icon.update_menu},
                daemon=True,
            ).start()
        # Refresh tray menu checkmarks after all state has been set
        icon.update_menu()

    log.info("Starting tray icon...")
    icon.run(setup=on_setup)
    log.info("Tray icon stopped")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Require administrator privileges before anything else
    if not gpu_power.is_admin():
        gpu_power.relaunch_as_admin()
        sys.exit(0)

    log.info("=" * 50)
    log.info("Little Helper starting")
    log.info(f"Log: {LOG_PATH}")

    _config = cfg.load_config()
    log.info(
        f"Config loaded: paste={_config['paste_hotkey']['modifier']}+"
        f"{_config['paste_hotkey']['key']}, "
        f"screenshot={_config['screenshot_hotkey']['modifier']}+"
        f"{_config['screenshot_hotkey']['key']}"
    )

    kill_previous_instance()

    system_overlay.init_nvml()
    system_overlay.init_lhm()

    try:
        _hidden_hwnd = create_hidden_window()
        create_tray_icon()
    except Exception as e:
        log.error(f"Fatal error: {e}", exc_info=True)
