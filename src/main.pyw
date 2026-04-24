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
import queue

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
import monitor_server
import fan_control
import auto_sleep

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

logging.getLogger().setLevel(logging.DEBUG)

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
             "little_helper.monitor_server", "little_helper.fan_control",
             "little_helper.auto_sleep"):
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
_settings_overlay_enabled_var = None
_ui_root         = None
_ui_thread       = None
_ui_ready        = threading.Event()
_ui_tasks: queue.Queue = queue.Queue()

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


def _sync_overlay_ui_state(enabled: bool) -> None:
    global _settings_overlay_enabled_var

    if _settings_overlay_enabled_var is not None:
        try:
            if bool(_settings_overlay_enabled_var.get()) != bool(enabled):
                _settings_overlay_enabled_var.set(bool(enabled))
        except Exception:
            _settings_overlay_enabled_var = None

    if _tray_icon:
        try:
            _tray_icon.update_menu()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Shared Tk UI thread
# ---------------------------------------------------------------------------

def _run_ui_loop() -> None:
    global _ui_root

    def _drain_ui_tasks() -> None:
        try:
            while True:
                callback = _ui_tasks.get_nowait()
                callback()
        except queue.Empty:
            pass

        if _ui_root is not None:
            _ui_root.after(20, _drain_ui_tasks)

    root = tk.Tk()
    root.withdraw()
    _ui_root = root
    system_overlay.set_ui_root(root)
    screenshot_mod.set_ui_root(root)
    root.after(20, _drain_ui_tasks)
    _ui_ready.set()
    log.info("Shared Tk UI thread started")

    try:
        root.mainloop()
    finally:
        system_overlay.set_ui_root(None)
        screenshot_mod.set_ui_root(None)
        _ui_root = None
        _ui_ready.clear()
        log.info("Shared Tk UI thread stopped")


def ensure_ui_thread() -> None:
    global _ui_thread

    if _ui_thread is not None and _ui_thread.is_alive() and _ui_root is not None:
        return

    _ui_ready.clear()
    _ui_thread = threading.Thread(target=_run_ui_loop, daemon=True, name="tk-ui-thread")
    _ui_thread.start()
    if not _ui_ready.wait(timeout=5):
        raise RuntimeError("Timed out starting shared Tk UI thread")


def _schedule_on_ui_thread(callback) -> None:
    ensure_ui_thread()
    if _ui_root is None:
        raise RuntimeError("Shared Tk UI root is not available")

    if threading.current_thread() is _ui_thread:
        callback()
    else:
        _ui_tasks.put(callback)


def _shutdown_ui_thread() -> None:
    def _close_ui():
        global _settings_root

        if _settings_root is not None:
            try:
                if _settings_root.winfo_exists():
                    _settings_root.destroy()
            except Exception:
                pass
            _settings_root = None

        system_overlay.close_overlay()

        if _ui_root is not None:
            try:
                _ui_root.quit()
                _ui_root.destroy()
            except Exception:
                pass

    if _ui_root is not None:
        try:
            _schedule_on_ui_thread(_close_ui)
        except Exception as e:
            log.warning(f"Error shutting down shared UI thread: {e}")


# ---------------------------------------------------------------------------
# Shutdown
# ---------------------------------------------------------------------------

def _shutdown() -> None:
    """Unified cleanup: close overlay, restore GPU, stop hook."""
    log.info("Shutting down...")

    # Stop all components in a safe order with error handling
    try:
        _shutdown_ui_thread()
    except Exception as e:
        log.error(f"Error shutting down UI thread: {e}")

    try:
        monitor_server.stop_monitor_server()
    except Exception as e:
        log.error(f"Error stopping monitor server: {e}")
    
    try:
        fan_control.stop_fan_control()
    except Exception as e:
        log.error(f"Error stopping fan control: {e}")
    
    try:
        fan_control.stop_gpu_fan_control()
    except Exception as e:
        log.error(f"Error stopping GPU fan control: {e}")
    
    try:
        auto_sleep.stop_auto_sleep()
    except Exception as e:
        log.error(f"Error stopping auto sleep: {e}")
    
    try:
        gpu_power.restore_gpu_power_limit()
    except Exception as e:
        log.error(f"Error restoring GPU power limit: {e}")
    
    # Stop keyboard hook last (this may wait for thread to exit)
    try:
        hotkey.stop_keyboard_hook()
    except Exception as e:
        log.error(f"Error stopping keyboard hook: {e}")
    
    log.info("Shutdown complete")


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

def _force_window_focus(root: tk.Misc) -> None:
    """Reliably bring a tkinter window to the foreground on Windows."""
    root.lift()
    root.focus_force()
    try:
        hwnd = root.winfo_id()
        ctypes.windll.user32.ShowWindow(hwnd, 9)   # SW_RESTORE
        ctypes.windll.user32.SetForegroundWindow(hwnd)
    except Exception:
        pass


def _apply_window_icon(root: tk.Misc) -> None:
    """Apply the packaged app icon to Tk windows on Windows."""
    try:
        icon_path = cfg.get_resource_path("icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass


def show_settings_dialog() -> None:
    def _run():
        global _settings_root, _settings_overlay_enabled_var

        if _settings_root is not None:
            try:
                if _settings_root.winfo_exists():
                    _force_window_focus(_settings_root)
                    return
            except Exception:
                _settings_root = None

        _initing = [True]   # guard: suppress side-effect callbacks during widget setup

        root = tk.Toplevel(_ui_root)
        _settings_root = root
        root.title("Settings - Little Helper")
        _apply_window_icon(root)
        root.resizable(True, True)

        width, height = 460, 940
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

        # Scrollable content area
        _canvas = tk.Canvas(root, borderwidth=0, highlightthickness=0)
        _vscroll = tk.Scrollbar(root, orient="vertical", command=_canvas.yview)
        _canvas.configure(yscrollcommand=_vscroll.set)
        _vscroll.pack(side="right", fill="y")
        _canvas.pack(side="left", fill="both", expand=True)
        outer = tk.Frame(_canvas, padx=16, pady=12)
        _outer_id = _canvas.create_window((0, 0), window=outer, anchor="nw")

        def _on_outer_configure(event):
            _canvas.configure(scrollregion=_canvas.bbox("all"))

        def _on_canvas_configure(event):
            _canvas.itemconfig(_outer_id, width=event.width)

        outer.bind("<Configure>", _on_outer_configure)
        _canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            _canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        root.bind("<MouseWheel>", _on_mousewheel)

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

        # ── Section 2: Overlay ────────────────────────────────────────────────
        ov_frame = _section(outer, "System Monitor Overlay")

        ov_enabled  = tk.BooleanVar(master=root, value=system_overlay.overlay_is_open())
        ov_opacity  = tk.DoubleVar( master=root, value=_config["overlay"]["opacity"])
        ov_refresh  = tk.IntVar(    master=root, value=_config["overlay"]["refresh_ms"])
        _settings_overlay_enabled_var = ov_enabled

        row4 = tk.Frame(ov_frame)
        row4.pack(fill="x", pady=3)
        ov_cb = tk.Checkbutton(row4, text="Enable",
                       variable=ov_enabled)
        ov_cb.pack(side="left")

        row5 = tk.Frame(ov_frame)
        row5.pack(fill="x", pady=3)
        tk.Label(row5, text="Opacity:", width=14, anchor="w").pack(side="left")
        ov_opacity_spin = tk.Spinbox(row5, from_=0.20, to=1.00, increment=0.05, format="%.2f",
                   textvariable=ov_opacity, width=6)
        ov_opacity_spin.pack(side="left", padx=4)

        row6 = tk.Frame(ov_frame)
        row6.pack(fill="x", pady=3)
        tk.Label(row6, text="Refresh (ms):", width=14, anchor="w").pack(side="left")
        ov_refresh_spin = tk.Spinbox(row6, from_=100, to=5000, increment=100,
                   textvariable=ov_refresh, width=6)
        ov_refresh_spin.pack(side="left", padx=4)

        def _toggle_overlay(*_):
            state = "normal" if ov_enabled.get() else "disabled"
            ov_opacity_spin.configure(state=state)
            ov_refresh_spin.configure(state=state)
        ov_enabled.trace_add("write", _toggle_overlay)
        _toggle_overlay()

        def _apply_overlay(*_):
            if _initing[0]: return
            _config["overlay"]["enabled"]    = ov_enabled.get()
            _config["overlay"]["opacity"]    = round(ov_opacity.get(), 2)
            _config["overlay"]["refresh_ms"] = ov_refresh.get()
            cfg.save_config(_config)
            system_overlay.set_overlay_enabled(
                _config,
                cfg.save_config,
                ov_enabled.get(),
                on_state_change_fn=_sync_overlay_ui_state,
                persist_config=False,
            )
            system_overlay.apply_overlay_opacity(_config["overlay"]["opacity"])

        ov_enabled.trace_add("write", _apply_overlay)
        ov_opacity.trace_add("write", _apply_overlay)
        ov_refresh.trace_add("write", _apply_overlay)

        # ── Section 3: GPU Power Limit ────────────────────────────────────────
        gpu_frame = _section(outer, "GPU Power Limit  (Nvidia only)")

        gpu_enabled = tk.BooleanVar(master=root, value=_config["gpu_power_limit"]["enabled"])
        gpu_watts   = tk.IntVar(   master=root, value=_config["gpu_power_limit"]["watts"])

        # Query GPU limits for Spinbox range
        limits = gpu_power.get_gpu_power_limits()
        w_min, w_max = (50, 500) if limits is None else (int(limits[0]), int(limits[1]))

        row2 = tk.Frame(gpu_frame)
        row2.pack(fill="x", pady=3)
        gpu_cb = tk.Checkbutton(row2, text="Enable", variable=gpu_enabled)
        gpu_cb.pack(side="left")

        row3 = tk.Frame(gpu_frame)
        row3.pack(fill="x", pady=3)
        tk.Label(row3, text="Target watts:", width=14, anchor="w").pack(side="left")
        gpu_spin = tk.Spinbox(row3, from_=w_min, to=w_max, increment=5,
                              textvariable=gpu_watts, width=6)
        gpu_spin.pack(side="left", padx=4)
        gpu_watts_label = tk.Label(row3, text=f"(GPU range: {w_min}–{w_max} W)",
                 font=("Arial", 8), fg="gray")
        gpu_watts_label.pack(side="left")

        def _toggle_gpu_power(*_):
            state = "normal" if gpu_enabled.get() else "disabled"
            gpu_spin.configure(state=state)
            gpu_watts_label.configure(fg="#808080" if state == "disabled" else "gray")
        gpu_enabled.trace_add("write", _toggle_gpu_power)
        _toggle_gpu_power()

        def _apply_gpu_power(*_):
            if _initing[0]: return
            _config["gpu_power_limit"]["enabled"] = gpu_enabled.get()
            _config["gpu_power_limit"]["watts"]   = gpu_watts.get()
            cfg.save_config(_config)
            gpu_power.apply_gpu_power_limit(_config, notify_fn=_notify)

        gpu_enabled.trace_add("write", _apply_gpu_power)
        gpu_watts.trace_add("write", _apply_gpu_power)

        # ── Section 4: CPU Fan Control ────────────────────────────────────────
        fc_frame = _section(outer, "CPU Fan Control")

        fc_cfg      = _config.get("fan_control", {})
        fc_enabled  = tk.BooleanVar(master=root, value=fc_cfg.get("enabled", False))
        fc_source   = tk.StringVar(master=root, value=fc_cfg.get("source", "gpu_temp"))
        fc_interval = tk.IntVar(   master=root, value=fc_cfg.get("interval_s", 3))
        fc_manual   = tk.DoubleVar(master=root, value=fc_cfg.get("manual_pct", 50))

        row_fc0 = tk.Frame(fc_frame)
        row_fc0.pack(fill="x", pady=3)
        fc_cb = tk.Checkbutton(row_fc0, text="Enable", variable=fc_enabled)
        fc_cb.pack(side="left")

        row_fc1 = tk.Frame(fc_frame)
        row_fc1.pack(fill="x", pady=3)
        tk.Label(row_fc1, text="Source:", width=14, anchor="w").pack(side="left")
        fc_source_menu = tk.OptionMenu(row_fc1, fc_source, "cpu_temp", "gpu_temp", "mixed", "manual")
        fc_source_menu.pack(side="left", padx=4)

        row_fc2 = tk.Frame(fc_frame)
        row_fc2.pack(fill="x", pady=3)
        tk.Label(row_fc2, text="Interval (s):", width=14, anchor="w").pack(side="left")
        fc_interval_spin = tk.Spinbox(row_fc2, from_=1, to=60, increment=1,
                   textvariable=fc_interval, width=5)
        fc_interval_spin.pack(side="left", padx=4)

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

        # CPU fan curve editor
        curve_hdr_row = tk.Frame(fc_frame)
        curve_hdr_row.pack(fill="x", pady=(8, 0))
        tk.Label(curve_hdr_row, text="Fan curve:", width=14, anchor="w").pack(side="left")
        tk.Label(curve_hdr_row, text="Temp °C", width=8, anchor="center",
                 font=("Arial", 8, "bold")).pack(side="left")
        tk.Label(curve_hdr_row, text="Fan %", width=8, anchor="center",
                 font=("Arial", 8, "bold")).pack(side="left")

        _fc_curve_vars = []
        _fc_curve_entries = []
        for _ct, _cp in _config.get("fan_control", {}).get("curve", []):
            _ctv = tk.StringVar(master=root, value=str(_ct))
            _cpv = tk.StringVar(master=root, value=str(_cp))
            _fc_curve_vars.append((_ctv, _cpv))
            _cr = tk.Frame(fc_frame)
            _cr.pack(fill="x", pady=1)
            tk.Label(_cr, width=14, anchor="w").pack(side="left")
            _e1 = tk.Entry(_cr, textvariable=_ctv, width=7); _e1.pack(side="left", padx=2)
            _e2 = tk.Entry(_cr, textvariable=_cpv, width=7); _e2.pack(side="left", padx=2)
            _fc_curve_entries.extend([_e1, _e2])

        def _apply_fc_curve():
            try:
                new_curve = sorted([
                    [int(tv.get()), max(0, min(100, int(pv.get())))]
                    for tv, pv in _fc_curve_vars
                ])
                if not new_curve:
                    return
                _config["fan_control"]["curve"] = new_curve
                cfg.save_config(_config)
                if fan_control.fan_control_is_active():
                    fan_control.stop_fan_control()
                    lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
                    if lhm_computer is not None:
                        fan_control.start_fan_control(_config, lhm_computer, lhm_lock)
            except (ValueError, TypeError) as e:
                log.debug(f"Invalid CPU fan curve: {e}")

        fc_apply_curve_btn = tk.Button(fc_frame, text="Apply Curve",
                  command=_apply_fc_curve)
        fc_apply_curve_btn.pack(anchor="w", pady=(2, 6))

        def _toggle_fc_enabled(*_):
            state = "normal" if fc_enabled.get() else "disabled"
            fc_source_menu.configure(state=state)
            fc_interval_spin.configure(state=state)
            manual_slider.configure(state=state)
            fc_apply_curve_btn.configure(state=state)
            for _e in _fc_curve_entries:
                _e.configure(state=state)

        fc_enabled.trace_add("write", _toggle_fc_enabled)
        _toggle_fc_enabled()

        # ── Section 5: GPU Fan Control ────────────────────────────────────────
        gfc_frame = _section(outer, "GPU Fan Control  (Nvidia only)")

        gfc_cfg      = _config.get("gpu_fan_control", {})
        gfc_enabled  = tk.BooleanVar(master=root, value=gfc_cfg.get("enabled", False))
        gfc_source   = tk.StringVar( master=root, value=gfc_cfg.get("source", "gpu_temp"))
        gfc_interval = tk.IntVar(    master=root, value=gfc_cfg.get("interval_s", 2))
        gfc_manual   = tk.DoubleVar( master=root, value=gfc_cfg.get("manual_pct", 50))

        row_gfc0 = tk.Frame(gfc_frame)
        row_gfc0.pack(fill="x", pady=3)
        gfc_cb = tk.Checkbutton(row_gfc0, text="Enable", variable=gfc_enabled)
        gfc_cb.pack(side="left")

        row_gfc1 = tk.Frame(gfc_frame)
        row_gfc1.pack(fill="x", pady=3)
        tk.Label(row_gfc1, text="Source:", width=14, anchor="w").pack(side="left")
        gfc_source_menu = tk.OptionMenu(row_gfc1, gfc_source, "gpu_temp", "manual")
        gfc_source_menu.pack(side="left", padx=4)

        row_gfc2 = tk.Frame(gfc_frame)
        row_gfc2.pack(fill="x", pady=3)
        tk.Label(row_gfc2, text="Interval (s):", width=14, anchor="w").pack(side="left")
        gfc_interval_spin = tk.Spinbox(row_gfc2, from_=1, to=60, increment=1,
                   textvariable=gfc_interval, width=5)
        gfc_interval_spin.pack(side="left", padx=4)

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

        # GPU fan curve editor
        gcurve_hdr_row = tk.Frame(gfc_frame)
        gcurve_hdr_row.pack(fill="x", pady=(8, 0))
        tk.Label(gcurve_hdr_row, text="Fan curve:", width=14, anchor="w").pack(side="left")
        tk.Label(gcurve_hdr_row, text="Temp °C", width=8, anchor="center",
                 font=("Arial", 8, "bold")).pack(side="left")
        tk.Label(gcurve_hdr_row, text="Fan %", width=8, anchor="center",
                 font=("Arial", 8, "bold")).pack(side="left")

        _gfc_curve_vars = []
        _gfc_curve_entries = []
        for _gt, _gp in _config.get("gpu_fan_control", {}).get("curve", []):
            _gtv = tk.StringVar(master=root, value=str(_gt))
            _gpv = tk.StringVar(master=root, value=str(_gp))
            _gfc_curve_vars.append((_gtv, _gpv))
            _grr = tk.Frame(gfc_frame)
            _grr.pack(fill="x", pady=1)
            tk.Label(_grr, width=14, anchor="w").pack(side="left")
            _ge1 = tk.Entry(_grr, textvariable=_gtv, width=7); _ge1.pack(side="left", padx=2)
            _ge2 = tk.Entry(_grr, textvariable=_gpv, width=7); _ge2.pack(side="left", padx=2)
            _gfc_curve_entries.extend([_ge1, _ge2])

        def _apply_gfc_curve():
            try:
                new_curve = sorted([
                    [int(tv.get()), max(0, min(100, int(pv.get())))]
                    for tv, pv in _gfc_curve_vars
                ])
                if not new_curve:
                    return
                _config["gpu_fan_control"]["curve"] = new_curve
                cfg.save_config(_config)
                if fan_control.gpu_fan_control_is_active():
                    fan_control.stop_gpu_fan_control()
                    fan_control.start_gpu_fan_control(_config)
            except (ValueError, TypeError) as e:
                log.debug(f"Invalid GPU fan curve: {e}")

        gfc_apply_curve_btn = tk.Button(gfc_frame, text="Apply Curve",
                  command=_apply_gfc_curve)
        gfc_apply_curve_btn.pack(anchor="w", pady=(2, 6))

        def _toggle_gfc_enabled(*_):
            state = "normal" if gfc_enabled.get() else "disabled"
            gfc_source_menu.configure(state=state)
            gfc_interval_spin.configure(state=state)
            gfc_manual_slider.configure(state=state)
            gfc_apply_curve_btn.configure(state=state)
            for _e in _gfc_curve_entries:
                _e.configure(state=state)

        gfc_enabled.trace_add("write", _toggle_gfc_enabled)
        _toggle_gfc_enabled()

        # ── Section 6: Auto Sleep ──────────────────────────────────────────────
        as_frame = _section(outer, "Auto Sleep")

        as_cfg           = _config.get("auto_sleep", {})
        as_enabled       = tk.BooleanVar(master=root, value=as_cfg.get("enabled", False))
        as_idle_seconds  = tk.IntVar(master=root, value=as_cfg.get("idle_seconds", 60))
        as_cpu_threshold = tk.IntVar(master=root, value=as_cfg.get("cpu_threshold", 5))
        as_gpu_threshold = tk.IntVar(master=root, value=as_cfg.get("gpu_threshold", 5))
        as_disk_threshold= tk.DoubleVar(master=root, value=as_cfg.get("disk_threshold_mbps", 1.0))
        as_countdown     = tk.IntVar(master=root, value=as_cfg.get("countdown_seconds", 10))

        row_as0 = tk.Frame(as_frame)
        row_as0.pack(fill="x", pady=3)
        as_cb = tk.Checkbutton(row_as0, text="Enable", variable=as_enabled)
        as_cb.pack(side="left")

        row_as1 = tk.Frame(as_frame)
        row_as1.pack(fill="x", pady=3)
        tk.Label(row_as1, text="Idle time (s):", width=16, anchor="w").pack(side="left")
        as_idle_spin = tk.Spinbox(row_as1, from_=1, to=3600, increment=1,
                   textvariable=as_idle_seconds, width=5)
        as_idle_spin.pack(side="left", padx=4)

        row_as2 = tk.Frame(as_frame)
        row_as2.pack(fill="x", pady=3)
        tk.Label(row_as2, text="CPU threshold (%):", width=16, anchor="w").pack(side="left")
        as_cpu_spin = tk.Spinbox(row_as2, from_=1, to=100, increment=1,
                   textvariable=as_cpu_threshold, width=5)
        as_cpu_spin.pack(side="left", padx=4)

        row_as3 = tk.Frame(as_frame)
        row_as3.pack(fill="x", pady=3)
        tk.Label(row_as3, text="GPU threshold (%):", width=16, anchor="w").pack(side="left")
        as_gpu_spin = tk.Spinbox(row_as3, from_=1, to=100, increment=1,
                   textvariable=as_gpu_threshold, width=5)
        as_gpu_spin.pack(side="left", padx=4)

        row_as4 = tk.Frame(as_frame)
        row_as4.pack(fill="x", pady=3)
        tk.Label(row_as4, text="Disk threshold (MB/s):", width=16, anchor="w").pack(side="left")
        as_disk_spin = tk.Spinbox(row_as4, from_=0.1, to=100, increment=0.1,
                   textvariable=as_disk_threshold, width=5)
        as_disk_spin.pack(side="left", padx=4)

        row_as5 = tk.Frame(as_frame)
        row_as5.pack(fill="x", pady=3)
        tk.Label(row_as5, text="Countdown (s):", width=16, anchor="w").pack(side="left")
        as_countdown_spin = tk.Spinbox(row_as5, from_=5, to=60, increment=1,
                   textvariable=as_countdown, width=5)
        as_countdown_spin.pack(side="left", padx=4)

        def _apply_as_enabled(*_):
            if _initing[0]: return
            _config["auto_sleep"]["enabled"] = as_enabled.get()
            cfg.save_config(_config)
            if as_enabled.get():
                if not auto_sleep.is_auto_sleep_active():
                    auto_sleep.start_auto_sleep(_config)
            else:
                auto_sleep.stop_auto_sleep()

        as_enabled.trace_add("write", _apply_as_enabled)

        def _apply_as_settings(*_):
            """Apply Auto Sleep settings changes and restart monitoring."""
            if _initing[0]: return
            _config["auto_sleep"]["idle_seconds"] = as_idle_seconds.get()
            _config["auto_sleep"]["cpu_threshold"] = as_cpu_threshold.get()
            _config["auto_sleep"]["gpu_threshold"] = as_gpu_threshold.get()
            _config["auto_sleep"]["disk_threshold_mbps"] = as_disk_threshold.get()
            _config["auto_sleep"]["countdown_seconds"] = as_countdown.get()
            cfg.save_config(_config)
            if auto_sleep.is_auto_sleep_active():
                auto_sleep.stop_auto_sleep()
                auto_sleep.start_auto_sleep(_config)
            _notify("Auto Sleep settings applied")

        # Auto Sleep buttons row
        row_as_buttons = tk.Frame(as_frame)
        row_as_buttons.pack(fill="x", pady=(8, 0))

        as_apply_btn = tk.Button(
            row_as_buttons,
            text="Apply Changes",
            command=_apply_as_settings,
        )
        as_apply_btn.pack(side="left", padx=2)

        def _toggle_as_enabled(*_):
            state = "normal" if as_enabled.get() else "disabled"
            as_idle_spin.configure(state=state)
            as_cpu_spin.configure(state=state)
            as_gpu_spin.configure(state=state)
            as_disk_spin.configure(state=state)
            as_countdown_spin.configure(state=state)
            as_apply_btn.configure(state=state)

        as_enabled.trace_add("write", _toggle_as_enabled)
        _toggle_as_enabled()

        # ── Section 8: Monitor Server ─────────────────────────────────────────
        ms_frame = _section(outer, "Monitor Server")

        ms_cfg = _config.get("monitor_server", {})
        ms_enabled = tk.BooleanVar(master=root, value=ms_cfg.get("enabled", False))
        ms_host = tk.StringVar(master=root, value=ms_cfg.get("host", monitor_server.DEFAULT_HOST))
        ms_port = tk.IntVar(master=root, value=ms_cfg.get("port", monitor_server.DEFAULT_PORT))
        ms_token = tk.StringVar(master=root, value=ms_cfg.get("token", ""))
        ms_mdns = tk.BooleanVar(master=root, value=ms_cfg.get("mdns", True))
        ms_hint = tk.StringVar(master=root)

        row_ms0 = tk.Frame(ms_frame)
        row_ms0.pack(fill="x", pady=3)
        tk.Checkbutton(row_ms0, text="Enable", variable=ms_enabled).pack(side="left")
        tk.Checkbutton(row_ms0, text="mDNS Broadcast", variable=ms_mdns).pack(side="left", padx=(16, 0))

        row_ms1 = tk.Frame(ms_frame)
        row_ms1.pack(fill="x", pady=3)
        tk.Label(row_ms1, text="Bind IP:", width=14, anchor="w").pack(side="left")
        ms_host_entry = tk.Entry(row_ms1, textvariable=ms_host, width=18)
        ms_host_entry.pack(side="left", padx=4)

        row_ms2 = tk.Frame(ms_frame)
        row_ms2.pack(fill="x", pady=3)
        tk.Label(row_ms2, text="Port:", width=14, anchor="w").pack(side="left")
        ms_port_spin = tk.Spinbox(row_ms2, from_=1, to=65535, increment=1,
                                  textvariable=ms_port, width=8)
        ms_port_spin.pack(side="left", padx=4)

        row_ms3 = tk.Frame(ms_frame)
        row_ms3.pack(fill="x", pady=3)
        tk.Label(row_ms3, text="Token:", width=14, anchor="w").pack(side="left")
        ms_token_entry = tk.Entry(row_ms3, textvariable=ms_token, width=24)
        ms_token_entry.pack(side="left", padx=4)

        tk.Label(
            ms_frame,
            text="Leave token blank to disable authentication. mDNS broadcasts service on local network.",
            font=("Arial", 8),
            fg="gray",
            anchor="w",
            justify="left",
        ).pack(fill="x", pady=(0, 2))

        ms_hint_label = tk.Label(
            ms_frame,
            textvariable=ms_hint,
            font=("Arial", 8),
            fg="gray",
            anchor="w",
            justify="left",
            wraplength=390,
        )
        ms_hint_label.pack(fill="x", pady=(0, 4))

        def _collect_monitor_server_config() -> dict:
            try:
                port_value = ms_port.get()
            except tk.TclError:
                port_value = monitor_server.DEFAULT_PORT
            return monitor_server.normalize_monitor_server_config(
                {
                    "monitor_server": {
                        "enabled": ms_enabled.get(),
                        "host": ms_host.get(),
                        "port": port_value,
                        "token": ms_token.get(),
                        "mdns": ms_mdns.get(),
                    }
                }
            )

        def _update_monitor_server_hint(*_):
            preview_cfg = _collect_monitor_server_config()
            urls = monitor_server.get_monitor_urls(preview_cfg)
            auth_mode = "token required" if preview_cfg.get("token") else "no auth"
            state = "enabled" if preview_cfg.get("enabled") else "disabled"
            mdns_state = "on" if preview_cfg.get("mdns") else "off"
            ms_hint.set(
                f"HTTP: {urls['http']}\nWS: {urls['websocket']}\nState: {state}, {auth_mode}, mDNS: {mdns_state}"
            )

        def _apply_monitor_server_settings(notify_user: bool = True):
            if _initing[0]:
                return
            new_cfg = _collect_monitor_server_config()
            old_cfg = dict(_config.get("monitor_server", {}))
            _config["monitor_server"] = new_cfg
            cfg.save_config(_config)

            try:
                if new_cfg.get("enabled"):
                    if monitor_server.monitor_server_is_running() and old_cfg != new_cfg:
                        monitor_server.restart_monitor_server(_config)
                    elif not monitor_server.monitor_server_is_running():
                        monitor_server.start_monitor_server(_config)
                else:
                    monitor_server.stop_monitor_server()
                if notify_user:
                    urls = monitor_server.get_monitor_urls(new_cfg)
                    if new_cfg.get("enabled"):
                        _notify(
                            f"Monitor server active\nHTTP: {urls['http']}\nWS: {urls['websocket']}",
                            "Monitor Server",
                        )
                    else:
                        _notify("Monitor server stopped", "Monitor Server")
            except Exception as e:
                log.error(f"Error applying monitor server settings: {e}", exc_info=True)
                if notify_user:
                    _notify(f"Monitor server error: {e}", "Monitor Server")

        for _var in (ms_enabled, ms_host, ms_port, ms_token, ms_mdns):
            _var.trace_add("write", _update_monitor_server_hint)
        _update_monitor_server_hint()

        ms_buttons_row = tk.Frame(ms_frame)
        ms_buttons_row.pack(fill="x", pady=(4, 0))

        ms_apply_btn = tk.Button(
            ms_buttons_row,
            text="Apply Changes",
            command=lambda: _apply_monitor_server_settings(notify_user=True),
        )
        ms_apply_btn.pack(side="left", padx=2)

        ms_enabled.trace_add(
            "write",
            lambda *_: None if _initing[0] else _apply_monitor_server_settings(notify_user=True),
        )

        # ── Close: save key fields and slider defaults ────────────────────────
        def _close():
            global _settings_root, _settings_overlay_enabled_var
            dirty = False

            new_monitor_cfg = _collect_monitor_server_config()
            old_monitor_cfg = dict(_config.get("monitor_server", {}))
            if old_monitor_cfg != new_monitor_cfg:
                _config["monitor_server"] = new_monitor_cfg
                dirty = True

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
            if old_monitor_cfg != new_monitor_cfg:
                try:
                    if new_monitor_cfg.get("enabled"):
                        monitor_server.restart_monitor_server(_config)
                    else:
                        monitor_server.stop_monitor_server()
                except Exception as e:
                    log.error(f"Error restarting monitor server on settings close: {e}", exc_info=True)
            if _poll_id[0]:
                root.after_cancel(_poll_id[0])
            if _gfc_poll_id[0]:
                root.after_cancel(_gfc_poll_id[0])
            _settings_overlay_enabled_var = None
            _settings_root = None
            root.destroy()

        root.protocol("WM_DELETE_WINDOW", _close)

        # ── Bottom buttons ────────────────────────────────────────────────────
        btn_frame = tk.Frame(outer, pady=6)
        btn_frame.pack(fill="x")

        # All widgets built — allow trace callbacks to fire from now on
        _initing[0] = False

    _schedule_on_ui_thread(_run)


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
        system_overlay.toggle_overlay(
            _config,
            cfg.save_config,
            on_state_change_fn=_sync_overlay_ui_state,
        )

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
        try:
            gpu_power.apply_gpu_power_limit(_config, notify_fn=_notify)
        except Exception as e:
            log.error(f"Error applying GPU power limit: {e}", exc_info=True)
        
        # Sync manual_pct from config into the module
        try:
            fan_control.set_manual_pct(_config.get("fan_control", {}).get("manual_pct", 50))
            fan_control.set_gpu_manual_pct(_config.get("gpu_fan_control", {}).get("manual_pct", 50))
        except Exception as e:
            log.error(f"Error setting manual fan percentages: {e}", exc_info=True)
        
        # Register auto_sleep keyboard activity callback
        try:
            hotkey.register_activity_callback(auto_sleep.notify_keyboard_activity)
        except Exception as e:
            log.error(f"Error registering activity callback: {e}", exc_info=True)
        
        # Auto-start auto_sleep if configured
        try:
            auto_sleep.start_auto_sleep(_config)
        except Exception as e:
            log.error(f"Error starting auto_sleep: {e}", exc_info=True)
        
        # Auto-start CPU fan control if configured
        try:
            if _config.get("fan_control", {}).get("enabled", False):
                lhm_computer, lhm_lock = system_overlay.get_lhm_computer()
                if lhm_computer is not None:
                    fan_control.start_fan_control(_config, lhm_computer, lhm_lock)
                else:
                    log.warning("CPU fan control enabled in config but LHM is not available")
        except Exception as e:
            log.error(f"Error starting CPU fan control: {e}", exc_info=True)
        
        # Auto-start GPU fan control if configured
        try:
            if _config.get("gpu_fan_control", {}).get("enabled", False):
                fan_control.start_gpu_fan_control(_config)
        except Exception as e:
            log.error(f"Error starting GPU fan control: {e}", exc_info=True)
        
        # Auto-open overlay if configured
        try:
            if _config.get("overlay", {}).get("enabled", False):
                system_overlay.set_overlay_enabled(
                    _config,
                    cfg.save_config,
                    True,
                    on_state_change_fn=_sync_overlay_ui_state,
                )
        except Exception as e:
            log.error(f"Error starting overlay: {e}", exc_info=True)

        try:
            if _config.get("monitor_server", {}).get("enabled", False):
                monitor_server.start_monitor_server(_config)
        except Exception as e:
            log.error(f"Error starting monitor server: {e}", exc_info=True)
        
        # Refresh tray menu checkmarks after all state has been set
        try:
            icon.update_menu()
        except Exception as e:
            log.error(f"Error updating tray menu: {e}", exc_info=True)

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
    ensure_ui_thread()
    
    # Set up auto_sleep UI callback
    auto_sleep.set_ui_callback(_schedule_on_ui_thread)

    try:
        _hidden_hwnd = create_hidden_window()
        create_tray_icon()
    except Exception as e:
        log.error(f"Fatal error: {e}", exc_info=True)
