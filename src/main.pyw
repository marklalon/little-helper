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

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_PATH = cfg.get_log_path()
if os.path.exists(LOG_PATH):
    open(LOG_PATH, "w", encoding="utf-8").close()

logging.getLogger().setLevel(logging.WARNING)

log = logging.getLogger("little_helper")
log.setLevel(logging.DEBUG)
log.propagate = False

_fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
_fh.setLevel(logging.DEBUG)
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s %(message)s"))
log.addHandler(_fh)

if sys.stderr and hasattr(sys.stderr, "write"):
    _sh = logging.StreamHandler(sys.stderr)
    _sh.setLevel(logging.DEBUG)
    _sh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s %(message)s"))
    log.addHandler(_sh)

# Propagate to sub-module loggers
for _mod in ("little_helper.config", "little_helper.clipboard_paste",
             "little_helper.screenshot", "little_helper.hotkey",
             "little_helper.gpu_power", "little_helper.system_overlay"):
    _ml = logging.getLogger(_mod)
    _ml.setLevel(logging.DEBUG)
    _ml.propagate = False
    _ml.addHandler(_fh)
    if sys.stderr and hasattr(sys.stderr, "write"):
        _ml.addHandler(_sh)

# ---------------------------------------------------------------------------
# Global state
# ---------------------------------------------------------------------------

_config: dict    = {}
_tray_icon       = None

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
    system_overlay.close_overlay()
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

def show_settings_dialog() -> None:
    global _config

    root = tk.Tk()
    root.title("Settings - Little Helper")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    width, height = 420, 480
    root.update_idletasks()
    x = (root.winfo_screenwidth()  // 2) - (width  // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    root.lift()
    root.after(100, lambda: (root.lift(), root.focus_force()))

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

    row0 = tk.Frame(hk_frame)
    row0.pack(fill="x", pady=3)
    tk.Label(row0, text="Paste:", width=14, anchor="w").pack(side="left")
    tk.OptionMenu(row0, paste_mod, "Ctrl", "Alt").pack(side="left", padx=4)
    tk.Entry(row0, textvariable=paste_key, width=5).pack(side="left")

    ss_mod = tk.StringVar(master=root, value=_config["screenshot_hotkey"]["modifier"].capitalize())
    ss_key = tk.StringVar(master=root, value=_config["screenshot_hotkey"]["key"].upper())

    row1 = tk.Frame(hk_frame)
    row1.pack(fill="x", pady=3)
    tk.Label(row1, text="Screenshot:", width=14, anchor="w").pack(side="left")
    tk.OptionMenu(row1, ss_mod, "Ctrl", "Alt").pack(side="left", padx=4)
    tk.Entry(row1, textvariable=ss_key, width=5).pack(side="left")

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

    # ── Buttons ───────────────────────────────────────────────────────────
    def on_save():
        global _config
        _config["paste_hotkey"]["modifier"]      = paste_mod.get().lower()
        _config["paste_hotkey"]["key"]            = paste_key.get().upper()
        _config["screenshot_hotkey"]["modifier"]  = ss_mod.get().lower()
        _config["screenshot_hotkey"]["key"]       = ss_key.get().upper()
        _config["gpu_power_limit"]["enabled"]     = gpu_enabled.get()
        _config["gpu_power_limit"]["watts"]       = gpu_watts.get()
        _config["overlay"]["enabled"]             = ov_enabled.get()
        _config["overlay"]["opacity"]             = round(ov_opacity.get(), 2)
        _config["overlay"]["refresh_ms"]          = ov_refresh.get()

        cfg.save_config(_config)
        # Re-apply GPU limit with new settings
        gpu_power.apply_gpu_power_limit(_config, notify_fn=_notify)
        root.destroy()
        _notify("Settings saved.", "Settings")

    btn_frame = tk.Frame(outer)
    btn_frame.pack(pady=18)
    tk.Button(btn_frame, text="Save",   width=10, command=on_save ).pack(side="left", padx=8)
    tk.Button(btn_frame, text="Cancel", width=10, command=root.destroy).pack(side="left", padx=8)

    root.mainloop()


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
            daemon=True,
        ).start()

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
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Settings", _on_settings),
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
        # Auto-open overlay if configured
        if _config.get("overlay", {}).get("enabled", False):
            threading.Thread(
                target=system_overlay.toggle_overlay,
                args=(_config, cfg.save_config),
                daemon=True,
            ).start()

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
