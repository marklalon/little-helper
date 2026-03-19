"""
Little Helper - System monitoring overlay window.

Shows RAM, CPU, GPU stats in a draggable, resizable, semi-transparent overlay.
Runs in its own daemon thread with a Tkinter mainloop.
"""

import os
import queue
import threading
import logging
import tkinter as tk
from tkinter import font as tkfont

log = logging.getLogger("little_helper.system_overlay")

# --- NVML state (initialised once at startup) ---
_nvml_available = False
_nvml_handle    = None

# --- LibreHardwareMonitor state ---
_lhm_available = False
_lhm_computer  = None
_lhm_cpu_temp  = None  # ISensor reference
_lhm_cpu_power = None  # ISensor reference


def init_nvml() -> bool:
    """Attempt to initialise pynvml for GPU index 0. Call once at startup."""
    global _nvml_available, _nvml_handle
    # Prime psutil cpu_percent so the first background fetch returns a real value
    # (first call with interval=None always returns 0.0 unless primed)
    try:
        import psutil
        psutil.cpu_percent(interval=None)
    except Exception:
        pass
    try:
        import pynvml
        pynvml.nvmlInit()
        _nvml_handle    = pynvml.nvmlDeviceGetHandleByIndex(0)
        _nvml_available = True
        name = pynvml.nvmlDeviceGetName(_nvml_handle)
        log.info(f"NVML initialised: {name}")
        return True
    except Exception as e:
        log.warning(f"NVML init failed (no Nvidia GPU?): {e}")
        _nvml_available = False
        return False


def init_lhm() -> bool:
    """Attempt to initialise LibreHardwareMonitorLib for CPU sensors. Call once at startup."""
    global _lhm_available, _lhm_computer, _lhm_cpu_temp, _lhm_cpu_power
    try:
        import clr
        # Find the DLL path
        dll_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "lib", "lhm")
        if not os.path.exists(dll_dir):
            log.debug(f"LibreHardwareMonitor DLLs not found at {dll_dir}")
            return False
        
        # Add reference to the DLL
        clr.AddReference(os.path.join(dll_dir, "LibreHardwareMonitorLib.dll"))
        from LibreHardwareMonitor.Hardware import Computer
        
        _lhm_computer = Computer()
        _lhm_computer.IsCpuEnabled = True
        _lhm_computer.IsGpuEnabled = False
        _lhm_computer.IsMemoryEnabled = False
        _lhm_computer.IsMotherboardEnabled = False
        _lhm_computer.IsControllerEnabled = False
        _lhm_computer.IsNetworkEnabled = False
        _lhm_computer.IsStorageEnabled = False
        _lhm_computer.Open()
        
        # Find CPU temperature and power sensors
        for hardware in _lhm_computer.Hardware:
            if hardware.HardwareType.ToString() == "Cpu":
                hardware.Update()
                for sensor in hardware.Sensors:
                    sensor_type = sensor.SensorType.ToString()
                    name = sensor.Name.lower()
                    if sensor_type == "Temperature" and _lhm_cpu_temp is None:
                        if "core" in name or "package" in name or "cpu" in name:
                            _lhm_cpu_temp = sensor
                            log.debug(f"Found CPU temp sensor: {sensor.Name}")
                    elif sensor_type == "Power" and _lhm_cpu_power is None:
                        if "package" in name or "cpu" in name:
                            _lhm_cpu_power = sensor
                            log.debug(f"Found CPU power sensor: {sensor.Name}")
        
        _lhm_available = True
        log.info("LibreHardwareMonitorLib initialised for CPU sensors")
        return True
    except Exception as e:
        log.warning(f"LibreHardwareMonitorLib init failed: {e}")
        _lhm_available = False
        return False


# ---------------------------------------------------------------------------
# Data collection
# ---------------------------------------------------------------------------

def get_gpu_stats() -> dict:
    """Return GPU metrics dict; any unavailable metric is None."""
    result = {
        "vram_used_mb":  None,
        "vram_total_mb": None,
        "gpu_util_pct":  None,
        "gpu_temp_c":    None,
        "gpu_power_w":   None,
    }
    if not _nvml_available:
        return result
    try:
        import pynvml
        h = _nvml_handle
        try:
            mem = pynvml.nvmlDeviceGetMemoryInfo(h)
            result["vram_used_mb"]  = mem.used  / 1024**2
            result["vram_total_mb"] = mem.total / 1024**2
        except Exception:
            pass
        try:
            result["gpu_temp_c"] = pynvml.nvmlDeviceGetTemperature(
                h, pynvml.NVML_TEMPERATURE_GPU
            )
        except Exception:
            pass
        try:
            result["gpu_power_w"] = pynvml.nvmlDeviceGetPowerUsage(h) / 1000.0
        except Exception:
            pass
        try:
            result["gpu_util_pct"] = pynvml.nvmlDeviceGetUtilizationRates(h).gpu
        except Exception:
            pass
    except Exception as e:
        log.debug(f"get_gpu_stats error: {e}")
    return result


def get_system_stats() -> dict:
    """Return system metrics dict."""
    result = {
        "ram_used_gb":  None,
        "ram_total_gb": None,
        "ram_pct":      None,
        "cpu_pct":      None,
        "cpu_temp_c":   None,
        "cpu_power_w":  None,
    }
    try:
        import psutil
        vm = psutil.virtual_memory()
        result["ram_used_gb"]  = vm.used  / 1024**3
        result["ram_total_gb"] = vm.total / 1024**3
        result["ram_pct"]      = vm.percent
        # Use 0.1s interval for accurate measurement (blocks fetch thread briefly)
        result["cpu_pct"]      = psutil.cpu_percent(interval=0.1)

        # CPU temperature/power via LibreHardwareMonitorLib
        if _lhm_available and _lhm_computer is not None:
            try:
                # Update CPU hardware
                for hardware in _lhm_computer.Hardware:
                    if hardware.HardwareType.ToString() == "Cpu":
                        hardware.Update()
                        break
                
                if _lhm_cpu_temp is not None:
                    result["cpu_temp_c"] = _lhm_cpu_temp.Value
                if _lhm_cpu_power is not None:
                    result["cpu_power_w"] = _lhm_cpu_power.Value
            except Exception as e:
                log.debug(f"LHM sensor read error: {e}")

    except Exception as e:
        log.error(f"get_system_stats error: {e}", exc_info=True)

    return result


# ---------------------------------------------------------------------------
# Overlay window
# ---------------------------------------------------------------------------

_BG        = "#1a1a1a"
_TITLE_BG  = "#252525"
_FG_NORMAL = "#00e676"
_FG_WARM   = "#ffdd00"
_FG_HOT    = "#ff4444"
_FG_NA     = "#555555"
_FONT      = ("Consolas", 9)
_FONT_BOLD = ("Consolas", 9, "bold")


def _temp_color(temp_c):
    if temp_c is None:
        return _FG_NA
    if temp_c >= 80:
        return _FG_HOT
    if temp_c >= 70:
        return _FG_WARM
    return _FG_NORMAL


def _fmt(val, fmt, unit="", na="N/A"):
    if val is None:
        return na
    return f"{val:{fmt}}{unit}"


class SystemMonitorOverlay:
    """
    Semi-transparent always-on-top overlay.
    Runs inside its own tk.Tk() mainloop on a dedicated daemon thread.
    """

    def __init__(self, config: dict, save_config_fn):
        self.config        = config
        self.save_config   = save_config_fn
        self._running      = False
        self._fetch_running = False
        self._q: queue.Queue = queue.Queue(maxsize=1)

        # drag state
        self._drag_offset_x = 0
        self._drag_offset_y = 0

        self.root   = None
        self._labels = {}  # key -> tk.Label

    # -----------------------------------------------------------------------
    # Public API (called from other threads)
    # -----------------------------------------------------------------------

    def run(self) -> None:
        """Build UI and run mainloop (blocks until close())."""
        self._running = True
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", self.config["overlay"]["opacity"])
        self.root.configure(bg=_BG)
        self.root.resizable(False, False)

        self._build_ui()
        self._position_window()
        self.root.protocol("WM_DELETE_WINDOW", self.close)

        # Kick off stats loop
        self.root.after(100, self._update_stats)

        self.root.mainloop()
        self._running = False

    def close(self) -> None:
        """Destroy the window (safe to call from any thread)."""
        global _overlay_instance
        self._running = False
        if self.root:
            try:
                self.root.after(0, self.root.destroy)
            except Exception:
                pass
        # Clear global instance so tray menu checkbox updates correctly
        _overlay_instance = None

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------

    def _build_ui(self) -> None:
        root = self.root

        # ── Title bar ──────────────────────────────────────────────────────
        self._title_bar = tk.Frame(root, bg=_TITLE_BG, height=22, cursor="fleur")
        self._title_bar.pack(fill="x", side="top")
        self._title_bar.pack_propagate(False)

        tk.Label(
            self._title_bar, text="◈ MONITOR", bg=_TITLE_BG,
            fg=_FG_NA, font=_FONT_BOLD, anchor="w",
        ).pack(side="left", padx=6)

        self._close_btn = tk.Label(
            self._title_bar, text="[×]", bg=_TITLE_BG,
            fg=_FG_NA, font=_FONT, cursor="hand2",
        )
        self._close_btn.pack(side="right", padx=4)
        self._close_btn.bind("<Button-1>", lambda e: self.close())

        # Drag bindings on title bar (skip compact button so it keeps its click handler)
        self._title_bar.bind("<ButtonPress-1>",   self._drag_start)
        self._title_bar.bind("<B1-Motion>",        self._drag_motion)
        self._title_bar.bind("<ButtonRelease-1>",  self._drag_stop)
        for child in self._title_bar.winfo_children():
            if child is self._close_btn:
                continue
            child.bind("<ButtonPress-1>",  self._drag_start)
            child.bind("<B1-Motion>",       self._drag_motion)
            child.bind("<ButtonRelease-1>", self._drag_stop)

        # ── Content frame ─────────────────────────────────────────────────
        self._content = tk.Frame(root, bg=_BG)
        self._content.pack(fill="both", expand=True)

        # System section
        self._sys_frame = tk.Frame(self._content, bg=_BG)
        self._sys_frame.pack(fill="x", padx=6, pady=(4, 2))

        self._make_row(self._sys_frame, "cpu", "CPU")
        self._make_row(self._sys_frame, "ram", "RAM")

        tk.Frame(self._content, bg="#333333", height=1).pack(fill="x", padx=6, pady=2)

        # GPU section
        self._gpu_frame = tk.Frame(self._content, bg=_BG)
        self._gpu_frame.pack(fill="x", padx=6, pady=(2, 4))

        self._make_row(self._gpu_frame, "gpu",  "GPU")
        self._make_row(self._gpu_frame, "vram", "VRAM")


    def _make_row(self, parent, key: str, label: str) -> None:
        row = tk.Frame(parent, bg=_BG)
        row.pack(fill="x", pady=1)
        tk.Label(row, text=f"{label:<4}", bg=_BG, fg=_FG_NA,
                 font=_FONT, width=4, anchor="w").pack(side="left")
        lbl = tk.Label(row, text="...", bg=_BG, fg=_FG_NORMAL,
                       font=_FONT, anchor="w")
        lbl.pack(side="left", fill="x", expand=True)
        self._labels[key] = lbl

    # -----------------------------------------------------------------------
    # Stats update (queue-based, non-blocking UI)
    # -----------------------------------------------------------------------

    def _update_stats(self) -> None:
        if not self._running:
            return

        # Drain queue
        try:
            stats = self._q.get_nowait()
            self._apply_stats(stats)
        except queue.Empty:
            pass

        # Spawn fetch thread if idle
        if not self._fetch_running:
            self._fetch_running = True
            threading.Thread(target=self._fetch_thread, daemon=True).start()

        refresh = self.config["overlay"].get("refresh_ms", 1000)
        self.root.after(refresh, self._update_stats)

    def _fetch_thread(self) -> None:
        try:
            sys_stats = get_system_stats()
            gpu_stats = get_gpu_stats()
            combined  = {**sys_stats, **gpu_stats}
            try:
                self._q.put_nowait(combined)
            except queue.Full:
                pass
        finally:
            self._fetch_running = False

    def _apply_stats(self, s: dict) -> None:
        """Update label text and colours from a stats dict."""
        # CPU: [usage%] [temp°C] [power implicit via temp color]
        cpu_parts = []
        if s.get("cpu_pct") is not None:
            cpu_parts.append(f"{s['cpu_pct']:.0f}%")
        if s.get("cpu_temp_c") is not None:
            cpu_parts.append(f"{s['cpu_temp_c']:.0f}°C")
        if s.get("cpu_power_w") is not None:
            cpu_parts.append(f"{s['cpu_power_w']:.0f}W")
        cpu_color = _temp_color(s["cpu_temp_c"]) if s.get("cpu_temp_c") is not None else _FG_NORMAL
        self._set(self._labels["cpu"], "  ".join(cpu_parts) if cpu_parts else "N/A", cpu_color)

        # RAM: used/total only, no percentage
        if s.get("ram_used_gb") is not None:
            self._set(self._labels["ram"],
                      f"{s['ram_used_gb']:.1f}/{s['ram_total_gb']:.0f}GB", _FG_NORMAL)
        else:
            self._set(self._labels["ram"], "N/A", _FG_NA)

        # GPU: [usage%] [temp°C] [power W]
        gpu_parts = []
        if s.get("gpu_util_pct") is not None:
            gpu_parts.append(f"{s['gpu_util_pct']}%")
        if s.get("gpu_temp_c") is not None:
            gpu_parts.append(f"{s['gpu_temp_c']:.0f}°C")
        if s.get("gpu_power_w") is not None:
            gpu_parts.append(f"{s['gpu_power_w']:.0f}W")
        gpu_color = _temp_color(s["gpu_temp_c"]) if s.get("gpu_temp_c") is not None else _FG_NORMAL
        self._set(self._labels["gpu"], "  ".join(gpu_parts) if gpu_parts else "N/A", gpu_color)

        # VRAM: used/total
        if s.get("vram_used_mb") is not None:
            self._set(self._labels["vram"],
                      f"{s['vram_used_mb']/1024:.1f}/{s['vram_total_mb']/1024:.0f}GB", _FG_NORMAL)
        else:
            self._set(self._labels["vram"], "N/A", _FG_NA)

    @staticmethod
    def _set(label: tk.Label, text: str, color: str) -> None:
        label.configure(text=text, fg=color)

    # -----------------------------------------------------------------------
    # Drag to move
    # -----------------------------------------------------------------------

    def _drag_start(self, event) -> None:
        self._drag_offset_x = event.x_root - self.root.winfo_x()
        self._drag_offset_y = event.y_root - self.root.winfo_y()

    def _drag_motion(self, event) -> None:
        x = event.x_root - self._drag_offset_x
        y = event.y_root - self._drag_offset_y
        self.root.geometry(f"+{x}+{y}")

    def _drag_stop(self, event) -> None:
        self._save_position()

    # -----------------------------------------------------------------------
    # Position helpers
    # -----------------------------------------------------------------------

    def _position_window(self) -> None:
        cfg = self.config["overlay"]
        self.root.update_idletasks()  # allow content to determine natural size
        if cfg["x"] == -1 or cfg["y"] == -1:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            w  = self.root.winfo_reqwidth() or 210
            h  = self.root.winfo_reqheight() or 120
            x  = sw - w - 10
            y  = sh - h - 50  # Account for taskbar
        else:
            x, y = cfg["x"], cfg["y"]
        self.root.geometry(f"+{x}+{y}")

    def _save_position(self) -> None:
        self.config["overlay"]["x"] = self.root.winfo_x()
        self.config["overlay"]["y"] = self.root.winfo_y()
        self.save_config(self.config)




# ---------------------------------------------------------------------------
# Module-level toggle helper (called from tray menu)
# ---------------------------------------------------------------------------

_overlay_instance: SystemMonitorOverlay | None = None


def toggle_overlay(config: dict, save_config_fn) -> None:
    """Show or hide the overlay. Safe to call from any thread."""
    global _overlay_instance
    if _overlay_instance is not None:
        _overlay_instance.close()
        _overlay_instance = None
    else:
        _overlay_instance = SystemMonitorOverlay(config, save_config_fn)
        threading.Thread(target=_overlay_instance.run, daemon=True).start()


def close_overlay() -> None:
    """Close overlay if open (called during shutdown)."""
    global _overlay_instance
    if _overlay_instance is not None:
        _overlay_instance.close()
        _overlay_instance = None


def overlay_is_open() -> bool:
    return _overlay_instance is not None
