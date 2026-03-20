"""
Little Helper - Fan control via LibreHardwareMonitor.

Controls chassis fan PWM based on GPU/CPU temperature (or manual input).
Fans with RPM >= 3000 at discovery time are treated as pumps and skipped.
Requires administrator privileges and LHM with motherboard hardware enabled.
"""

import threading
import logging

log = logging.getLogger("little_helper.fan_control")

_controller_thread: threading.Thread | None = None
_stop_event    = threading.Event()
_controls      = []          # cached IControl .NET objects
_controls_lock = threading.Lock()

# User intent: True if fan control was started (even if no fans found yet)
_fan_control_enabled = False

# Manual override percentage (0–100), used when source == "manual"
_manual_pct: float = 50.0


# ---------------------------------------------------------------------------
# Hardware discovery
# ---------------------------------------------------------------------------

def _discover_fan_controls(lhm_computer, fan_indices: list) -> list:
    """
    Walk Motherboard → SubHardware → Sensors(SensorType=Control) → sensor.Control.
    Fans with current RPM >= 3000 are treated as pumps and skipped automatically.
    Includes controls with ControlMode=Undefined — LHM sometimes cannot read the
    current mode but SetSoftware() still works on the chip.
    Must be called with lhm_lock already held by the caller.
    """
    controls = []
    names    = []
    try:
        for hw in lhm_computer.Hardware:
            if hw.HardwareType.ToString() != "Motherboard":
                continue
            for sub_hw in hw.SubHardware:
                try:
                    sub_hw.Update()
                except Exception:
                    pass
                # Collect RPM sensors by position for cross-reference and pump detection
                rpm_sensors = [s for s in sub_hw.Sensors
                               if s.SensorType.ToString() == "Fan"]
                raw_idx = 0   # counts all Control sensors (for RPM index alignment)
                for sensor in sub_hw.Sensors:
                    try:
                        if sensor.SensorType.ToString() != "Control":
                            continue
                        ctrl = sensor.Control
                        ri = raw_idx
                        raw_idx += 1
                        if ctrl is None:
                            continue

                        mode       = ctrl.ControlMode.ToString()
                        rpm_sensor = rpm_sensors[ri] if ri < len(rpm_sensors) else None
                        rpm_val    = None
                        if rpm_sensor is not None:
                            try:
                                v = rpm_sensor.Value
                                rpm_val = float(v) if v is not None else None
                            except Exception:
                                pass
                        rpm_name = rpm_sensor.Name if rpm_sensor else ""

                        # Skip pumps (high RPM headers)
                        if rpm_val is not None and rpm_val >= 3000:
                            log.info(
                                f"Skipping {sensor.Name} on {sub_hw.Name} "
                                f"(RPM={rpm_val:.0f} >= 3000, likely pump)"
                            )
                            continue

                        fan_idx   = len(controls)
                        rpm_hint  = f', RPM sensor: "{rpm_name}"' if rpm_name else ""
                        rpm_cur   = f" | RPM={rpm_val:.0f}" if rpm_val is not None else ""
                        log.info(
                            f"[{fan_idx}] {sensor.Name} on {sub_hw.Name} "
                            f"| id={sensor.Identifier} | mode={mode}{rpm_hint}{rpm_cur}"
                        )
                        controls.append(ctrl)
                        names.append(sensor.Name)
                    except Exception as e:
                        log.debug(f"Control probe error on {sub_hw.Name}: {e}")
    except Exception as e:
        log.warning(f"Fan control discovery error: {e}")

    if fan_indices:
        filtered = [(c, n) for i, (c, n) in enumerate(zip(controls, names))
                    if i in fan_indices]
        controls = [c for c, _ in filtered]
        log.info(f"Filtered to fan_indices={fan_indices}: {len(controls)} control(s)")

    if not controls:
        log.warning(
            "No fan controls found. Check: (1) running as admin, "
            "(2) Fan Control software is NOT open (conflicts with LHM), "
            "(3) your motherboard SuperIO chip is supported by LHM."
        )
    return controls


# ---------------------------------------------------------------------------
# Fan curve interpolation
# ---------------------------------------------------------------------------

def _interpolate_curve(source_val: float, curve: list) -> float:
    """Linear interpolation; returns a percentage clamped to [0, 100]."""
    if not curve or source_val is None:
        return 30.0

    if source_val <= curve[0][0]:
        return float(curve[0][1])
    if source_val >= curve[-1][0]:
        return float(curve[-1][1])

    for i in range(len(curve) - 1):
        x0, y0 = curve[i]
        x1, y1 = curve[i + 1]
        if x0 <= source_val <= x1:
            t = (source_val - x0) / (x1 - x0)
            return max(0.0, min(100.0, y0 + t * (y1 - y0)))

    return float(curve[-1][1])


# ---------------------------------------------------------------------------
# Source value readers
# ---------------------------------------------------------------------------

def _get_source_value(source: str):
    """Return the current control input value (temp °C or manual %).
    Sources: gpu_temp | cpu_temp | mixed (max of GPU and CPU) | manual
    """
    if source == "manual":
        return _manual_pct
    try:
        import system_overlay
        gpu_temp = None
        cpu_temp = None

        if source in ("gpu_temp", "mixed"):
            try:
                import pynvml
                if system_overlay._nvml_available and system_overlay._nvml_handle is not None:
                    h = system_overlay._nvml_handle
                    gpu_temp = float(pynvml.nvmlDeviceGetTemperature(
                        h, pynvml.NVML_TEMPERATURE_GPU))
            except Exception as e:
                log.debug(f"GPU temp read error: {e}")

        if source in ("cpu_temp", "mixed"):
            try:
                s = system_overlay._lhm_cpu_temp
                if s is not None:
                    v = s.Value
                    cpu_temp = float(v) if v is not None else None
            except Exception as e:
                log.debug(f"CPU temp read error: {e}")

        if source == "gpu_temp":
            return gpu_temp * 1.1 if gpu_temp is not None else None
        elif source == "cpu_temp":
            return cpu_temp
        elif source == "mixed":
            # GPU temperature is weighted +10% to make fans more responsive to GPU load
            vals = [v for v in (
                gpu_temp * 1.1 if gpu_temp is not None else None,
                cpu_temp,
            ) if v is not None]
            return max(vals) if vals else None
    except Exception as e:
        log.debug(f"_get_source_value({source}) error: {e}")
    return None


# ---------------------------------------------------------------------------
# Control loop
# ---------------------------------------------------------------------------

def _control_loop(config: dict, lhm_computer, lhm_lock: threading.Lock) -> None:
    fc_cfg      = config.get("fan_control", {})
    source      = fc_cfg.get("source", "gpu_temp")
    interval_s  = max(1, fc_cfg.get("interval_s", 3))
    raw_curve   = fc_cfg.get("curve", [[40, 30], [60, 50], [75, 75], [85, 100]])
    fan_indices = fc_cfg.get("fan_indices", [])

    curve = sorted([(float(x), float(y)) for x, y in raw_curve])

    with lhm_lock:
        fan_controls = _discover_fan_controls(lhm_computer, fan_indices)

    with _controls_lock:
        _controls.clear()
        _controls.extend(fan_controls)

    if not fan_controls:
        log.warning("Fan control loop exiting: no controllable fans found")
        return

    log.info(
        f"Fan control loop: source={source}, interval={interval_s}s, "
        f"fans={len(fan_controls)}, curve={curve}"
    )

    last_target: float | None = None

    while not _stop_event.is_set():
        try:
            val = _get_source_value(source)
            if val is not None:
                # Manual mode: val IS the target percentage
                target_pct = val if source == "manual" else _interpolate_curve(val, curve)
                if last_target is None or abs(target_pct - last_target) >= 1.0:
                    log.debug(f"Fan control: {source}={val:.1f} -> {target_pct:.1f}%")
                    last_target = target_pct
                with lhm_lock:
                    for ctrl in fan_controls:
                        try:
                            ctrl.SetSoftware(target_pct)
                        except Exception as e:
                            log.warning(f"SetSoftware failed: {e}")
            else:
                log.debug(f"Fan control: source unavailable for '{source}'")
        except Exception as e:
            log.error(f"Fan control loop error: {e}", exc_info=True)

        _stop_event.wait(interval_s)

    # Restore automatic fan control on exit
    log.info("Fan control stopping, restoring automatic mode...")
    with lhm_lock:
        with _controls_lock:
            failed = []
            for i, ctrl in enumerate(_controls):
                try:
                    ctrl.SetDefault()
                except Exception as e:
                    failed.append(f"#{i+1}: {e}")
            n = len(_controls)
            if failed:
                log.warning(f"SetDefault failed for {failed}")
            else:
                log.info(f"Restored {n} fan(s) to automatic mode")
            _controls.clear()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def set_manual_pct(pct: float) -> None:
    """Set the manual fan speed percentage (0–100). Applied immediately if running."""
    global _manual_pct
    new = max(0.0, min(100.0, float(pct)))
    if new != _manual_pct:
        _manual_pct = new
        log.debug(f"Manual fan pct set to {_manual_pct:.1f}%")


def start_fan_control(config: dict, lhm_computer, lhm_lock: threading.Lock) -> None:
    """Start the fan control thread. Idempotent."""
    global _controller_thread, _fan_control_enabled
    _fan_control_enabled = True
    if _controller_thread is not None and _controller_thread.is_alive():
        log.debug("Fan control already running")
        return

    _stop_event.clear()
    _controller_thread = threading.Thread(
        target=_control_loop,
        args=(config, lhm_computer, lhm_lock),
        daemon=True,
        name="fan-control",
    )
    _controller_thread.start()
    log.info("Fan control thread started")


def stop_fan_control() -> None:
    """Stop the fan control thread and restore automatic fan mode."""
    global _controller_thread, _fan_control_enabled
    _fan_control_enabled = False
    _stop_event.set()
    if _controller_thread is not None:
        _controller_thread.join(timeout=5)
        _controller_thread = None
    log.info("Fan control stopped")


def fan_control_is_active() -> bool:
    """True if the control thread is alive (fans actually being controlled)."""
    return _controller_thread is not None and _controller_thread.is_alive()


def fan_control_is_enabled() -> bool:
    """True if the user has turned on fan control (reflects intent, not hardware state)."""
    return _fan_control_enabled


# ---------------------------------------------------------------------------
# GPU Fan Control (pynvml — Nvidia only, source fixed to GPU temp)
# ---------------------------------------------------------------------------

_gpu_controller_thread: threading.Thread | None = None
_gpu_stop_event    = threading.Event()
_gpu_fan_control_enabled = False

_gpu_manual_pct: float = 50.0


def _gpu_control_loop(config: dict) -> None:
    fc_cfg     = config.get("gpu_fan_control", {})
    source     = fc_cfg.get("source", "gpu_temp")
    interval_s = max(1, fc_cfg.get("interval_s", 2))
    raw_curve  = fc_cfg.get("curve", [[40, 30], [60, 50], [70, 75], [80, 100]])
    curve      = sorted([(float(x), float(y)) for x, y in raw_curve])

    try:
        import pynvml
        import system_overlay
        if not system_overlay._nvml_available or system_overlay._nvml_handle is None:
            log.warning("GPU fan control: pynvml not available or no GPU")
            return
        handle = system_overlay._nvml_handle
        try:
            n_fans = pynvml.nvmlDeviceGetNumFans(handle)
        except Exception:
            n_fans = 1  # older driver: assume 1 fan
    except Exception as e:
        log.warning(f"GPU fan control init error: {e}")
        return

    if n_fans == 0:
        log.warning("GPU fan control: no fans found on GPU")
        return

    log.info(
        f"GPU fan control: {n_fans} fan(s), source={source}, "
        f"interval={interval_s}s, curve={curve}"
    )

    last_target: float | None = None

    while not _gpu_stop_event.is_set():
        try:
            if source == "manual":
                target_pct = _gpu_manual_pct
                val_str    = f"manual={target_pct:.1f}"
                val_ok     = True
            else:
                gpu_temp = None
                try:
                    gpu_temp = float(pynvml.nvmlDeviceGetTemperature(
                        handle, pynvml.NVML_TEMPERATURE_GPU))
                except Exception as e:
                    log.debug(f"GPU temp read error: {e}")
                if gpu_temp is None:
                    log.debug("GPU fan control: GPU temp unavailable")
                    _gpu_stop_event.wait(interval_s)
                    continue
                target_pct = _interpolate_curve(gpu_temp, curve)
                val_str    = f"temp={gpu_temp:.1f}"
                val_ok     = True

            if val_ok:
                if last_target is None or abs(target_pct - last_target) >= 1.0:
                    log.debug(f"GPU fan: {val_str} -> {target_pct:.1f}%")
                    last_target = target_pct
                for i in range(n_fans):
                    try:
                        policy_manual = getattr(pynvml, "NVML_FAN_POLICY_MANUAL", 1)
                        pynvml.nvmlDeviceSetFanControlPolicy(handle, i, policy_manual)
                    except Exception:
                        pass
                    try:
                        pynvml.nvmlDeviceSetFanSpeed_v2(handle, i, int(target_pct))
                    except AttributeError:
                        try:
                            pynvml.nvmlDeviceSetFanSpeed(handle, i, int(target_pct))
                        except Exception as e:
                            log.warning(f"GPU fan #{i} set error: {e}")
                    except Exception as e:
                        log.warning(f"GPU fan #{i} set error: {e}")
        except Exception as e:
            log.error(f"GPU fan control loop error: {e}", exc_info=True)

        _gpu_stop_event.wait(interval_s)

    # Restore auto mode
    log.info("GPU fan control stopping, restoring auto mode...")
    failed = []
    for i in range(n_fans):
        try:
            policy_auto = getattr(pynvml, "NVML_FAN_POLICY_TEMPERATURE_CONTINOUS_SW", 0)
            pynvml.nvmlDeviceSetFanControlPolicy(handle, i, policy_auto)
        except Exception as e:
            failed.append(f"#{i}: {e}")
    if failed:
        log.warning(f"GPU fan restore failed: {failed}")
    else:
        log.info(f"Restored {n_fans} GPU fan(s) to auto mode")


def set_gpu_manual_pct(pct: float) -> None:
    """Set the manual GPU fan speed percentage (0–100). Applied immediately if running."""
    global _gpu_manual_pct
    new = max(0.0, min(100.0, float(pct)))
    if new != _gpu_manual_pct:
        _gpu_manual_pct = new
        log.debug(f"Manual GPU fan pct set to {_gpu_manual_pct:.1f}%")


def start_gpu_fan_control(config: dict) -> None:
    """Start the GPU fan control thread. Idempotent."""
    global _gpu_controller_thread, _gpu_fan_control_enabled
    _gpu_fan_control_enabled = True
    if _gpu_controller_thread is not None and _gpu_controller_thread.is_alive():
        log.debug("GPU fan control already running")
        return

    _gpu_stop_event.clear()
    _gpu_controller_thread = threading.Thread(
        target=_gpu_control_loop,
        args=(config,),
        daemon=True,
        name="gpu-fan-control",
    )
    _gpu_controller_thread.start()
    log.info("GPU fan control thread started")


def stop_gpu_fan_control() -> None:
    """Stop the GPU fan control thread and restore auto fan mode."""
    global _gpu_controller_thread, _gpu_fan_control_enabled
    _gpu_fan_control_enabled = False
    _gpu_stop_event.set()
    if _gpu_controller_thread is not None:
        _gpu_controller_thread.join(timeout=5)
        _gpu_controller_thread = None
    log.info("GPU fan control stopped")


def gpu_fan_control_is_active() -> bool:
    """True if the GPU fan control thread is alive."""
    return _gpu_controller_thread is not None and _gpu_controller_thread.is_alive()


def gpu_fan_control_is_enabled() -> bool:
    """True if the user has turned on GPU fan control."""
    return _gpu_fan_control_enabled
