"""
Little Helper - Nvidia GPU power limit control via nvidia-smi.
Requires administrator privileges to set the power limit.
"""

import os
import sys
import ctypes
import subprocess
import logging

log = logging.getLogger("little_helper.gpu_power")

# Fallback path for nvidia-smi if not on PATH
_NVSMI_FALLBACK = r"C:\Program Files\NVIDIA Corporation\NVSMI\nvidia-smi.exe"

_original_watts: float | None = None  # stored on first apply, restored on exit


def _run_nvidia_smi(args: list[str]) -> tuple[int, str, str]:
    """Run nvidia-smi with given args. Returns (returncode, stdout, stderr)."""
    for exe in ("nvidia-smi", _NVSMI_FALLBACK):
        try:
            result = subprocess.run(
                [exe] + args,
                capture_output=True,
                text=True,
                timeout=8,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            return result.returncode, result.stdout.strip(), result.stderr.strip()
        except FileNotFoundError:
            continue
        except subprocess.TimeoutExpired:
            return -1, "", "nvidia-smi timed out"
        except Exception as e:
            return -1, "", str(e)
    return -1, "", "nvidia-smi not found"


def get_gpu_power_limits() -> tuple[float, float, float] | None:
    """
    Query GPU min/max/current power limits.
    Returns (min_w, max_w, current_w) or None if unavailable.
    """
    rc, out, err = _run_nvidia_smi([
        "--query-gpu=power.min_limit,power.max_limit,power.limit",
        "--format=csv,noheader,nounits",
    ])
    if rc != 0 or not out:
        log.debug(f"get_gpu_power_limits failed: rc={rc} err={err}")
        return None
    try:
        parts = [p.strip() for p in out.split(",")]
        return float(parts[0]), float(parts[1]), float(parts[2])
    except Exception as e:
        log.warning(f"Failed to parse power limits '{out}': {e}")
        return None


def is_admin() -> bool:
    """Return True if the current process has administrator privileges."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> None:
    """Re-launch this process with UAC elevation, then exit the current one."""
    script = sys.argv[0]
    params = " ".join(f'"{a}"' for a in sys.argv[1:])
    log.info(f"Relaunching as admin: {script} {params}")
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{script}" {params}', None, 1
    )


def apply_gpu_power_limit(config: dict, notify_fn=None) -> None:
    """
    Read config["gpu_power_limit"] and apply the target watt limit.
    Stores the original limit so it can be restored on exit.
    """
    global _original_watts

    gpu_cfg = config.get("gpu_power_limit", {})
    if not gpu_cfg.get("enabled", False):
        return

    target_w = int(gpu_cfg.get("watts", 150))

    # Query and store original limit
    limits = get_gpu_power_limits()
    if limits is None:
        log.warning("No Nvidia GPU found or nvidia-smi unavailable, skipping power limit")
        return

    min_w, max_w, current_w = limits
    _original_watts = current_w
    log.info(f"GPU power limits: min={min_w}W max={max_w}W current={current_w}W target={target_w}W")

    # Clamp to valid range
    target_w = max(int(min_w), min(int(max_w), target_w))

    rc, out, err = _run_nvidia_smi(["-pl", str(target_w)])
    if rc == 0:
        log.info(f"GPU power limit set to {target_w}W")
    else:
        log.warning(f"Failed to set GPU power limit: {err}")
        if not is_admin():
            msg = "GPU power limit requires administrator rights.\nUse 'Relaunch as Administrator' in the tray menu."
        else:
            msg = f"Failed to set GPU power limit:\n{err}"
        if notify_fn:
            notify_fn(msg, "GPU Power Limit")


def restore_gpu_power_limit() -> None:
    """Restore the original GPU power limit (idempotent)."""
    global _original_watts
    if _original_watts is None:
        return
    watts = _original_watts
    _original_watts = None
    log.info(f"Restoring GPU power limit to {watts}W")
    rc, out, err = _run_nvidia_smi(["-pl", str(int(watts))])
    if rc == 0:
        log.info(f"GPU power limit restored to {watts}W")
    else:
        log.warning(f"Failed to restore GPU power limit: {err}")
