"""
Little Helper - System monitoring overlay window.

Shows RAM, CPU, GPU stats in a draggable, resizable, semi-transparent overlay.
"""

import copy
import os
import sys
import queue
import time
import threading
import logging
import subprocess
from collections import Counter, defaultdict
import tkinter as tk
from tkinter import font as tkfont
from datetime import datetime, timezone

log = logging.getLogger("little_helper.system_overlay")

# --- NVML state (initialised once at startup) ---
_nvml_available = False
_nvml_handle    = None

# --- LibreHardwareMonitor state ---
_lhm_available = False
_lhm_computer  = None
_lhm_cpu_temp  = None   # ISensor reference
_lhm_cpu_power = None   # ISensor reference
_lhm_ram_temps = []     # list of ISensor references (one per DIMM)
_lhm_disk_temps = {}    # dict: {unique_disk_name: ISensor} for disk temperatures
_lhm_disk_activity = {} # dict: {unique_disk_name: ISensor} for disk activity time percentage
_lhm_disk_storage = {}  # dict: {unique_disk_name: storage object} for runtime disk naming
_lhm_disk_display_name_lookup = {}
_lhm_lock      = threading.Lock()  # serialises all LHM .NET object access

# --- Disk wake-up state (HDD management) ---
_disk_type_cache = {}   # {drive_number: "HDD"|"SSD"|"Unknown"} - cached disk types
_disk_wakeup_targets = {}  # {drive_number: tuple[str, ...]} - cached logical drives for wake-up
_disk_inventory_signature = ()  # tuple[int, ...] of current LHM drive numbers for cache invalidation
_hdd_wakeup_drive_letters = ()  # tuple[str, ...] - cached drive letters for HDD wake-up only
_disk_wakeup_lock = threading.Lock()  # serialises disk wake-up operations
_ui_root       = None
_ui_thread_id  = None
_ui_tasks: queue.Queue = queue.Queue()
_snapshot_lock = threading.Lock()
_snapshot_cache = None
_snapshot_cache_at = 0.0
_MIN_SNAPSHOT_CACHE_MS = 500


def _set_overlay_enabled_in_config(config: dict, save_config_fn, enabled: bool) -> bool:
    overlay_cfg = config.setdefault("overlay", {})
    enabled = bool(enabled)
    if bool(overlay_cfg.get("enabled", False)) == enabled:
        return False

    overlay_cfg["enabled"] = enabled
    save_config_fn(config)
    return True


def _disk_temp_sensor_priority(sensor_name: str) -> tuple[int, int]:
    name = (sensor_name or "").strip().lower()
    if name == "temperature":
        return (0, 0)
    if name.startswith("temperature #"):
        try:
            return (1, int(name.split("#", 1)[1].strip()))
        except (IndexError, ValueError):
            return (1, 99)
    if "warning" in name:
        return (2, 0)
    if "critical" in name:
        return (3, 0)
    return (4, 0)


def _disk_activity_sensor_priority(sensor_name: str) -> tuple[int, int]:
    """Priority for selecting disk activity time sensor."""
    name = (sensor_name or "").strip().lower()
    if name == "active time":
        return (0, 0)
    if "active" in name and "time" in name:
        return (1, 0)
    if name == "load":
        return (2, 0)
    if "activity" in name or "utilization" in name or "usage" in name:
        return (3, 0)
    return (4, 0)


def _serial_suffix(serial_number) -> str | None:
    serial = "".join(ch for ch in str(serial_number or "").upper() if ch.isalnum())
    if not serial:
        return None
    return serial[-4:] if len(serial) >= 4 else serial


def _normalize_disk_name(name) -> str:
    text = str(name or "Unknown").strip()
    # Remove trailing parentheses and their content (virtual identifiers)
    import re
    model = re.sub(r'\s*\([^)]*\)\s*$', '', text).strip()
    return " ".join(model.split()) or "Unknown"


def _build_windows_disk_serial_suffix_map(entries) -> dict[str, list[str]]:
    suffixes: dict[str, list[str]] = defaultdict(list)
    for entry in sorted(entries or [], key=lambda item: int(item.get("Index", 0))):
        model = _normalize_disk_name(entry.get("Model"))
        suffix = _serial_suffix(entry.get("SerialNumber"))
        if model and suffix:
            suffixes[model].append(suffix)
    return dict(suffixes)


def _get_lhm_disk_serial_suffix_map() -> dict[str, list[str]]:
    suffixes: dict[str, list[str]] = defaultdict(list)
    storages = sorted(
        _lhm_disk_storage.items(),
        key=lambda item: getattr(item[1], "DriveNumber", sys.maxsize),
    )
    for disk_name, storage in storages:
        model = _normalize_disk_name(getattr(storage, "Model", None) or disk_name)
        suffix = _serial_suffix(getattr(storage, "SerialNumber", None))
        if model and suffix:
            suffixes[model].append(suffix)
    return dict(suffixes)


def _build_lhm_disk_display_name_lookup(storages) -> dict[tuple[str, str | None], str]:
    ordered_storages = sorted(
        storages,
        key=lambda item: getattr(item[1], "DriveNumber", sys.maxsize) if item[1] is not None else sys.maxsize,
    )
    serial_suffix_map: dict[str, list[str]] = defaultdict(list)
    for hardware, storage in ordered_storages:
        model = _normalize_disk_name(getattr(storage, "Model", None) or hardware.Name)
        suffix = _serial_suffix(getattr(storage, "SerialNumber", None))
        if model and suffix:
            serial_suffix_map[model].append(suffix)

    display_names = _assign_unique_disk_names(
        [_normalize_disk_name(getattr(storage, "Model", None) or hardware.Name) for hardware, storage in ordered_storages],
        dict(serial_suffix_map),
    )
    display_lookup = {}
    for (hardware, storage), display_name in zip(ordered_storages, display_names):
        model = _normalize_disk_name(getattr(storage, "Model", None) or hardware.Name)
        if storage is None:
            continue

        drive_number = getattr(storage, "DriveNumber", None)
        if drive_number is not None:
            try:
                display_lookup[(model, f"index:{int(drive_number)}")] = display_name
            except (TypeError, ValueError):
                pass

        serial_suffix = _serial_suffix(getattr(storage, "SerialNumber", None))
        if serial_suffix:
            display_lookup[(model, f"serial:{serial_suffix}")] = display_name

    return display_lookup


def _lookup_disk_display_names(display_name_lookup, normalized_name: str) -> set[str]:
    return {
        display_name
        for (model, lookup_key), display_name in (display_name_lookup or {}).items()
        if model == normalized_name and str(lookup_key).startswith("index:")
    }


def _resolve_disk_display_name(
    disk_name,
    serial_number,
    drive_number=None,
    display_name_lookup: dict[tuple[str, str | None], str] | None = None,
) -> str:
    normalized_name = _normalize_disk_name(disk_name)
    lookup = display_name_lookup or {}
    serial_suffix = _serial_suffix(serial_number)
    if serial_suffix:
        serial_key = (normalized_name, f"serial:{serial_suffix}")
        if serial_key in lookup:
            return lookup[serial_key]

        if len(_lookup_disk_display_names(lookup, normalized_name)) > 1:
            return f"{normalized_name} ({serial_suffix})"

    if drive_number is not None:
        try:
            drive_key = (normalized_name, f"index:{int(drive_number)}")
            if drive_key in lookup:
                return lookup[drive_key]
        except (TypeError, ValueError):
            pass

    return normalized_name


def _iter_hardware_tree(root_hardware):
    stack = [root_hardware]
    while stack:
        hardware = stack.pop()
        yield hardware
        try:
            stack.extend(reversed(list(hardware.SubHardware)))
        except Exception:
            pass


def _iter_storage_hardware(_lhm_computer_instance):
    if _lhm_computer_instance is None:
        return
    for hardware in _lhm_computer_instance.Hardware:
        for node in _iter_hardware_tree(hardware):
            try:
                if node.HardwareType.ToString() == "Storage":
                    yield node
            except Exception:
                continue


def _get_storage_object(hardware):
    storage_prop = hardware.GetType().GetProperty("Storage")
    return storage_prop.GetValue(hardware) if storage_prop is not None else None


def _get_disk_temp_sensor_candidates(hardware):
    candidates = []
    for node in _iter_hardware_tree(hardware):
        try:
            node.Update()
        except Exception:
            pass
        for sensor in node.Sensors:
            if sensor.SensorType.ToString() != "Temperature":
                continue
            candidates.append(sensor)
    return candidates


def _select_best_disk_temp_sensor(hardware):
    selected_sensor = None
    for sensor in _get_disk_temp_sensor_candidates(hardware):
        if selected_sensor is None or _disk_temp_sensor_priority(sensor.Name) < _disk_temp_sensor_priority(selected_sensor.Name):
            selected_sensor = sensor
    return selected_sensor


def _get_disk_activity_sensor_candidates(hardware):
    """Get all Load/Activity time sensors from a disk hardware node."""
    candidates = []
    for node in _iter_hardware_tree(hardware):
        try:
            node.Update()
        except Exception:
            pass
        for sensor in node.Sensors:
            sensor_type = sensor.SensorType.ToString()
            # Load sensors typically represent disk activity percentage
            if sensor_type in ("Load",):
                candidates.append(sensor)
    return candidates


def _select_best_disk_activity_sensor(hardware):
    """Select the best disk activity sensor from candidates."""
    selected_sensor = None
    for sensor in _get_disk_activity_sensor_candidates(hardware):
        if selected_sensor is None or _disk_activity_sensor_priority(sensor.Name) < _disk_activity_sensor_priority(selected_sensor.Name):
            selected_sensor = sensor
    return selected_sensor


def _refresh_lhm_storage_state(refresh_sensor_bindings: bool = False) -> None:
    global _lhm_disk_display_name_lookup, _lhm_disk_storage, _lhm_disk_temps, _lhm_disk_activity

    if _lhm_computer is None:
        return

    storage_nodes = []
    for hardware in _iter_storage_hardware(_lhm_computer):
        try:
            hardware.Update()
        except Exception:
            pass

        try:
            stor_obj = _get_storage_object(hardware)
        except Exception as exc:
            stor_obj = None
            log.debug("Failed to get Storage object for %s: %s", _normalize_disk_name(hardware.Name), exc)

        storage_nodes.append((hardware, stor_obj))

    if refresh_sensor_bindings or not _lhm_disk_display_name_lookup:
        _lhm_disk_display_name_lookup = _build_lhm_disk_display_name_lookup(storage_nodes)

    updated_disk_storage = {}
    updated_disk_temps = {}
    updated_disk_activity = {}

    for hardware, stor_obj in storage_nodes:

        disk_name = _resolve_disk_display_name(
            hardware.Name,
            getattr(stor_obj, 'SerialNumber', None),
            getattr(stor_obj, 'DriveNumber', None),
            _lhm_disk_display_name_lookup,
        )

        if stor_obj is not None:
            updated_disk_storage[disk_name] = stor_obj

        if refresh_sensor_bindings or disk_name not in _lhm_disk_temps:
            selected_sensor = _select_best_disk_temp_sensor(hardware)
            if selected_sensor is not None:
                updated_disk_temps[disk_name] = selected_sensor
        elif disk_name in _lhm_disk_temps:
            updated_disk_temps[disk_name] = _lhm_disk_temps[disk_name]

        if refresh_sensor_bindings or disk_name not in _lhm_disk_activity:
            selected_activity_sensor = _select_best_disk_activity_sensor(hardware)
            if selected_activity_sensor is not None:
                updated_disk_activity[disk_name] = selected_activity_sensor
        elif disk_name in _lhm_disk_activity:
            updated_disk_activity[disk_name] = _lhm_disk_activity[disk_name]

    _refresh_disk_wakeup_cache(updated_disk_storage)

    _lhm_disk_storage = updated_disk_storage
    _lhm_disk_temps = updated_disk_temps
    _lhm_disk_activity = updated_disk_activity


def _assign_unique_disk_names(disk_names: list[str], serial_suffix_map: dict[str, list[str]] | None = None) -> list[str]:
    normalized_names = [_normalize_disk_name(name) for name in disk_names]
    counts = Counter(normalized_names)
    occurrences: dict[str, int] = defaultdict(int)
    serial_suffix_map = {
        _normalize_disk_name(name): list(suffixes)
        for name, suffixes in (serial_suffix_map or {}).items()
    }
    unique_names = []

    for disk_name in normalized_names:
        known_duplicates = len(serial_suffix_map.get(disk_name, [])) > 1
        if counts[disk_name] <= 1 and not known_duplicates:
            unique_names.append(disk_name)
            continue

        idx = occurrences[disk_name]
        occurrences[disk_name] += 1
        suffixes = serial_suffix_map.get(disk_name, [])
        if idx < len(suffixes):
            suffix = suffixes[idx]
        else:
            suffix = str(idx + 1)
        unique_names.append(f"{disk_name} ({suffix})")

    return unique_names


def _rename_disk_temp_values(disk_values: dict[str, float]) -> dict[str, float]:
    return dict(disk_values)


def _classify_disk_media_type(media_type) -> str:
    if media_type is None:
        return "Unknown"

    media_str = str(media_type).strip().lower()
    if "ssd" in media_str or "nvme" in media_str or "solid" in media_str:
        return "SSD"
    if "fixed hard disk" in media_str or "hard disk" in media_str:
        return "HDD"

    try:
        media_int = int(media_type)
    except (TypeError, ValueError):
        return "Unknown"

    if media_int == 3:
        return "HDD"
    if media_int == 4:
        return "SSD"
    return "Unknown"


def _escape_wmi_associator_value(value) -> str:
    return str(value or "").replace("\\", "\\\\")


def _logical_disk_name_to_letter(name) -> str | None:
    text = str(name or "").strip().rstrip(":").upper()
    if len(text) == 1 and text.isalpha():
        return text
    return None


def _get_disk_logical_drive_letters(wmi_client, disk_device_id) -> tuple[str, ...]:
    if not disk_device_id:
        return ()

    letters = []
    seen = set()
    escaped_disk_id = _escape_wmi_associator_value(disk_device_id)
    partitions = wmi_client.query(
        f'ASSOCIATORS OF {{Win32_DiskDrive.DeviceID="{escaped_disk_id}"}} '
        'WHERE AssocClass = Win32_DiskDriveToDiskPartition'
    )
    for partition in partitions:
        partition_id = getattr(partition, 'DeviceID', None)
        if not partition_id:
            continue

        escaped_partition_id = _escape_wmi_associator_value(partition_id)
        logical_disks = wmi_client.query(
            f'ASSOCIATORS OF {{Win32_DiskPartition.DeviceID="{escaped_partition_id}"}} '
            'WHERE AssocClass = Win32_LogicalDiskToPartition'
        )
        for logical_disk in logical_disks:
            drive_letter = _logical_disk_name_to_letter(getattr(logical_disk, 'Name', None))
            if drive_letter and drive_letter not in seen:
                seen.add(drive_letter)
                letters.append(drive_letter)

    return tuple(letters)


def _get_windows_disk_inventory() -> dict[int, dict[str, object]]:
    inventory: dict[int, dict[str, object]] = {}
    try:
        import wmi

        w = wmi.WMI()
        disks = w.query("SELECT DeviceID, Index, MediaType FROM Win32_DiskDrive")
        for disk in disks:
            drive_number = getattr(disk, 'Index', None)
            try:
                drive_number = int(drive_number)
            except (TypeError, ValueError):
                continue

            inventory[drive_number] = {
                "disk_type": _classify_disk_media_type(getattr(disk, 'MediaType', None)),
                "drive_letters": _get_disk_logical_drive_letters(w, getattr(disk, 'DeviceID', None)),
            }
    except Exception as exc:
        log.debug(f"Failed to build Windows disk inventory: {exc}")

    return inventory


def _refresh_disk_wakeup_cache(storages: dict[str, object] | None = None) -> None:
    global _disk_inventory_signature, _disk_type_cache, _disk_wakeup_targets, _hdd_wakeup_drive_letters

    storage_map = storages if storages is not None else _lhm_disk_storage
    drive_numbers = []
    for storage in (storage_map or {}).values():
        if storage is None:
            continue
        drive_number = getattr(storage, 'DriveNumber', None)
        try:
            drive_numbers.append(int(drive_number))
        except (TypeError, ValueError):
            continue

    signature = tuple(sorted(set(drive_numbers)))
    if not signature:
        _disk_inventory_signature = ()
        _disk_wakeup_targets = {}
        _hdd_wakeup_drive_letters = ()
        return

    cache_complete = (
        signature == _disk_inventory_signature
        and all(drive_number in _disk_type_cache for drive_number in signature)
        and all(drive_number in _disk_wakeup_targets for drive_number in signature)
    )
    if cache_complete:
        return

    inventory = _get_windows_disk_inventory()
    updated_types = {}
    updated_targets = {}
    for drive_number in signature:
        entry = inventory.get(drive_number, {})
        updated_types[drive_number] = entry.get("disk_type", "Unknown")
        updated_targets[drive_number] = tuple(entry.get("drive_letters", ()))

    _disk_type_cache.update(updated_types)
    _disk_wakeup_targets.update(updated_targets)
    _disk_inventory_signature = signature
    _hdd_wakeup_drive_letters = tuple(
        drive_letter
        for drive_number in signature
        if _disk_type_cache.get(drive_number) == "HDD"
        for drive_letter in _disk_wakeup_targets.get(drive_number, ())
    )


def _detect_disk_type_via_wmi(drive_number: int) -> str:
    """Detect if disk is HDD or SSD via Windows WMI.
    
    Args:
        drive_number: Physical disk index from Storage object (0, 1, 2, ...)
        
    Returns:
        "HDD", "SSD", or "Unknown"
        
    MediaType values from Win32_DiskDrive:
    - 3: Internal Fixed Disk (HDD)
    - 4: Removable Media / SSD
    - 5: External Hard Disk
    - Enum name: "Fixed hard disk media", "Removable media other than floppy", etc.
    """
    # Check cache first
    if drive_number in _disk_type_cache:
        return _disk_type_cache[drive_number]

    result = "Unknown"
    try:
        inventory = _get_windows_disk_inventory()
        if drive_number in inventory:
            result = inventory[drive_number].get("disk_type", "Unknown")
            _disk_wakeup_targets[drive_number] = tuple(inventory[drive_number].get("drive_letters", ()))
            log.debug(f"Disk #{drive_number}: cached inventory -> {result}")
        else:
            log.debug(f"No disk found with Index={drive_number} in WMI")
    except Exception as e:
        log.debug(f"Failed to detect disk type for drive #{drive_number}: {e}")

    _disk_type_cache[drive_number] = result
    return result


def _get_drive_letter_from_number(drive_number: int) -> str | None:
    """Map physical disk number to the first cached drive letter.
    
    Args:
        drive_number: Physical disk index (0=first disk, 1=second disk, ...)
        
    Returns:
        Drive letter (e.g., "C", "D") or None if not found
    """
    drive_letters = _disk_wakeup_targets.get(drive_number, ())
    if drive_letters:
        return drive_letters[0]
    return None


def _wake_hdd_via_io(drive_letter: str) -> bool:
    """Wake HDD by performing small I/O read on disk root.
    
    This triggers Windows driver to spin up the disk.
    
    Args:
        drive_letter: Drive letter (e.g., "C", "D") without colon
        
    Returns:
        True if operation completed, False on error
    """
    try:
        import pathlib
        root_path = pathlib.Path(f"{drive_letter}:\\")
        
        if root_path.exists():
            # Trigger disk I/O by listing root directory
            # This causes Windows to wake up the disk if it's in sleep mode
            list(root_path.iterdir())
            log.debug(f"Woke up disk {drive_letter}: via I/O")
            return True
    except Exception as e:
        log.debug(f"Failed to wake up disk {drive_letter}: {e}")
    
    return False


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
    """Attempt to initialise LibreHardwareMonitorLib for CPU/RAM/Disk sensors. Call once at startup."""
    global _lhm_available, _lhm_computer, _lhm_cpu_temp, _lhm_cpu_power, _lhm_ram_temps, _lhm_disk_temps, _lhm_disk_activity, _lhm_disk_storage, _lhm_disk_display_name_lookup
    try:
        import clr
        # Find the DLL path (works for both source and PyInstaller frozen EXE)
        if getattr(sys, 'frozen', False):
            dll_dir = os.path.join(sys._MEIPASS, "lhm")
        else:
            dll_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "lib", "lhm")
        if not os.path.exists(dll_dir):
            log.debug(f"LibreHardwareMonitor DLLs not found at {dll_dir}")
            return False

        # Add reference to the DLL
        clr.AddReference(os.path.join(dll_dir, "LibreHardwareMonitorLib.dll"))
        from LibreHardwareMonitor.Hardware import Computer

        _lhm_cpu_temp = None
        _lhm_cpu_power = None
        _lhm_ram_temps = []
        _lhm_disk_temps = {}
        _lhm_disk_activity = {}
        _lhm_disk_storage = {}
        _lhm_disk_display_name_lookup = {}

        _lhm_computer = Computer()
        _lhm_computer.IsCpuEnabled = True
        _lhm_computer.IsGpuEnabled = False
        _lhm_computer.IsMemoryEnabled = True
        _lhm_computer.IsMotherboardEnabled = True
        _lhm_computer.IsControllerEnabled = True   # needed for SMBus (DIMM temps)
        _lhm_computer.IsNetworkEnabled = False
        _lhm_computer.IsStorageEnabled = True    # needed for disk temperatures
        _lhm_computer.Open()

        for hardware in _lhm_computer.Hardware:
            hw_type = hardware.HardwareType.ToString()
            hardware.Update()

            if hw_type == "Cpu":
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

            elif hw_type != "Storage":
                # RAM temps may appear under SMBus, EmbeddedController, or other
                # hardware types — scan all non-CPU hardware for DIMM/DDR temp sensors
                _RAM_KEYWORDS = ("ddr", "dimm", "memory", "mem ", "mem#", "channel")
                for node in list(hardware.SubHardware) + [hardware]:
                    try:
                        node.Update()
                    except Exception:
                        pass
                    for sensor in node.Sensors:
                        if sensor.SensorType.ToString() != "Temperature":
                            continue
                        name_lower = sensor.Name.lower()
                        if any(kw in name_lower for kw in _RAM_KEYWORDS):
                            _lhm_ram_temps.append(sensor)
                            log.debug(f"Found RAM temp sensor: {sensor.Name} on {hw_type}")

        _refresh_lhm_storage_state(refresh_sensor_bindings=True)

        _lhm_available = True
        log.info(
            f"LibreHardwareMonitorLib initialised: CPU sensors found, "
            f"{len(_lhm_ram_temps)} RAM temp sensor(s), "
            f"{len(_lhm_disk_temps)} disk sensor(s)"
        )
        return True
    except Exception as e:
        log.warning(f"LibreHardwareMonitorLib init failed: {e}")
        _lhm_available = False
        return False


def get_lhm_computer():
    """Return (computer, lock) for fan_control to share the LHM instance."""
    return _lhm_computer, _lhm_lock


def set_ui_root(root) -> None:
    """Register the shared Tk UI root used by overlay windows."""
    global _ui_root, _ui_thread_id

    def _drain_ui_tasks() -> None:
        try:
            while True:
                callback = _ui_tasks.get_nowait()
                callback()
        except queue.Empty:
            pass

        if _ui_root is not None:
            _ui_root.after(20, _drain_ui_tasks)

    _ui_root = root
    _ui_thread_id = None if root is None else threading.get_ident()
    if root is not None:
        root.after(20, _drain_ui_tasks)


def _run_on_ui_thread(callback) -> None:
    if _ui_root is None:
        raise RuntimeError("Shared Tk UI root is not available")

    if threading.get_ident() == _ui_thread_id:
        callback()
    else:
        _ui_tasks.put(callback)


def lhm_is_available() -> bool:
    return _lhm_available


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
        "ram_temps":    None,   # list of temperatures for each RAM module
        "ram_temp_c":   None,   # average RAM temperature
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

        # CPU temperature/power and RAM temps via LibreHardwareMonitorLib
        if _lhm_available and _lhm_computer is not None:
            try:
                with _lhm_lock:
                    for hardware in _lhm_computer.Hardware:
                        hw_type = hardware.HardwareType.ToString()
                        if hw_type == "Cpu":
                            hardware.Update()
                        elif hw_type == "Memory":
                            hardware.Update()
                            for sub in hardware.SubHardware:
                                try:
                                    sub.Update()
                                except Exception:
                                    pass
                    cpu_temp  = _lhm_cpu_temp.Value  if _lhm_cpu_temp  is not None else None
                    cpu_power = _lhm_cpu_power.Value if _lhm_cpu_power is not None else None
                    ram_vals  = []
                    for s in _lhm_ram_temps:
                        try:
                            v = s.Value
                            if v is not None:
                                ram_vals.append(float(v))
                        except Exception:
                            pass
                result["cpu_temp_c"]  = cpu_temp
                result["cpu_power_w"] = cpu_power
                if ram_vals:
                    result["ram_temps"] = ram_vals
                    result["ram_temp_c"] = sum(ram_vals) / len(ram_vals)
            except Exception as e:
                log.debug(f"LHM sensor read error: {e}")

    except Exception as e:
        log.error(f"get_system_stats error: {e}", exc_info=True)

    return result


def get_disk_stats() -> dict:
    """Return disk temperature and activity time statistics.
    
    Before reading HDD status, wake them up via I/O to ensure accurate readings.
    Non-HDD drives are skipped.
    """
    result = {"disk_temps": {}, "disk_activity": {}}
    if not (_lhm_available and _lhm_computer is not None):
        return result
    try:
        with _lhm_lock:
            _refresh_lhm_storage_state()
            
            # --- HDD Wake-up Phase ---
            # Before reading sensors, wake up any HDD drives
            try:
                with _disk_wakeup_lock:
                    for drive_letter in _hdd_wakeup_drive_letters:
                        try:
                            if _wake_hdd_via_io(drive_letter):
                                continue
                        except Exception as e:
                            log.debug(f"Error waking HDD {drive_letter}: {e}")
            except Exception as e:
                log.debug(f"HDD wake-up phase error: {e}")
            
            # --- Normal sensor reading phase ---
            disk_temps = {}
            disk_activity = {}
            for disk_name, sensor in _lhm_disk_temps.items():
                try:
                    v = sensor.Value
                    if v is not None:
                        disk_temps[disk_name] = float(v)
                except Exception:
                    pass
            for disk_name, sensor in _lhm_disk_activity.items():
                try:
                    v = sensor.Value
                    if v is not None:
                        disk_activity[disk_name] = float(v)
                except Exception:
                    pass
        result["disk_temps"] = _rename_disk_temp_values(disk_temps)
        result["disk_activity"] = disk_activity
    except Exception as e:
        log.debug(f"get_disk_stats error: {e}")
    return result


def get_fan_stats() -> dict:
    """Return fan speed statistics (RPM values) for all fans with RPM > 0.
    
    Returns all fans that currently have RPM > 0, which indicates they are
    physically connected and spinning.
    """
    result = {"fan_speeds": {}}
    if not (_lhm_available and _lhm_computer is not None):
        return result
    
    try:
        with _lhm_lock:
            fan_list = []
            for hardware in _lhm_computer.Hardware:
                hw_type = hardware.HardwareType.ToString()
                if hw_type != "Motherboard":
                    continue
                try:
                    hardware.Update()
                except Exception:
                    pass
                for sub_hw in hardware.SubHardware:
                    try:
                        sub_hw.Update()
                    except Exception:
                        pass
                    for sensor in sub_hw.Sensors:
                        if sensor.SensorType.ToString() != "Fan":
                            continue
                        try:
                            v = sensor.Value
                            if v is not None:
                                rpm = float(v)
                                # Only include fans with RPM > 0
                                if rpm > 0:
                                    fan_name = sensor.Name
                                    result["fan_speeds"][fan_name] = rpm
                        except Exception:
                            pass
    except Exception as e:
        log.debug(f"get_fan_stats error: {e}")
    return result


def get_monitor_stats() -> dict:
    return {**get_system_stats(), **get_gpu_stats(), **get_disk_stats()}


def _temp_level(temp_c):
    if temp_c is None:
        return "na"
    if temp_c >= 80:
        return "hot"
    if temp_c >= 70:
        return "warm"
    return "normal"


def _level_to_color(level: str) -> str:
    if level == "hot":
        return _FG_HOT
    if level == "warm":
        return _FG_WARM
    if level == "na":
        return _FG_NA
    return _FG_NORMAL


def build_overlay_rows(stats: dict) -> dict:
    cpu_parts = []
    if stats.get("cpu_pct") is not None:
        cpu_parts.append(f"{stats['cpu_pct']:.0f}%")
    if stats.get("cpu_temp_c") is not None:
        cpu_parts.append(f"{stats['cpu_temp_c']:.0f}°C")
    if stats.get("cpu_power_w") is not None:
        cpu_parts.append(f"{stats['cpu_power_w']:.0f}W")

    ram_text = "N/A"
    ram_level = "na"
    if stats.get("ram_used_gb") is not None and stats.get("ram_total_gb") is not None:
        ram_text = f"{stats['ram_used_gb']:.1f}/{stats['ram_total_gb']:.0f}GB"
        if stats.get("ram_temp_c") is not None:
            ram_text += f"  {stats['ram_temp_c']:.0f}°C"
            ram_level = _temp_level(stats.get("ram_temp_c"))
        else:
            ram_level = "normal"

    gpu_parts = []
    if stats.get("gpu_util_pct") is not None:
        gpu_parts.append(f"{stats['gpu_util_pct']}%")
    if stats.get("gpu_temp_c") is not None:
        gpu_parts.append(f"{stats['gpu_temp_c']:.0f}°C")
    if stats.get("gpu_power_w") is not None:
        gpu_parts.append(f"{stats['gpu_power_w']:.0f}W")

    vram_text = "N/A"
    vram_level = "na"
    if stats.get("vram_used_mb") is not None and stats.get("vram_total_mb") is not None:
        vram_text = f"{stats['vram_used_mb'] / 1024:.1f}/{stats['vram_total_mb'] / 1024:.0f}GB"
        vram_level = "normal"

    cpu_level = _temp_level(stats.get("cpu_temp_c")) if stats.get("cpu_temp_c") is not None else ("normal" if cpu_parts else "na")
    gpu_level = _temp_level(stats.get("gpu_temp_c")) if stats.get("gpu_temp_c") is not None else ("normal" if gpu_parts else "na")

    return {
        "cpu": {
            "text": "  ".join(cpu_parts) if cpu_parts else "N/A",
            "level": cpu_level,
        },
        "ram": {
            "text": ram_text,
            "level": ram_level,
        },
        "gpu": {
            "text": "  ".join(gpu_parts) if gpu_parts else "N/A",
            "level": gpu_level,
        },
        "vram": {
            "text": vram_text,
            "level": vram_level,
        },
    }


def get_monitor_snapshot(max_age_ms: int = 500, type: str = "default") -> dict:
    global _snapshot_cache, _snapshot_cache_at

    max_age_s = max(_MIN_SNAPSHOT_CACHE_MS, int(max_age_ms)) / 1000.0
    now = time.monotonic()

    with _snapshot_lock:
        if type == "disk":
            # Lazy: only read disk sensors, no cache
            disk_stats = get_disk_stats()
            return {
                "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
                "disk_temps": disk_stats["disk_temps"],
                "disk_activity": disk_stats["disk_activity"],
            }

        if type == "fan":
            # Lazy: only read fan sensors, no cache
            fan_stats = get_fan_stats()
            return {
                "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
                "fan_speeds": fan_stats["fan_speeds"],
            }

        # Default type: CPU/RAM/GPU only (no disk)
        if (
            _snapshot_cache is not None
            and max_age_s > 0
            and (now - _snapshot_cache_at) <= max_age_s
        ):
            return copy.deepcopy(_snapshot_cache)

        stats = {**get_system_stats(), **get_gpu_stats()}
        snapshot = {
            "timestamp": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
            "sources": {
                "nvml": _nvml_available,
                "lhm": _lhm_available,
            },
            "stats": stats,
        }
        _snapshot_cache = snapshot
        _snapshot_cache_at = now
        return copy.deepcopy(snapshot)


# ---------------------------------------------------------------------------
# Overlay window
# ---------------------------------------------------------------------------

_BG        = "#1a1a1a"
_TITLE_BG  = "#252525"
_FG_NORMAL = "#00e676"
_FG_WARM   = "#ffdd00"
_FG_HOT    = "#ff4444"
_FG_NA     = "#777777"
_FONT      = ("Consolas", 9)
_FONT_BOLD = ("Consolas", 9, "bold")


def _temp_color(temp_c):
    return _level_to_color(_temp_level(temp_c))


def _fmt(val, fmt, unit="", na="N/A"):
    if val is None:
        return na
    return f"{val:{fmt}}{unit}"


class SystemMonitorOverlay:
    """
    Semi-transparent always-on-top overlay.
    Lives on the shared Tk UI thread as a Toplevel window.
    """

    def __init__(self, config: dict, save_config_fn, on_state_change_fn=None):
        self.config        = config
        self.save_config   = save_config_fn
        self._on_state_change_fn = on_state_change_fn
        self._running      = False
        self._fetch_running = False
        self._q: queue.Queue = queue.Queue(maxsize=1)

        # drag state
        self._drag_offset_x = 0
        self._drag_offset_y = 0
        self._closed = False
        self._sync_config_on_close = True
        self._notify_state_on_close = True

        self.root   = None
        self._labels = {}  # key -> tk.Label

    def _notify_state_change(self, enabled: bool) -> None:
        if self._on_state_change_fn:
            try:
                self._on_state_change_fn(enabled)
            except Exception:
                pass

    # -----------------------------------------------------------------------
    # Public API (called from other threads)
    # -----------------------------------------------------------------------

    def show(self, parent) -> None:
        """Build the overlay window on the shared Tk UI thread."""
        global _overlay_instance
        self._running = True
        self.root = tk.Toplevel(parent)
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", self.config["overlay"]["opacity"])
        self.root.configure(bg=_BG)
        self.root.resizable(False, False)
        self.root.bind("<Destroy>", self._on_destroy, add="+")

        self._build_ui()
        self._position_window()
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        _set_overlay_enabled_in_config(self.config, self.save_config, True)
        self._notify_state_change(True)

        # Kick off stats loop
        self.root.after(100, self._update_stats)

    def close(self, sync_config: bool = True, notify_state: bool = True) -> None:
        """Destroy the window (safe to call from any thread)."""
        self._sync_config_on_close = sync_config
        self._notify_state_on_close = notify_state

        def _close_impl():
            self._finalize_close()
            if self.root is not None:
                try:
                    if self.root.winfo_exists():
                        self.root.destroy()
                except Exception:
                    pass

        if self.root is not None:
            try:
                _run_on_ui_thread(_close_impl)
            except Exception:
                pass
        else:
            self._finalize_close()

    def _finalize_close(self) -> None:
        global _overlay_instance

        if self._closed:
            return

        self._closed = True
        self._running = False
        if _overlay_instance is self:
            _overlay_instance = None
        if self._sync_config_on_close:
            _set_overlay_enabled_in_config(self.config, self.save_config, False)
        if self._notify_state_on_close:
            self._notify_state_change(False)

    def _on_destroy(self, event) -> None:
        if event.widget is self.root:
            self._finalize_close()

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

        try:
            # Check if the underlying window still exists (OS may have destroyed it)
            if not self.root.winfo_exists():
                log.warning("Overlay window was destroyed externally, cleaning up")
                self.close()
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

            # Re-assert topmost every 30 cycles (~30s at 1000ms refresh) to
            # counteract Windows DWM occasionally dropping the flag.
            self._topmost_counter = getattr(self, "_topmost_counter", 0) + 1
            if self._topmost_counter >= 30:
                self._topmost_counter = 0
                self.root.attributes("-topmost", False)
                self.root.attributes("-topmost", True)

        except Exception:
            log.exception("Error in overlay update loop")

        # Always reschedule as long as we're running, even if an error occurred
        if self._running:
            refresh = self.config["overlay"].get("refresh_ms", 1000)
            self.root.after(refresh, self._update_stats)

    def _fetch_thread(self) -> None:
        try:
            snapshot = get_monitor_snapshot()
            try:
                self._q.put_nowait(snapshot)
            except queue.Full:
                pass
        finally:
            self._fetch_running = False

    def _apply_stats(self, snapshot: dict) -> None:
        """Update label text and colours from a monitor snapshot."""
        rows = snapshot.get("overlay") or build_overlay_rows(snapshot.get("stats", snapshot))
        for key, label in self._labels.items():
            row = rows.get(key, {"text": "N/A", "level": "na"})
            self._set(label, row["text"], _level_to_color(row["level"]))

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


def set_overlay_enabled(
    config: dict,
    save_config_fn,
    enabled: bool,
    on_state_change_fn=None,
    persist_config: bool = True,
) -> None:
    """Apply the requested overlay enabled state and keep runtime/UI state aligned."""
    desired = bool(enabled)

    def _sync_impl():
        global _overlay_instance
        config_changed = False

        if persist_config:
            config_changed = _set_overlay_enabled_in_config(config, save_config_fn, desired)

        if (_overlay_instance is not None) == desired:
            if config_changed and on_state_change_fn:
                try:
                    on_state_change_fn(desired)
                except Exception:
                    pass
            return

        if desired:
            instance = SystemMonitorOverlay(
                config,
                save_config_fn,
                on_state_change_fn=on_state_change_fn,
            )
            _overlay_instance = instance
            try:
                instance.show(_ui_root)
            except Exception:
                if _overlay_instance is instance:
                    _overlay_instance = None
                _set_overlay_enabled_in_config(config, save_config_fn, False)
                if on_state_change_fn:
                    try:
                        on_state_change_fn(False)
                    except Exception:
                        pass
                raise
            return

        _overlay_instance.close(sync_config=persist_config, notify_state=True)

    try:
        _run_on_ui_thread(_sync_impl)
    except Exception as e:
        log.error(f"Error setting overlay state: {e}", exc_info=True)


def toggle_overlay(config: dict, save_config_fn, on_state_change_fn=None) -> None:
    """Toggle the overlay based on the actual runtime window state."""
    set_overlay_enabled(
        config,
        save_config_fn,
        not overlay_is_open(),
        on_state_change_fn=on_state_change_fn,
    )


def close_overlay() -> None:
    """Close overlay if open (called during shutdown)."""
    def _close_impl():
        if _overlay_instance is not None:
            _overlay_instance.close(sync_config=False, notify_state=False)

    try:
        _run_on_ui_thread(_close_impl)
    except Exception:
        if _overlay_instance is not None:
            _overlay_instance.close(sync_config=False, notify_state=False)


def overlay_is_open() -> bool:
    return _overlay_instance is not None


def apply_overlay_opacity(opacity: float) -> None:
    """Apply opacity to the running overlay immediately (no restart needed)."""
    def _apply_impl():
        if _overlay_instance is not None and _overlay_instance.root is not None:
            try:
                _overlay_instance.root.attributes("-alpha", opacity)
            except Exception:
                pass

    try:
        _run_on_ui_thread(_apply_impl)
    except Exception:
        pass

