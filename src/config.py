"""
Little Helper - Configuration management and path utilities.
"""

import os
import sys
import json
import logging

log = logging.getLogger("little_helper.config")

DEFAULT_CONFIG = {
    "paste_hotkey": {"modifier": "ctrl", "key": "V"},
    "screenshot_hotkey": {"modifier": "alt", "key": "A"},
    "gpu_power_limit": {
        "enabled": False,
        "watts": 150,
    },
    "overlay": {
        "enabled": False,   # auto-show on startup
        "x": -1,            # -1 = auto top-right corner
        "y": -1,
        "width": 220,
        "height": 200,
        "opacity": 0.85,
        "refresh_ms": 1000,
    },
}


def get_data_dir() -> str:
    """Directory for config/log files (exe dir when frozen, project root otherwise)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    # src/ is one level below project root
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def get_resource_path(filename: str) -> str:
    """Path to a bundled resource; works in both dev and PyInstaller."""
    if getattr(sys, "frozen", False):
        return os.path.join(sys._MEIPASS, filename)
    # dev: resources live in res/ at project root
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(project_root, "res", filename)


def get_config_path() -> str:
    return os.path.join(get_data_dir(), "config.json")


def get_log_path() -> str:
    return os.path.join(get_data_dir(), "little_helper.log")


def load_config() -> dict:
    """Load config from file; merge missing keys from DEFAULT_CONFIG."""
    config_path = get_config_path()
    try:
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            # Deep-merge top-level dict keys only (sufficient for our schema)
            for key, default_val in DEFAULT_CONFIG.items():
                if key not in config:
                    config[key] = default_val
                elif isinstance(default_val, dict):
                    for sub_key, sub_val in default_val.items():
                        if sub_key not in config[key]:
                            config[key][sub_key] = sub_val
            return config
    except Exception as e:
        log.error(f"Error loading config: {e}")
    return DEFAULT_CONFIG.copy()


def save_config(config: dict) -> None:
    """Persist config to disk."""
    config_path = get_config_path()
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
        log.info(f"Config saved to {config_path}")
    except Exception as e:
        log.error(f"Error saving config: {e}")
