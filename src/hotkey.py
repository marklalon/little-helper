"""
Little Helper - Low-level keyboard hook (WH_KEYBOARD_LL, no admin required).
"""

import ctypes
import ctypes.wintypes
import threading
import logging

log = logging.getLogger("little_helper.hotkey")

# --- Win32 constants ---
WH_KEYBOARD_LL = 13
WM_KEYDOWN     = 0x0100
WM_KEYUP       = 0x0101
WM_SYSKEYDOWN  = 0x0104
VK_V           = 0x56
VK_A           = 0x41
HC_ACTION      = 0

# --- DLL references ---
_user32   = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32

# Fix 64-bit pointer types
_kernel32.GetModuleHandleW.restype  = ctypes.wintypes.HMODULE
_kernel32.GetModuleHandleW.argtypes = [ctypes.wintypes.LPCWSTR]

_user32.SetWindowsHookExW.restype  = ctypes.c_void_p
_user32.SetWindowsHookExW.argtypes = [
    ctypes.c_int,
    ctypes.c_void_p,
    ctypes.wintypes.HINSTANCE,
    ctypes.wintypes.DWORD,
]

_user32.CallNextHookEx.restype  = ctypes.c_long
_user32.CallNextHookEx.argtypes = [
    ctypes.c_void_p,
    ctypes.c_int,
    ctypes.wintypes.WPARAM,
    ctypes.wintypes.LPARAM,
]

_user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]

_user32.GetAsyncKeyState.restype  = ctypes.c_short
_user32.GetAsyncKeyState.argtypes = [ctypes.c_int]

# PostMessageW signature (also used by single-instance code in main)
_user32.PostMessageW.argtypes = [
    ctypes.wintypes.HWND,
    ctypes.wintypes.UINT,
    ctypes.wintypes.WPARAM,
    ctypes.wintypes.LPARAM,
]

HOOKPROC = ctypes.CFUNCTYPE(
    ctypes.wintypes.LPARAM,
    ctypes.c_int,
    ctypes.wintypes.WPARAM,
    ctypes.wintypes.LPARAM,
)


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode",      ctypes.wintypes.DWORD),
        ("scanCode",    ctypes.wintypes.DWORD),
        ("flags",       ctypes.wintypes.DWORD),
        ("time",        ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.wintypes.ULONG)),
    ]


# VK code mapping for common keys
VK_CODES = {
    "A": 0x41, "B": 0x42, "C": 0x43, "D": 0x44, "E": 0x45,
    "F": 0x46, "G": 0x47, "H": 0x48, "I": 0x49, "J": 0x4A,
    "K": 0x4B, "L": 0x4C, "M": 0x4D, "N": 0x4E, "O": 0x4F,
    "P": 0x50, "Q": 0x51, "R": 0x52, "S": 0x53, "T": 0x54,
    "U": 0x55, "V": 0x56, "W": 0x57, "X": 0x58, "Y": 0x59, "Z": 0x5A,
    "0": 0x30, "1": 0x31, "2": 0x32, "3": 0x33, "4": 0x34,
    "5": 0x35, "6": 0x36, "7": 0x37, "8": 0x38, "9": 0x39,
    "F1": 0x70,  "F2": 0x71,  "F3": 0x72,  "F4": 0x73,
    "F5": 0x74,  "F6": 0x75,  "F7": 0x76,  "F8": 0x77,
    "F9": 0x78,  "F10": 0x79, "F11": 0x7A, "F12": 0x7B,
}

# Module-level state
_hook_handle = None
_hook_proc   = None  # kept alive to prevent GC


def _is_key_down(vk_code: int) -> bool:
    return bool(_user32.GetAsyncKeyState(vk_code) & 0x8000)


def start_keyboard_hook(config: dict, on_paste_fn, on_screenshot_fn) -> None:
    """
    Install WH_KEYBOARD_LL hook and run the message loop (blocks until stopped).
    config, on_paste_fn, on_screenshot_fn are captured via closure.
    """
    global _hook_handle, _hook_proc

    def _proc(nCode, wParam, lParam):
        try:
            if nCode == HC_ACTION:
                kb  = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vk  = kb.vkCode

                ctrl_down = _is_key_down(0x11)   # VK_CONTROL
                alt_down  = _is_key_down(0x12)   # VK_MENU

                paste_key      = VK_CODES.get(config["paste_hotkey"]["key"].upper(), VK_V)
                paste_mod      = config["paste_hotkey"]["modifier"].lower()
                screenshot_key = VK_CODES.get(config["screenshot_hotkey"]["key"].upper(), VK_A)
                screenshot_mod = config["screenshot_hotkey"]["modifier"].lower()

                paste_mod_ok = (ctrl_down and paste_mod == "ctrl") or \
                               (alt_down  and paste_mod == "alt")
                if vk == paste_key and wParam == WM_KEYDOWN and paste_mod_ok:
                    hotkey_str = f"{paste_mod.capitalize()}+{config['paste_hotkey']['key'].upper()}"
                    log.info(f"{hotkey_str} detected via hook!")
                    threading.Thread(target=on_paste_fn, daemon=True).start()

                # Accept SYSKEYDOWN too so Alt+X fires when another window has focus
                screenshot_mod_ok = (ctrl_down and screenshot_mod == "ctrl") or \
                                    (alt_down  and screenshot_mod == "alt")
                if vk == screenshot_key and \
                   wParam in (WM_KEYDOWN, WM_SYSKEYDOWN) and \
                   screenshot_mod_ok:
                    hotkey_str = f"{screenshot_mod.capitalize()}+{config['screenshot_hotkey']['key'].upper()}"
                    log.info(f"{hotkey_str} detected via hook!")
                    threading.Thread(target=on_screenshot_fn, daemon=True).start()

        except Exception as e:
            log.error(f"Exception in hook callback: {e}", exc_info=True)

        return _user32.CallNextHookEx(_hook_handle, nCode, wParam, lParam)

    _hook_proc = HOOKPROC(_proc)

    try:
        log.info("Installing keyboard hook...")
        hmod = _kernel32.GetModuleHandleW("user32.dll")
        _hook_handle = _user32.SetWindowsHookExW(
            WH_KEYBOARD_LL,
            ctypes.cast(_hook_proc, ctypes.c_void_p).value,
            hmod,
            0,
        )
        if not _hook_handle:
            err = ctypes.GetLastError()
            log.error(f"Failed to install hook, error: {err}")
            return

        log.info(f"Keyboard hook installed (handle={_hook_handle})")

        msg = ctypes.wintypes.MSG()
        while _user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            _user32.TranslateMessage(ctypes.byref(msg))
            _user32.DispatchMessageW(ctypes.byref(msg))

        log.info("Keyboard hook message loop ended")
    except Exception as e:
        log.error(f"Exception in start_keyboard_hook: {e}", exc_info=True)


def stop_keyboard_hook() -> None:
    """Remove the keyboard hook."""
    global _hook_handle
    if _hook_handle:
        _user32.UnhookWindowsHookEx(_hook_handle)
        _hook_handle = None
        log.info("Keyboard hook removed")
