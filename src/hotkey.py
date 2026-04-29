"""
Little Helper - Low-level keyboard hook (WH_KEYBOARD_LL, no admin required).
"""

import ctypes
import ctypes.wintypes
import threading
import logging

from clipboard_paste import should_skip_paste

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
_hook_thread = None  # reference to the hook thread for clean shutdown
_stop_event  = threading.Event()  # signal to stop the message loop

_activity_callbacks = []  # callbacks for user activity (called on any keydown)

import time as _time
_last_paste_t      = 0.0
_last_screenshot_t = 0.0
_HOTKEY_COOLDOWN   = 0.5  # seconds


def _is_key_down(vk_code: int) -> bool:
    return bool(_user32.GetAsyncKeyState(vk_code) & 0x8000)


def register_activity_callback(fn) -> None:
    """Register a callback to be called on any keyboard activity (WM_KEYDOWN)."""
    if fn not in _activity_callbacks:
        _activity_callbacks.append(fn)


def start_keyboard_hook(config: dict, on_paste_fn, on_screenshot_fn) -> None:
    """
    Install WH_KEYBOARD_LL hook and run the message loop (blocks until stopped).
    config, on_paste_fn, on_screenshot_fn are captured via closure.
    Should be called from a daemon thread to allow clean shutdown.
    """
    global _hook_handle, _hook_proc, _hook_thread, _stop_event

    _hook_thread = threading.current_thread()
    _stop_event.clear()

    def _proc(nCode, wParam, lParam):
        try:
            if nCode == HC_ACTION:
                kb  = ctypes.cast(lParam, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                vk  = kb.vkCode

                # Call activity callbacks on any WM_KEYDOWN
                if wParam == WM_KEYDOWN:
                    for callback in _activity_callbacks:
                        try:
                            callback()
                        except Exception as e:
                            log.error(f"Error in activity callback: {e}", exc_info=True)

                ctrl_down = _is_key_down(0x11)   # VK_CONTROL
                alt_down  = _is_key_down(0x12)   # VK_MENU

                paste_key      = VK_CODES.get(config["paste_hotkey"]["key"].upper(), VK_V)
                paste_mod      = config["paste_hotkey"]["modifier"].lower()
                screenshot_key = VK_CODES.get(config["screenshot_hotkey"]["key"].upper(), VK_A)
                screenshot_mod = config["screenshot_hotkey"]["modifier"].lower()

                paste_mod_ok = (ctrl_down and paste_mod == "ctrl") or \
                               (alt_down  and paste_mod == "alt")
                if vk == paste_key and wParam == WM_KEYDOWN and paste_mod_ok:
                    global _last_paste_t
                    now = _time.monotonic()
                    if now - _last_paste_t >= _HOTKEY_COOLDOWN:
                        _last_paste_t = now
                        hotkey_str = f"{paste_mod.capitalize()}+{config['paste_hotkey']['key'].upper()}"
                        
                        # Check if we should skip paste in editable contexts
                        if should_skip_paste():
                            log.debug(f"{hotkey_str} detected but skipping - in editable context")
                        else:
                            log.debug(f"{hotkey_str} detected via hook")
                            threading.Thread(target=on_paste_fn, daemon=True).start()

                # Accept SYSKEYDOWN too so Alt+X fires when another window has focus
                screenshot_mod_ok = (ctrl_down and screenshot_mod == "ctrl") or \
                                    (alt_down  and screenshot_mod == "alt")
                if vk == screenshot_key and \
                   wParam in (WM_KEYDOWN, WM_SYSKEYDOWN) and \
                   screenshot_mod_ok:
                    global _last_screenshot_t
                    now = _time.monotonic()
                    if now - _last_screenshot_t >= _HOTKEY_COOLDOWN:
                        _last_screenshot_t = now
                        hotkey_str = f"{screenshot_mod.capitalize()}+{config['screenshot_hotkey']['key'].upper()}"
                        log.debug(f"{hotkey_str} detected via hook")
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
            # Check stop event to exit early (e.g., if UnhookWindowsHookEx was called)
            if _stop_event.is_set():
                log.debug("Stop event detected in message loop")
                break
            _user32.TranslateMessage(ctypes.byref(msg))
            _user32.DispatchMessageW(ctypes.byref(msg))

        log.info("Keyboard hook message loop ended")
    except Exception as e:
        log.error(f"Exception in start_keyboard_hook: {e}", exc_info=True)
    finally:
        # Ensure hook is unhooked
        if _hook_handle:
            try:
                _user32.UnhookWindowsHookEx(_hook_handle)
            except Exception as e:
                log.error(f"Error unhoking in finally: {e}")
        _hook_handle = None
        _hook_proc = None
        _hook_thread = None


def stop_keyboard_hook() -> None:
    """Remove the keyboard hook and stop the message loop."""
    global _hook_handle, _stop_event, _hook_thread
    
    # Signal the loop to exit
    _stop_event.set()
    
    if _hook_handle:
        log.info("Removing keyboard hook...")
        try:
            _user32.UnhookWindowsHookEx(_hook_handle)
        except Exception as e:
            log.error(f"Error unhoking: {e}")
        _hook_handle = None
    
    # Post a dummy message to wake up GetMessageW so it can check _stop_event
    if _hook_thread is not None:
        try:
            if _hook_thread.is_alive():
                _user32.PostThreadMessageW(_hook_thread.ident, 0x0012, 0, 0)
        except Exception:
            pass  # Thread may not have a message queue; ignore
    
    # Wait for the hook thread to finish (max 2 seconds)
    if _hook_thread is not None:
        try:
            if _hook_thread.is_alive():
                _hook_thread.join(timeout=2)
                if _hook_thread.is_alive():
                    log.warning("Hook thread did not stop within timeout")
        except Exception:
            pass  # Thread already gone or in bad state
    
    _hook_thread = None
    log.info("Keyboard hook stopped")
