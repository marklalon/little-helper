"""
Little Helper - Clipboard image paste and Explorer path detection.
"""

import os
import logging
import urllib.parse
from datetime import datetime
from io import BytesIO

import pythoncom
import win32gui
import win32com.client
import win32clipboard
import win32con
from PIL import Image, ImageGrab

log = logging.getLogger("little_helper.clipboard_paste")


def get_explorer_path() -> str | None:
    """Get the directory path of the currently focused Explorer window."""
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        log.debug("No foreground window")
        return None

    class_name = win32gui.GetClassName(hwnd)
    log.debug(f"Foreground window class: {class_name}, hwnd: {hwnd}")

    # Desktop window -> use Desktop path
    if class_name in ("WorkerW", "Progman"):
        desktop = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
        if os.path.isdir(desktop):
            log.debug(f"Desktop detected, path: {desktop}")
            return desktop
        log.debug("Desktop detected but path not found")
        return None

    if class_name not in ("CabinetWClass", "ExploreWClass"):
        log.debug("Not an Explorer window, skipping")
        return None

    try:
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
        log.debug(f"Shell.Windows count: {windows.Count}")

        for i in range(windows.Count):
            try:
                window = windows.Item(i)
                if window is None:
                    continue
                if window.HWND == hwnd:
                    url = window.LocationURL
                    log.debug(f"Matched window, LocationURL: {url}")
                    if url and url.startswith("file:///"):
                        path = urllib.parse.unquote(url[8:]).replace("/", "\\")
                        if os.path.isdir(path):
                            return path
                    folder_path = window.Document.Folder.Self.Path
                    log.debug(f"Fallback folder path: {folder_path}")
                    if os.path.isdir(folder_path):
                        return folder_path
            except Exception as e:
                log.debug(f"Error checking window {i}: {e}")
                continue
    except Exception as e:
        log.error(f"COM error in get_explorer_path: {e}")
    finally:
        pythoncom.CoUninitialize()

    return None


def generate_filename(directory: str) -> str:
    """Generate a unique clipboard-*.png filename."""
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    base = f"clipboard-{timestamp}"
    path = os.path.join(directory, f"{base}.png")
    if not os.path.exists(path):
        return path
    counter = 1
    while True:
        path = os.path.join(directory, f"{base}-{counter}.png")
        if not os.path.exists(path):
            return path
        counter += 1


def get_clipboard_image() -> Image.Image | None:
    """Return PIL Image from clipboard, or None."""
    img = ImageGrab.grabclipboard()
    if isinstance(img, Image.Image):
        log.debug("Got image via ImageGrab.grabclipboard()")
        return img
    if isinstance(img, list):
        for path in img:
            try:
                return Image.open(path)
            except Exception:
                pass
    return None


def copy_image_to_clipboard(img: Image.Image) -> None:
    """Copy a PIL Image to the Windows clipboard as CF_DIB."""
    if img.mode != "RGB":
        img = img.convert("RGB")

    output = BytesIO()
    try:
        img.save(output, "BMP")
        dib_data = output.getvalue()[14:]
        log.debug(f"Prepared CF_DIB payload: {len(dib_data)} bytes")
    finally:
        output.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, dib_data)
        log.info("Image copied to clipboard")
    finally:
        win32clipboard.CloseClipboard()


def on_paste(config: dict, notify_fn=None) -> None:
    """Handle paste hotkey: save clipboard image to active Explorer directory."""
    log.debug("on_paste triggered")
    try:
        img = get_clipboard_image()
        if img is None:
            log.debug("Clipboard does not contain an image, passing through")
            return

        target_dir = get_explorer_path()
        log.debug(f"Target directory: {target_dir}")
        if not target_dir:
            return

        filepath = generate_filename(target_dir)
        img.save(filepath, "PNG")
        log.info(f"Saved clipboard image to: {filepath}")

    except Exception as e:
        log.error(f"Error in on_paste: {e}", exc_info=True)
