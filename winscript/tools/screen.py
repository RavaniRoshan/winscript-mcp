import base64
import io

try:
    import mss
except ImportError:
    mss = None

from PIL import Image

try:
    import win32gui
    import win32clipboard
    import win32con
except ImportError:
    win32gui = None
    win32clipboard = None
    win32con = None

from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError

def take_screenshot(region: dict = None) -> str:
    """Capture screen and return as base64 PNG string.
    region: optional {"top":0,"left":0,"width":1920,"height":1080}
    If no region: captures full primary monitor.
    Returns: "data:image/png;base64,<b64string>"
    Example: take_screenshot() or take_screenshot({"top":0,"left":0,"width":800,"height":600})"""
    args = {"region": str(region)}
    try:
        if mss is None:
            raise ImportError("mss not available")
        with mss.mss() as sct:
            monitor = region if region else sct.monitors[1]
            shot = sct.grab(monitor)
            img = Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        guard.record_success("take_screenshot", args)
        return f"data:image/png;base64,{b64}"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("take_screenshot", args)
        return f"ERROR taking screenshot: {str(e)}"

def get_active_window() -> str:
    """Get the title of the currently focused window.
    Example: get_active_window()"""
    try:
        if win32gui is None:
            raise ImportError("win32gui not available")
        hwnd = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(hwnd)
        return title if title else "(no active window)"
    except Exception as e:
        return f"ERROR: {str(e)}"

def get_clipboard() -> str:
    """Read current clipboard text content.
    Example: get_clipboard()"""
    try:
        if win32clipboard is None:
            raise ImportError("win32clipboard not available")
        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            return data
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        return f"ERROR reading clipboard: {str(e)}"

def set_clipboard(text: str) -> str:
    """Set clipboard to given text.
    Example: set_clipboard("Hello world")"""
    try:
        if win32clipboard is None:
            raise ImportError("win32clipboard not available")
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
        return f"Clipboard set ({len(text)} chars)"
    except Exception as e:
        return f"ERROR setting clipboard: {str(e)}"