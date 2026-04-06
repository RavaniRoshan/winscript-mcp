import re
from typing import Optional

try:
    import pywinauto
    from pywinauto import Desktop
except ImportError:
    pywinauto = None
    Desktop = None

APP_ALIASES = {
    "notepad": "notepad.exe",
    "chrome": "chrome.exe",
    "firefox": "firefox.exe",
    "edge": "msedge.exe",
    "excel": "EXCEL.EXE",
    "word": "WINWORD.EXE",
    "outlook": "OUTLOOK.EXE",
    "explorer": "explorer.exe",
    "cmd": "cmd.exe",
    "powershell": "powershell.exe",
    "terminal": "wt.exe",
    "vscode": "Code.exe",
    "cursor": "Cursor.exe",
}

def resolve_app_exe(name: str) -> str:
    return APP_ALIASES.get(name.lower().strip(), name)

def find_window(title_hint: str) -> Optional[any]:
    """
    Find window by partial title. Three strategies in order:
    1. Exact match
    2. Case-insensitive contains
    3. Regex match
    Returns first match or None.
    """
    if Desktop is None:
        raise ImportError("pywinauto is not available")
        
    desktop = Desktop(backend="uia")
    windows = desktop.windows()

    for w in windows:
        try:
            if w.window_text() == title_hint:
                return w
        except Exception:
            continue

    for w in windows:
        try:
            if title_hint.lower() in w.window_text().lower():
                return w
        except Exception:
            continue

    try:
        pattern = re.compile(title_hint, re.IGNORECASE)
        for w in windows:
            try:
                if pattern.search(w.window_text()):
                    return w
            except Exception:
                continue
    except re.error:
        pass

    return None

def get_all_windows() -> list[dict]:
    if Desktop is None:
        raise ImportError("pywinauto is not available")
        
    desktop = Desktop(backend="uia")
    results = []
    for w in desktop.windows():
        try:
            title = w.window_text()
            pid = w.process_id()
            if title.strip():
                results.append({"title": title, "pid": pid})
        except Exception:
            continue
    return results

def connect_to_process(exe_name: str) -> Optional[any]:
    if pywinauto is None:
        raise ImportError("pywinauto is not available")
        
    try:
        return pywinauto.Application(backend="uia").connect(path=exe_name)
    except Exception:
        return None