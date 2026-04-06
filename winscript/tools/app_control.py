import subprocess
import time
from winscript.core.window_resolver import resolve_app_exe, find_window, get_all_windows
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError

def open_app(name: str, wait_seconds: float = 2.0) -> str:
    """Open a Windows app by name or alias. wait_seconds: pause after launch."""
    args = {"name": name}
    try:
        exe = resolve_app_exe(name)
        subprocess.Popen(exe)
        time.sleep(wait_seconds)
        guard.record_success("open_app", args)
        return f"Opened {name} ({exe})"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("open_app", args)
        return f"ERROR opening {name}: {str(e)}"

def close_app(title_hint: str) -> str:
    """Close a window by title (partial match)."""
    args = {"title_hint": title_hint}
    try:
        window = find_window(title_hint)
        if window is None:
            guard.record_failure("close_app", args)
            return f"ERROR: No window found matching '{title_hint}'"
        window.close()
        guard.record_success("close_app", args)
        return f"Closed window: {window.window_text()}"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("close_app", args)
        return f"ERROR closing '{title_hint}': {str(e)}"

def focus_app(title_hint: str) -> str:
    """Bring a window to the foreground by title (partial match)."""
    args = {"title_hint": title_hint}
    try:
        window = find_window(title_hint)
        if window is None:
            guard.record_failure("focus_app", args)
            return f"ERROR: No window found matching '{title_hint}'"
        window.set_focus()
        guard.record_success("focus_app", args)
        return f"Focused: {window.window_text()}"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("focus_app", args)
        return f"ERROR focusing '{title_hint}': {str(e)}"

def get_running_apps() -> str:
    """List all open windows with titles and PIDs."""
    try:
        windows = get_all_windows()
        if not windows:
            return "No windows found."
        return "\n".join(f"[PID {w['pid']}] {w['title']}" for w in windows)
    except Exception as e:
        return f"ERROR listing windows: {str(e)}"