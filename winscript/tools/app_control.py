import subprocess
import time
from winscript.core.window_resolver import resolve_app_exe, find_window, get_all_windows
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError
from winscript.core.state_diff import StateCapture

def open_app(name: str, wait_seconds: float = 2.0) -> str:
    """Open a Windows app by name or alias. wait_seconds: pause after launch."""
    args = {"name": name}
    with StateCapture("open_app", args) as capture:
        try:
            exe = resolve_app_exe(name)
            subprocess.Popen(exe)
            time.sleep(wait_seconds)
            guard.record_success("open_app", args)
            from winscript.core.memory import remember_window
            remember_window(name)
            result = f"Opened {name} ({exe})"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("open_app", args)
            result = f"ERROR opening {name}: {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"

def close_app(title_hint: str) -> str:
    """Close a window by title (partial match)."""
    args = {"title_hint": title_hint}
    with StateCapture("close_app", args) as capture:
        try:
            window = find_window(title_hint)
            if window is None:
                guard.record_failure("close_app", args)
                result = f"ERROR: No window found matching '{title_hint}'"
            else:
                window.close()
                guard.record_success("close_app", args)
                result = f"Closed window: {window.window_text()}"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("close_app", args)
            result = f"ERROR closing '{title_hint}': {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"

def focus_app(title_hint: str) -> str:
    """Bring a window to the foreground by title (partial match)."""
    args = {"title_hint": title_hint}
    with StateCapture("focus_app", args) as capture:
        try:
            window = find_window(title_hint)
            if window is None:
                guard.record_failure("focus_app", args)
                result = f"ERROR: No window found matching '{title_hint}'"
            else:
                window.set_focus()
                guard.record_success("focus_app", args)
                from winscript.core.memory import remember_window
                remember_window(window.window_text())
                result = f"Focused: {window.window_text()}"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("focus_app", args)
            result = f"ERROR focusing '{title_hint}': {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"

def get_running_apps() -> str:
    """List all open windows with titles and PIDs."""
    try:
        windows = get_all_windows()
        if not windows:
            return "No windows found."
        return "\n".join(f"[PID {w['pid']}] {w['title']}" for w in windows)
    except Exception as e:
        return f"ERROR listing windows: {str(e)}"

def wait_for_window(title_hint: str, timeout_seconds: int = 10) -> str:
    """Wait for a window to appear. Important for slow-loading applications.
    Example: wait_for_window("Notepad", 5)"""
    args = {"title_hint": title_hint, "timeout_seconds": timeout_seconds}
    try:
        start_time = time.time()
        while time.time() - start_time < timeout_seconds:
            window = find_window(title_hint)
            if window is not None:
                guard.record_success("wait_for_window", args)
                return f"Window appeared: {window.window_text()}"
            time.sleep(0.5)
        
        guard.record_failure("wait_for_window", args)
        return f"ERROR: Window '{title_hint}' did not appear within {timeout_seconds}s"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("wait_for_window", args)
        return f"ERROR waiting for window: {str(e)}"