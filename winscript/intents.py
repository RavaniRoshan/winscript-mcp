import os, glob
from pathlib import Path
from winscript.tools.app_control import open_app
from winscript.tools.filesystem import list_dir, read_file_text
from winscript.tools.screen import get_clipboard, set_clipboard
from winscript.tools.ui_interaction import type_text, press_key
from winscript.adapters.outlook_adapter import outlook
import subprocess, time

def open_latest_file(folder: str, extension: str = "*") -> str:
    """
    Find and open the most recently modified file in a folder.
    extension: filter by type e.g. "xlsx", "pdf", "txt". Default: any.
    Example: open_latest_file("C:/reports", "xlsx")
    Opens the file in its default application.
    """
    try:
        pattern = f"*.{extension}" if extension != "*" else "*"
        files = list(Path(folder).glob(pattern))
        if not files:
            return f"ERROR: No {extension} files found in {folder}"
        files = [f for f in files if f.is_file()]
        if not files:
            return f"ERROR: No {extension} files found in {folder}"
        latest = max(files, key=lambda f: f.stat().st_mtime)
        subprocess.Popen(["start", "", str(latest)], shell=True)
        time.sleep(1.5)
        from winscript.core.memory import remember_file
        remember_file(str(latest))
        return f"Opened latest file: {latest.name} (modified: {time.ctime(latest.stat().st_mtime)})"
    except Exception as e:
        return f"ERROR: {str(e)}"

def send_email_with_content(to: str, subject: str, content_source: str) -> str:
    """
    Send an email. content_source can be:
    - A file path: reads the file and uses it as body
    - A clipboard keyword "clipboard": uses current clipboard text
    - Any other string: used directly as body
    Example: send_email_with_content("a@b.com","Report","C:/report.txt")
    Example: send_email_with_content("a@b.com","Update","clipboard")
    """
    try:
        if content_source == "clipboard":
            body = get_clipboard()
        elif Path(content_source).exists():
            body = read_file_text(content_source, max_chars=5000)
        else:
            body = content_source
        return outlook.send(to, subject, body)
    except Exception as e:
        return f"ERROR: {str(e)}"

def find_in_folder(folder: str, search_term: str, extension: str = "*") -> str:
    """
    Find files whose name contains search_term in a folder.
    Extension filter optional. Returns matching file paths.
    Example: find_in_folder("C:/docs","invoice","pdf")
    """
    try:
        pattern = f"*.{extension}" if extension != "*" else "*"
        files = list(Path(folder).glob(pattern))
        matches = [f for f in files if search_term.lower() in f.name.lower() and f.is_file()]
        if not matches:
            return f"No files matching '{search_term}' in {folder}"
        lines = [str(f) for f in sorted(matches, key=lambda f: f.stat().st_mtime, reverse=True)]
        return "\n".join(lines)
    except Exception as e:
        return f"ERROR: {str(e)}"

def read_active_document() -> str:
    """
    Read text from the currently active window's document.
    Tries: clipboard select-all copy, then returns text.
    Works best with text editors, Notepad, Word.
    Example: read_active_document()
    """
    try:
        # Select all + copy to clipboard
        press_key("ctrl+a")
        time.sleep(0.2)
        press_key("ctrl+c")
        time.sleep(0.3)
        text = get_clipboard()
        return text if text else "ERROR: Could not read document content"
    except Exception as e:
        return f"ERROR: {str(e)}"

def summarize_screen() -> str:
    """
    Take a screenshot and return it as base64 PNG.
    The agent can then use vision to describe what is on screen.
    Returns: base64 PNG data URI for Claude vision.
    Example: summarize_screen()
    """
    from winscript.tools.screen import take_screenshot, get_active_window
    active = get_active_window()
    shot = take_screenshot()
    return f"Active window: {active}\nScreenshot: {shot}"
