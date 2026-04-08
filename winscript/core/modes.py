from enum import Enum

class ExecutionMode(Enum):
    STANDARD = "standard"
    SAFE = "safe"

_current_mode = ExecutionMode.STANDARD

BLOCKED_IN_SAFE_MODE = {
    "write_file_text", "delete_file", "move_file", "copy_file",
    "type_text", "click", "press_key", "outlook_send_email",
    "excel_write_cell", "open_app", "close_app", "workflow_replay",
    "notepad_type", "notepad_save", "notepad_close",
    "chrome_navigate", "chrome_new_tab", "chrome_close_tab",
    "set_clipboard",
}

def get_mode() -> ExecutionMode:
    return _current_mode

def set_mode(mode: ExecutionMode):
    global _current_mode
    _current_mode = mode

def is_blocked(tool_name: str) -> bool:
    return _current_mode == ExecutionMode.SAFE and tool_name in BLOCKED_IN_SAFE_MODE
