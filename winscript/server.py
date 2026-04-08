import subprocess
from fastmcp import FastMCP

from winscript.tools.app_control import open_app, close_app, focus_app, get_running_apps, wait_for_window
from winscript.tools.ui_interaction import click, type_text, read_text, press_key, get_ui_tree, coordinate_click
from winscript.tools.com_office import (
    excel_read_cell, excel_write_cell, excel_read_range,
    outlook_send_email, outlook_read_inbox
)
from winscript.tools.filesystem import (
    read_file_text, write_file_text, list_dir,
    move_file, copy_file, delete_file, file_exists
)
from winscript.tools.screen import take_screenshot, get_active_window, get_clipboard, set_clipboard
from winscript.tools.shell import run_powershell
from winscript.core.state_diff import DesktopSnapshot

from winscript.adapters.excel_adapter import excel
from winscript.adapters.outlook_adapter import outlook
from winscript.adapters.chrome_adapter import chrome
from winscript.adapters.explorer_adapter import explorer
from winscript.adapters.notepad_adapter import notepad

from winscript.core.workflow import recorder, Workflow, WORKFLOW_DIR
from winscript.core.replay_engine import register_tool, replay_workflow

mcp = FastMCP(
    name="winscript",
    version="0.1.0"
)

def get_state_snapshot() -> str:
    """Get current desktop state snapshot.
    Active window, all open windows, clipboard preview.
    Use this before any action sequence to establish baseline."""
    snap = DesktopSnapshot.capture()
    lines = [
        f"Active window: {snap.active_window}",
        f"Open windows ({len(snap.open_windows)}): {', '.join(snap.open_windows[:10])}",
        f"Clipboard: {snap.clipboard[:80]!r}" if snap.clipboard else "Clipboard: (empty)",
    ]
    return "\n".join(lines)

# App Control
mcp.tool()(open_app)
mcp.tool()(close_app)
mcp.tool()(focus_app)
mcp.tool()(get_running_apps)
mcp.tool()(wait_for_window)

# UI Interaction
mcp.tool()(click)
mcp.tool()(type_text)
mcp.tool()(read_text)
mcp.tool()(press_key)
mcp.tool()(get_ui_tree)
mcp.tool()(coordinate_click)

# COM Office
mcp.tool()(excel_read_cell)
mcp.tool()(excel_write_cell)
mcp.tool()(excel_read_range)
mcp.tool()(outlook_send_email)
mcp.tool()(outlook_read_inbox)

# File System
mcp.tool()(read_file_text)
mcp.tool()(write_file_text)
mcp.tool()(list_dir)
mcp.tool()(move_file)
mcp.tool()(copy_file)
mcp.tool()(delete_file)
mcp.tool()(file_exists)

# Screen
mcp.tool()(take_screenshot)
mcp.tool()(get_active_window)
mcp.tool()(get_clipboard)
mcp.tool()(set_clipboard)

# Shell
mcp.tool()(run_powershell)

# State
mcp.tool()(get_state_snapshot)

# Excel adapter
mcp.tool(name="excel_open")(excel.open)
mcp.tool(name="excel_save")(excel.save)
mcp.tool(name="excel_close")(excel.close)

# Chrome adapter
mcp.tool(name="chrome_open")(chrome.open)
mcp.tool(name="chrome_navigate")(chrome.navigate)
mcp.tool(name="chrome_get_url")(chrome.get_url)
mcp.tool(name="chrome_get_title")(chrome.get_title)
mcp.tool(name="chrome_new_tab")(chrome.new_tab)
mcp.tool(name="chrome_close_tab")(chrome.close_tab)
mcp.tool(name="chrome_find_on_page")(chrome.find_on_page)

# Explorer adapter
mcp.tool(name="explorer_open")(explorer.open)
mcp.tool(name="explorer_navigate")(explorer.navigate)

# Notepad adapter
mcp.tool(name="notepad_open")(notepad.open)
mcp.tool(name="notepad_type")(notepad.type)
mcp.tool(name="notepad_save")(notepad.save)
mcp.tool(name="notepad_close")(notepad.close)

# Outlook adapter
mcp.tool(name="outlook_open")(outlook.open)

# Workflow Tools
def workflow_record_start(name: str, description: str = "") -> str:
    """Start recording a workflow. All subsequent tool calls are captured.
    Example: workflow_record_start("daily_report","Opens report and emails it")"""
    return recorder.start(name, description)

def workflow_record_stop() -> str:
    """Stop recording and save the workflow.
    Example: workflow_record_stop()"""
    return recorder.stop()

def workflow_record_discard() -> str:
    """Discard current recording without saving."""
    return recorder.discard()

def workflow_replay(name: str, dry_run: bool = False) -> str:
    """Replay a saved workflow by name.
    dry_run=True shows steps without executing.
    Example: workflow_replay("daily_report"), workflow_replay("daily_report", dry_run=True)"""
    return replay_workflow(name, dry_run)

def workflow_list() -> str:
    """List all saved workflows with step counts and run history."""
    workflows = Workflow.list_all()
    if not workflows:
        return "No workflows saved yet."
    lines = []
    for wf in workflows:
        lines.append(
            f"'{wf['name']}' — {wf['steps']} steps | "
            f"Run {wf['run_count']}x | {wf['description']}"
        )
    return "\n".join(lines)

def workflow_delete(name: str) -> str:
    """Delete a saved workflow by name."""
    path = WORKFLOW_DIR / f"{name}.json"
    if not path.exists():
        return f"ERROR: Workflow '{name}' not found."
    path.unlink()
    return f"Deleted workflow '{name}'"

mcp.tool()(workflow_record_start)
mcp.tool()(workflow_record_stop)
mcp.tool()(workflow_record_discard)
mcp.tool()(workflow_replay)
mcp.tool()(workflow_list)
mcp.tool()(workflow_delete)

# Register tools for replay
for name, fn in [
    ("open_app", open_app), ("close_app", close_app),
    ("focus_app", focus_app), ("click", click),
    ("type_text", type_text), ("press_key", press_key),
    ("read_file_text", read_file_text), ("write_file_text", write_file_text),
    ("list_dir", list_dir), ("move_file", move_file),
    ("copy_file", copy_file), ("delete_file", delete_file),
    ("file_exists", file_exists), ("run_powershell", run_powershell),
]:
    register_tool(name, fn)

def cleanup_orphaned_com_processes():
    """Kill orphaned Excel and Outlook processes on server startup to prevent COM object leaks."""
    try:
        subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE", "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["taskkill", "/F", "/IM", "OUTLOOK.EXE", "/T"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

if __name__ == "__main__":
    cleanup_orphaned_com_processes()
    mcp.run()