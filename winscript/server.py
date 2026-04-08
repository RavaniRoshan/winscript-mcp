import subprocess
from fastmcp import FastMCP
import functools
import time

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
from winscript.core.audit import log_action, query_log, get_failure_summary

mcp = FastMCP(
    name="winscript",
    version="0.1.0"
)

def audited(tool_name: str, fn):
    """Wrap a tool function with audit logging."""
    @functools.wraps(fn)
    def wrapper(**kwargs):
        start = time.time()
        result = ""
        error = ""
        try:
            result = fn(**kwargs)
            if isinstance(result, str) and result.startswith("ERROR"):
                error = result
        except Exception as e:
            error = str(e)
            result = f"ERROR: {error}"
            raise
        finally:
            duration_ms = (time.time() - start) * 1000
            log_action(
                tool=tool_name,
                args=kwargs,
                result=str(result) if result is not None else "",
                error=error,
                duration_ms=duration_ms,
            )
        return result
    return wrapper

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
mcp.tool()(audited("open_app", open_app))
mcp.tool()(audited("close_app", close_app))
mcp.tool()(audited("focus_app", focus_app))
mcp.tool()(audited("get_running_apps", get_running_apps))
mcp.tool()(audited("wait_for_window", wait_for_window))

# UI Interaction
mcp.tool()(audited("click", click))
mcp.tool()(audited("type_text", type_text))
mcp.tool()(audited("read_text", read_text))
mcp.tool()(audited("press_key", press_key))
mcp.tool()(audited("get_ui_tree", get_ui_tree))
mcp.tool()(audited("coordinate_click", coordinate_click))

# COM Office
mcp.tool()(audited("excel_read_cell", excel_read_cell))
mcp.tool()(audited("excel_write_cell", excel_write_cell))
mcp.tool()(audited("excel_read_range", excel_read_range))
mcp.tool()(audited("outlook_send_email", outlook_send_email))
mcp.tool()(audited("outlook_read_inbox", outlook_read_inbox))

# File System
mcp.tool()(audited("read_file_text", read_file_text))
mcp.tool()(audited("write_file_text", write_file_text))
mcp.tool()(audited("list_dir", list_dir))
mcp.tool()(audited("move_file", move_file))
mcp.tool()(audited("copy_file", copy_file))
mcp.tool()(audited("delete_file", delete_file))
mcp.tool()(audited("file_exists", file_exists))

# Screen
mcp.tool()(audited("take_screenshot", take_screenshot))
mcp.tool()(audited("get_active_window", get_active_window))
mcp.tool()(audited("get_clipboard", get_clipboard))
mcp.tool()(audited("set_clipboard", set_clipboard))

# Shell
mcp.tool()(audited("run_powershell", run_powershell))

# State
mcp.tool()(audited("get_state_snapshot", get_state_snapshot))

# Excel adapter
mcp.tool(name="excel_open")(audited("excel_open", excel.open))
mcp.tool(name="excel_save")(audited("excel_save", excel.save))
mcp.tool(name="excel_close")(audited("excel_close", excel.close))

# Chrome adapter
mcp.tool(name="chrome_open")(audited("chrome_open", chrome.open))
mcp.tool(name="chrome_navigate")(audited("chrome_navigate", chrome.navigate))
mcp.tool(name="chrome_get_url")(audited("chrome_get_url", chrome.get_url))
mcp.tool(name="chrome_get_title")(audited("chrome_get_title", chrome.get_title))
mcp.tool(name="chrome_new_tab")(audited("chrome_new_tab", chrome.new_tab))
mcp.tool(name="chrome_close_tab")(audited("chrome_close_tab", chrome.close_tab))
mcp.tool(name="chrome_find_on_page")(audited("chrome_find_on_page", chrome.find_on_page))

# Explorer adapter
mcp.tool(name="explorer_open")(audited("explorer_open", explorer.open))
mcp.tool(name="explorer_navigate")(audited("explorer_navigate", explorer.navigate))

# Notepad adapter
mcp.tool(name="notepad_open")(audited("notepad_open", notepad.open))
mcp.tool(name="notepad_type")(audited("notepad_type", notepad.type))
mcp.tool(name="notepad_save")(audited("notepad_save", notepad.save))
mcp.tool(name="notepad_close")(audited("notepad_close", notepad.close))

# Outlook adapter
mcp.tool(name="outlook_open")(audited("outlook_open", outlook.open))

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

mcp.tool()(audited("workflow_record_start", workflow_record_start))
mcp.tool()(audited("workflow_record_stop", workflow_record_stop))
mcp.tool()(audited("workflow_record_discard", workflow_record_discard))
mcp.tool()(audited("workflow_replay", workflow_replay))
mcp.tool()(audited("workflow_list", workflow_list))
mcp.tool()(audited("workflow_delete", workflow_delete))

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

# Audit Logging Tools
def get_audit_log(limit: int = 20, tool_filter: str = "") -> str:
    """Get recent action audit log.
    tool_filter: optional tool name to filter by.
    Example: get_audit_log(10), get_audit_log(20, "click")"""
    rows = query_log(limit, tool_filter)
    if not rows:
        return "No audit entries found."
    lines = []
    for r in rows:
        ts = time.strftime("%H:%M:%S", time.localtime(r["timestamp"]))
        status = "✗" if r["error"] else "✓"
        lines.append(
            f"[{ts}] {status} {r['tool']}({r['args'][:60]}) → "
            f"{r['result'][:60]} [{r['duration_ms']:.0f}ms]"
        )
    return "\n".join(lines)

def get_failure_report() -> str:
    """Get a summary of which tools fail most often.
    Useful for diagnosing unreliable automations."""
    rows = get_failure_summary()
    if not rows:
        return "No data yet."
    lines = ["Tool failure summary:"]
    for r in rows:
        rate = (r["failures"] / r["total"] * 100) if r["total"] > 0 else 0
        lines.append(
            f"  {r['tool']}: {r['failures']}/{r['total']} failures "
            f"({rate:.0f}%) | avg {r['avg_ms']:.0f}ms"
        )
    return "\n".join(lines)

mcp.tool()(get_audit_log)
mcp.tool()(get_failure_report)

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
