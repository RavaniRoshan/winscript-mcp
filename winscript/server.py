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
from winscript.core.modes import get_mode, set_mode, is_blocked, ExecutionMode


from winscript.intents import (
    open_latest_file,
    send_email_with_content,
    find_in_folder,
    read_active_document,
    summarize_screen,
)

mcp = FastMCP(
    name="winscript",
    version="0.1.0"
)


def mode_guarded(tool_name: str, fn):
    import functools
    @functools.wraps(fn)
    def wrapper(**kwargs):
        if is_blocked(tool_name):
            return (
                f"BLOCKED: '{tool_name}' is not allowed in SAFE mode. "
                f"Switch to standard mode with set_execution_mode('standard')."
            )
        return fn(**kwargs)
    return wrapper

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
mcp.tool()(mode_guarded("open_app", audited("open_app", open_app)))
mcp.tool()(mode_guarded("close_app", audited("close_app", close_app)))
mcp.tool()(mode_guarded("focus_app", audited("focus_app", focus_app)))
mcp.tool()(mode_guarded("get_running_apps", audited("get_running_apps", get_running_apps)))
mcp.tool()(mode_guarded("wait_for_window", audited("wait_for_window", wait_for_window)))

# UI Interaction
mcp.tool()(mode_guarded("click", audited("click", click)))
mcp.tool()(mode_guarded("type_text", audited("type_text", type_text)))
mcp.tool()(mode_guarded("read_text", audited("read_text", read_text)))
mcp.tool()(mode_guarded("press_key", audited("press_key", press_key)))
mcp.tool()(mode_guarded("get_ui_tree", audited("get_ui_tree", get_ui_tree)))
mcp.tool()(mode_guarded("coordinate_click", audited("coordinate_click", coordinate_click)))

# COM Office
mcp.tool()(mode_guarded("excel_read_cell", audited("excel_read_cell", excel_read_cell)))
mcp.tool()(mode_guarded("excel_write_cell", audited("excel_write_cell", excel_write_cell)))
mcp.tool()(mode_guarded("excel_read_range", audited("excel_read_range", excel_read_range)))
mcp.tool()(mode_guarded("outlook_send_email", audited("outlook_send_email", outlook_send_email)))
mcp.tool()(mode_guarded("outlook_read_inbox", audited("outlook_read_inbox", outlook_read_inbox)))

# File System
mcp.tool()(mode_guarded("read_file_text", audited("read_file_text", read_file_text)))
mcp.tool()(mode_guarded("write_file_text", audited("write_file_text", write_file_text)))
mcp.tool()(mode_guarded("list_dir", audited("list_dir", list_dir)))
mcp.tool()(mode_guarded("move_file", audited("move_file", move_file)))
mcp.tool()(mode_guarded("copy_file", audited("copy_file", copy_file)))
mcp.tool()(mode_guarded("delete_file", audited("delete_file", delete_file)))
mcp.tool()(mode_guarded("file_exists", audited("file_exists", file_exists)))

# Screen
mcp.tool()(mode_guarded("take_screenshot", audited("take_screenshot", take_screenshot)))
mcp.tool()(mode_guarded("get_active_window", audited("get_active_window", get_active_window)))
mcp.tool()(mode_guarded("get_clipboard", audited("get_clipboard", get_clipboard)))
mcp.tool()(mode_guarded("set_clipboard", audited("set_clipboard", set_clipboard)))

# Shell
mcp.tool()(mode_guarded("run_powershell", audited("run_powershell", run_powershell)))

# State
mcp.tool()(mode_guarded("get_state_snapshot", audited("get_state_snapshot", get_state_snapshot)))

# Excel adapter
mcp.tool(name="excel_open")(mode_guarded("excel_open", audited("excel_open", excel.open)))
mcp.tool(name="excel_save")(mode_guarded("excel_save", audited("excel_save", excel.save)))
mcp.tool(name="excel_close")(mode_guarded("excel_close", audited("excel_close", excel.close)))

# Chrome adapter
mcp.tool(name="chrome_open")(mode_guarded("chrome_open", audited("chrome_open", chrome.open)))
mcp.tool(name="chrome_navigate")(mode_guarded("chrome_navigate", audited("chrome_navigate", chrome.navigate)))
mcp.tool(name="chrome_get_url")(mode_guarded("chrome_get_url", audited("chrome_get_url", chrome.get_url)))
mcp.tool(name="chrome_get_title")(mode_guarded("chrome_get_title", audited("chrome_get_title", chrome.get_title)))
mcp.tool(name="chrome_new_tab")(mode_guarded("chrome_new_tab", audited("chrome_new_tab", chrome.new_tab)))
mcp.tool(name="chrome_close_tab")(mode_guarded("chrome_close_tab", audited("chrome_close_tab", chrome.close_tab)))
mcp.tool(name="chrome_find_on_page")(mode_guarded("chrome_find_on_page", audited("chrome_find_on_page", chrome.find_on_page)))

# Explorer adapter
mcp.tool(name="explorer_open")(mode_guarded("explorer_open", audited("explorer_open", explorer.open)))
mcp.tool(name="explorer_navigate")(mode_guarded("explorer_navigate", audited("explorer_navigate", explorer.navigate)))

# Notepad adapter
mcp.tool(name="notepad_open")(mode_guarded("notepad_open", audited("notepad_open", notepad.open)))
mcp.tool(name="notepad_type")(mode_guarded("notepad_type", audited("notepad_type", notepad.type)))
mcp.tool(name="notepad_save")(mode_guarded("notepad_save", audited("notepad_save", notepad.save)))
mcp.tool(name="notepad_close")(mode_guarded("notepad_close", audited("notepad_close", notepad.close)))

# Outlook adapter
mcp.tool(name="outlook_open")(mode_guarded("outlook_open", audited("outlook_open", outlook.open)))

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

mcp.tool()(mode_guarded("workflow_record_start", audited("workflow_record_start", workflow_record_start)))
mcp.tool()(mode_guarded("workflow_record_stop", audited("workflow_record_stop", workflow_record_stop)))
mcp.tool()(mode_guarded("workflow_record_discard", audited("workflow_record_discard", workflow_record_discard)))
mcp.tool()(mode_guarded("workflow_replay", audited("workflow_replay", workflow_replay)))
mcp.tool()(mode_guarded("workflow_list", audited("workflow_list", workflow_list)))
mcp.tool()(mode_guarded("workflow_delete", audited("workflow_delete", workflow_delete)))

# Intents
mcp.tool()(mode_guarded("open_latest_file", audited("open_latest_file", open_latest_file)))
mcp.tool()(mode_guarded("send_email_with_content", audited("send_email_with_content", send_email_with_content)))
mcp.tool()(mode_guarded("find_in_folder", audited("find_in_folder", find_in_folder)))
mcp.tool()(mode_guarded("read_active_document", audited("read_active_document", read_active_document)))
mcp.tool()(mode_guarded("summarize_screen", audited("summarize_screen", summarize_screen)))

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


def set_execution_mode(mode: str) -> str:
    """Set execution mode. 'safe' = read-only, 'standard' = full access.
    Default is standard. Use safe when you only want to inspect, not change anything.
    Example: set_execution_mode("safe"), set_execution_mode("standard")"""
    if mode not in ("safe", "standard"):
        return "ERROR: mode must be 'safe' or 'standard'"
    set_mode(ExecutionMode(mode))
    from winscript.core.modes import BLOCKED_IN_SAFE_MODE
    blocked_count = len(BLOCKED_IN_SAFE_MODE) if mode == "safe" else 0
    return (
        f"Execution mode: {mode.upper()}. "
        f"{'Read-only — ' + str(blocked_count) + ' write tools blocked.' if mode == 'safe' else 'Full access enabled.'}"
    )

def get_execution_mode() -> str:
    """Get the current execution mode.
    Example: get_execution_mode()"""
    mode = get_mode()
    from winscript.core.modes import BLOCKED_IN_SAFE_MODE
    blocked = len(BLOCKED_IN_SAFE_MODE) if mode == ExecutionMode.SAFE else 0
    return f"Current mode: {mode.value.upper()} | {blocked} tools blocked"

mcp.tool()(set_execution_mode)
mcp.tool()(get_execution_mode)

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
