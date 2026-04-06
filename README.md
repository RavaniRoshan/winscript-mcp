# WinScript

> AppleScript for Windows — as an MCP server.

Control any Windows app from an AI agent. Open apps, click buttons,
read Excel cells, send Outlook emails, take screenshots. Zero manual interaction.

## Install

pip install -r requirements.txt

## Run

python winscript/server.py

## Wire into Claude Desktop

Add to `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "winscript": {
      "command": "python",
      "args": ["/home/roshandamm/winscript/winscript-mcp/winscript/server.py"]
    }
  }
}
```

Restart Claude Desktop. WinScript tools appear automatically.

## Tools (25 total)

App Control: open_app, close_app, focus_app, get_running_apps
UI: click, type_text, read_text, press_key, get_ui_tree
Office: excel_read_cell, excel_write_cell, excel_read_range, outlook_send_email, outlook_read_inbox
Files: read_file_text, write_file_text, list_dir, move_file, copy_file, delete_file, file_exists
Screen: take_screenshot, get_active_window, get_clipboard, set_clipboard

## Error Handling

- Tools return "ERROR: ..." strings to the agent on failure
- After 5 identical consecutive failures: WinScriptMaxRetriesError raised (hard stop)

## Limitations

- Windows only (by design)
- UWP/Electron apps: limited accessibility tree — use take_screenshot as fallback
- Elevated (admin) apps: cannot be automated from non-admin process