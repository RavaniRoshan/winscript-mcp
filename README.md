# 🚀 WinScript

<p align="center">
  <b>AppleScript for Windows — as a Model Context Protocol (MCP) server.</b>
</p>

<p align="center">
  <img alt="Python" src="https://img.shields.io/badge/python-3.9%2B-blue.svg">
  <img alt="MCP" src="https://img.shields.io/badge/MCP-Ready-brightgreen">
  <img alt="Windows Only" src="https://img.shields.io/badge/Platform-Windows-0078D6">
</p>

---

## 🌟 The Vision

Control any Windows application from an AI agent seamlessly. WinScript acts as the native bridge between intelligent agents (like Claude or local LLMs) and your local Windows environment. 

Whether it's manipulating Excel cells, clicking through legacy UIs, automating Outlook emails, or executing PowerShell escape hatches—WinScript handles it with zero manual interaction.

---

## 🏗️ Architecture

WinScript operates over standard MCP (stdio) to provide a clean API boundary over Windows' fragmented automation primitives:

```text
AI Agent (Claude / local LLM)
        │
        │  MCP protocol (stdio)
        ▼
┌──────────────────────────────┐
│       WinScript MCP Server   │  ← winscript/server.py
│       FastMCP + Python       │
└──────┬───────────────────────┘
       │
  ┌────┼──────────────┐
  ▼    ▼              ▼
pywinauto  win32com  subprocess
(UI ctrl) (COM/Office) (PowerShell)
  │         │
  └────┬────┘
       ▼
Any Windows App / System
```

---

## 🛠️ Installation & Setup

Get started in three simple steps:

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```
*For OCR fallback: install Tesseract https://github.com/tesseract-ocr/tesseract*

### 2. Run the Server
```bash
python winscript/server.py
```

### 3. Wire into Claude Desktop
Edit your `%APPDATA%\Claude\claude_desktop_config.json` to include the WinScript server using the absolute path to your script:

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
*Restart Claude Desktop, and the WinScript tool suite will automatically appear.*

---

## ⚡ Core Capabilities (52 Tools)

WinScript provides a comprehensive, battle-tested toolset for deep desktop integration:

| Category | Tools | Description |
| :--- | :--- | :--- |
| 🪟 **App Control** | `open_app`, `close_app`, `focus_app`, `get_running_apps`, `wait_for_window` | Open, close, and manage foreground states, with resilient polling for slow-loading software. |
| 🖱️ **UI Interaction** | `click`, `coordinate_click`, `type_text`, `read_text`, `press_key`, `get_ui_tree` | Deep semantic UI tree interaction, falling back to raw coordinate pixel clicks when required. |
| 📊 **COM Office** | `excel_read_cell`, `excel_write_cell`, `excel_read_range`, `outlook_send_email`, `outlook_read_inbox` | Native integration with Microsoft Office. Zero zombie processes thanks to aggressive background cleanup. |
| 📁 **File System** | `read_file_text`, `write_file_text`, `list_dir`, `move_file`, `copy_file`, `delete_file`, `file_exists` | Safe, sandboxed local file and directory manipulation. |
| 📸 **Screen & Clipboard** | `take_screenshot`, `get_active_window`, `get_clipboard`, `set_clipboard` | Instantly pipe base64 inline screenshots to Vision models or modify clipboard state. |
| 🐚 **Shell** | `run_powershell` | The ultimate escape hatch for admin-level operations and process queries. |
| 🔌 **Adapters** | `excel_*`, `outlook_*`, `chrome_*`, `explorer_*`, `notepad_*` | High-level semantic interfaces for common applications. |
| 🔄 **State Diffing** | `get_state_snapshot` | Mutating tools automatically return a computed `StateDelta` (e.g. `Active window changed: '' → 'Notepad' | Duration: 2100ms`). |
| 🤖 **Workflows** | `workflow_record_start`, `workflow_record_stop`, `workflow_replay`, `workflow_list`, `workflow_delete` | The "macro compiler for agents." Record any multi-step sequence and replay it on demand. |

---

## 🛡️ Resilience & Error Handling

Windows automation is notoriously brittle. WinScript is built for reality:

- **Graceful Failures:** Tools never crash the server. They return descriptive `"ERROR: ..."` strings to the agent so it can strategize a new approach.
- **Circuit Breaker:** After 5 identical consecutive tool failures, a `WinScriptMaxRetriesError` is raised (Hard Stop) to prevent runaway LLM loops.
- **Zombie Cleanup:** The server aggressively purges orphaned `EXCEL.EXE` and `OUTLOOK.EXE` processes on startup to prevent COM object memory leaks.

---

## 🔒 Security & Permissions

**WinScript is Open by Default.** 

Friction kills agent usability, so we designed WinScript without an interactive auth layer. However, exposing your desktop and shell directly to an AI agent comes with inherent risks. 

**Please ensure you are comfortable with the following before running:**
- The agent has full read/write access to your user files.
- The agent can execute arbitrary PowerShell commands.
- The agent can send emails from your authenticated Outlook account.

---

## ⚠️ Known Limitations

> **Platform Lock:** By design, WinScript strictly targets **Windows** via `pywinauto` and Win32 COM APIs. 
> 
> **Electron & UWP Apps:** Applications like VS Code, Discord, or Slack often have obfuscated accessibility trees. For these, rely on the `take_screenshot` + `coordinate_click` vision loop.
> 
> **Permissions:** Elevated (Admin) applications cannot be automated unless the Python environment running WinScript is also launched with Administrator privileges.