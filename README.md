<div align="center">

```
██╗    ██╗██╗███╗   ██╗███████╗ ██████╗██████╗ ██╗██████╗ ████████╗
██║    ██║██║████╗  ██║██╔════╝██╔════╝██╔══██╗██║██╔══██╗╚══██╔══╝
██║ █╗ ██║██║██╔██╗ ██║███████╗██║     ██████╔╝██║██████╔╝   ██║   
██║███╗██║██║██║╚██╗██║╚════██║██║     ██╔══██╗██║██╔═══╝    ██║   
╚███╔███╔╝██║██║ ╚████║███████║╚██████╗██║  ██║██║██║        ██║   
 ╚══╝╚══╝ ╚═╝╚═╝  ╚═══╝╚══════╝ ╚═════╝╚═╝  ╚═╝╚═╝╚═╝        ╚═╝   
```

**AppleScript for Windows. Built for AI agents.**

Windows 10/11 · Python 3.10+ · MCP Protocol

[![Python](https://img.shields.io/badge/python-3.10+-blue.svg)](https://python.org)
[![MCP](https://img.shields.io/badge/protocol-MCP-orange.svg)](https://modelcontextprotocol.io)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Tools](https://img.shields.io/badge/tools-59-purple.svg)](#tools)
[![PyPI](https://img.shields.io/pypi/v/winscript.svg)](https://pypi.org/project/winscript/)
[![smithery](https://img.shields.io/badge/Smithery-winscript-blue)](https://smithery.ai/server/winscript)

</div>

---

macOS has AppleScript.  
Windows had nothing clean for AI agents.  
Until now.

WinScript is a **state-aware, replayable, audited** Windows automation server for AI agents. It wraps 4 fragmented Windows automation primitives — UI Automation, COM, Win32, and OCR — into a single MCP server that any agent can call.

Not a wrapper. Not a toy. Infrastructure.

---

## Quick Start — Choose Your Deployment

No more `pip install` friction. Pick what works for you:

### Option 1: Smithery.ai (Zero Setup)
Add WinScript to Claude Desktop, Cursor, or any MCP client in one click:

[![Deploy to Smithery](https://smithery.ai/deploy.svg)](https://smithery.ai/server/winscript)

### Option 2: PyPI (One Command)
```bash
pip install winscript
winscript
```

### Option 3: Docker (Isolated)
```bash
docker run -v %USERPROFILE%/.winscript:~/.winscript ghcr.io/roshandamm/winscript-mcp:latest
```

### Option 4: Direct Download (No Install)
```bash
git clone https://github.com/roshandamm/winscript-mcp.git
cd winscript-mcp
python winscript-server.py
```

All options start an MCP server that any AI agent can connect to.

---

## The difference

Every other Windows automation tool gives you actions.  
WinScript gives you **actions + state**.

```
# What others give you:
click("Submit")
→ "Clicked Submit"

# What WinScript gives you:
click("Submit")
→ "Clicked 'Submit' via uia_name [confidence 1.0] |
   Active window: 'Form' → 'Confirmation' |
   New windows: ['Success Dialog'] |
   Duration: 312ms"
```

You don't just know what you did. You know what changed.

---

## Detailed Installation

### Option 1: Install from PyPI
```bash
pip install winscript
```
Then run: `winscript` or `python -m winscript.server`

### Option 2: Use Smithery.ai (Recommended for AI Agents)
1. Go to [smithery.ai/server/winscript](https://smithery.ai/server/winscript)
2. Click "Add to Claude" (or your MCP client)
3. Done — no local setup needed

### Option 3: Run with Docker
```bash
# Pull and run
docker run -d --name winscript \
  -v %USERPROFILE%/.winscript:~/.winscript \
  ghcr.io/roshandamm/winscript-mcp:latest

# Or build locally
docker build -t winscript:latest .
docker run -d --name winscript -v %USERPROFILE%/.winscript:~/.winscript winscript:latest
```

### Option 4: Run from Source (No Install)
```bash
git clone https://github.com/roshandamm/winscript-mcp.git
cd winscript-mcp
pip install -r requirements.txt
python winscript-server.py
```

### Optional: OCR Fallback (Layer 4)
For better element detection in broken UI trees:
```bash
# Install Tesseract: https://github.com/tesseract-ocr/tesseract
pip install pytesseract
```
```

---

## Wire into Claude Desktop

```json
{
  "mcpServers": {
    "winscript": {
      "command": "python",
      "args": ["-m", "winscript.server"]
    }
  }
}
```

Restart Claude Desktop. 59 tools appear automatically.

---

## Five things that make WinScript different

### 1. Five-layer selector fallback chain

Other tools fail when the UI tree is bad (Electron apps, UWP, legacy Win32).  
WinScript tries 5 strategies before giving up.

```
Layer 1 → UIA by element name       (fast, exact)
Layer 2 → UIA by automation_id      (for apps that label controls)
Layer 3 → UIA fuzzy role match      (partial name, control type)
Layer 4 → OCR scan + bounding box   (when UI tree is broken)
Layer 5 → Raw coordinates           (click("x=412,y=308"))
```

Every tool call tells you which layer succeeded:
```
"Clicked 'Login' in 'Slack' [via ocr, confidence 0.91]"
```

### 2. State diffing after every action

Before you act, WinScript snapshots the desktop.  
After you act, it snapshots again and diffs.

```python
# The agent knows what actually happened:
open_app("excel")
→ "Opened Excel | Active window: '' → 'Book1 - Excel' | 
   New windows: ['Microsoft Excel - Book1'] | Duration: 2140ms"

type_text("Notepad", "hello")
→ "Typed 5 chars | No window change detected | Duration: 89ms"
```

No more "did it work?" loops.

### 3. Workflow recorder and replay

Record any successful multi-step sequence. Replay it on demand.  
No human-written macros. No brittle scripts.

```python
# Record:
workflow_record_start("daily_report", "Opens report and emails it")
open_latest_file("C:/reports", "xlsx")
read_active_document()
send_email_with_content("team@co.com", "Daily Report", "clipboard")
workflow_record_stop()
→ "Workflow 'daily_report' saved: 3 steps"

# Replay any time:
workflow_replay("daily_report")
→ "Step 1 ✓ open_latest_file → Opened q1_2026.xlsx [2100ms]
   Step 2 ✓ read_active_document → [clipboard content] [340ms]
   Step 3 ✓ send_email_with_content → Email sent [890ms]"

# Preview before running:
workflow_replay("daily_report", dry_run=True)
```

### 4. Semantic intent layer

Five high-level intents so agents don't have to think in clicks.

```python
open_latest_file("C:/reports", "xlsx")     # Find + open newest xlsx
send_email_with_content("a@b.com", "Re", "clipboard")  # Clipboard → email
find_in_folder("C:/docs", "invoice", "pdf")  # Find matching files
read_active_document()                      # Select-all copy current doc
summarize_screen()                          # Screenshot → agent vision
```

### 5. Full audit log + local memory

Every action, input, output, state delta, selector layer, and failure logged to `~/.winscript/audit.db`.

```python
get_audit_log(10)
→ "[14:23:01] ✓ open_app({'name':'notepad'}) → Opened notepad [2100ms]
   [14:23:03] ✓ type_text({'text':'hello'}) → Typed 5 chars [89ms]
   [14:23:11] ✗ click({'element':'Submit'}) → ERROR: No element found [412ms]"

get_failure_report()
→ "click: 3/12 failures (25%) | avg 380ms
   open_app: 0/8 failures (0%) | avg 2100ms"
```

And memory persists across sessions:

```python
what_files_have_i_opened(5, "xlsx")
→ "C:/reports/q1_2026.xlsx — opened 4x | last: 14:23 08/04"

what_did_i_do(5)
→ "[14:23] open_app → Opened notepad
   [14:22] excel_read_cell → 47230.5
   [14:21] outlook_send_email → Email sent to team@co.com"
```

---

## All 59 tools

<details>
<summary><b>App Control (4)</b></summary>

| Tool | What it does |
|---|---|
| `open_app(name)` | Open any app by name or alias |
| `close_app(title_hint)` | Close by partial window title |
| `focus_app(title_hint)` | Bring to foreground |
| `get_running_apps()` | List all open windows + PIDs |

</details>

<details>
<summary><b>UI Interaction (5)</b></summary>

| Tool | What it does |
|---|---|
| `click(app_title, element_name)` | Click element — 5-layer fallback |
| `type_text(app_title, text)` | Type into focused element |
| `read_text(app_title, element_name)` | Read text from element |
| `press_key(key, app_title)` | Keyboard shortcuts |
| `get_ui_tree(app_title, depth)` | Discover all UI elements |

</details>

<details>
<summary><b>COM Office (5)</b></summary>

| Tool | What it does |
|---|---|
| `excel_read_cell(filepath, sheet, cell)` | Read one cell |
| `excel_write_cell(filepath, sheet, cell, value)` | Write one cell + save |
| `excel_read_range(filepath, sheet, start, end)` | Read range as CSV |
| `outlook_send_email(to, subject, body)` | Send email |
| `outlook_read_inbox(count)` | Read N recent emails |

</details>

<details>
<summary><b>File System (7)</b></summary>

`read_file_text` · `write_file_text` · `list_dir` · `move_file` · `copy_file` · `delete_file` · `file_exists`

</details>

<details>
<summary><b>Screen + Clipboard (4)</b></summary>

| Tool | What it does |
|---|---|
| `take_screenshot(region)` | Base64 PNG — agent sees your screen |
| `get_active_window()` | Current focused window title |
| `get_clipboard()` | Read clipboard |
| `set_clipboard(text)` | Write clipboard |

</details>

<details>
<summary><b>App Adapters (15)</b></summary>

Typed semantic APIs for specific apps. No more clicking blind.

```python
# Excel
excel_open(filepath)  ·  excel_save()  ·  excel_close(save)

# Chrome
chrome_open(url)  ·  chrome_navigate(url)  ·  chrome_get_url()
chrome_get_title()  ·  chrome_new_tab()  ·  chrome_close_tab()
chrome_find_on_page(text)

# Notepad
notepad_open(filepath)  ·  notepad_type(text)
notepad_save()  ·  notepad_close(save)

# Explorer
explorer_open(path)  ·  explorer_navigate(path)

# Outlook
outlook_open()
```

</details>

<details>
<summary><b>Workflow Recorder + Replay (6)</b></summary>

```python
workflow_record_start(name, description)
workflow_record_stop()
workflow_record_discard()
workflow_replay(name, dry_run)
workflow_list()
workflow_delete(name)
```

</details>

<details>
<summary><b>Semantic Intents (5)</b></summary>

```python
open_latest_file(folder, extension)
send_email_with_content(to, subject, content_source)
find_in_folder(folder, search_term, extension)
read_active_document()
summarize_screen()
```

</details>

<details>
<summary><b>Audit + Memory + State (10)</b></summary>

```python
# Audit
get_audit_log(limit, tool_filter)
get_failure_report()

# Memory
what_windows_have_i_seen(limit)
what_files_have_i_opened(limit, extension)
what_did_i_do(limit)

# State
get_state_snapshot()

# Modes
set_execution_mode(mode)   # "safe" | "standard"
get_execution_mode()
```

</details>

---

## App aliases

```python
open_app("notepad")    # notepad.exe
open_app("chrome")     # chrome.exe
open_app("firefox")    # firefox.exe
open_app("edge")       # msedge.exe
open_app("excel")      # EXCEL.EXE
open_app("word")       # WINWORD.EXE
open_app("outlook")    # OUTLOOK.EXE
open_app("explorer")   # explorer.exe
open_app("terminal")   # wt.exe
open_app("vscode")     # Code.exe
open_app("cursor")     # Cursor.exe
```

---

## Error handling

Tools return `"ERROR: ..."` strings to the agent on failure — never crash your agent.

After **5 consecutive identical failures** on the same tool + args:  
`WinScriptMaxRetriesError` is raised. Hard stop. Change your args and try again.

```python
get_failure_report()
# See which tools are failing and why before they hit the limit
```

---

## Execution modes

```python
set_execution_mode("safe")
# Read-only: screenshots, reads, audits only
# Blocks: write, delete, click, type, send email, open apps

set_execution_mode("standard")
# Full access (default)
```

---

## Where recordings live

```
~/.winscript/
├── audit.db          # every action ever taken
├── memory.db         # windows, files, action history
└── workflows/
    ├── daily_report.json
    └── your_workflow.json
```

Auto-purge: audit logs older than 30 days are deleted on startup.

---

## Limitations (honest)

- **Windows only.** By design. This is not a bug.
- **Elevated (admin) apps** cannot be automated from a non-admin process.
- **UWP + Electron apps** have broken accessibility trees. WinScript falls back to OCR then coordinates — but complex UIs still sometimes fail.
- **Requires Tesseract** for OCR fallback (Layer 4). Without it, WinScript skips to Layer 5.
- **COM automation** (Excel, Outlook) requires those apps installed and licensed.

---

## Built on

| Layer | Library |
|---|---|
| MCP server | FastMCP |
| UI automation | pywinauto + uiautomation |
| COM automation | pywin32 |
| OCR fallback | pytesseract + Tesseract |
| Screenshots | mss + Pillow |
| State + memory | SQLite |

---

## License

MIT

---

<div align="center">

**WinScript** — 59 tools. State-aware. Replayable. Audited. Memory-backed.

*Built by [Roshan Ravani](https://github.com/ravaniroshan)*

</div>