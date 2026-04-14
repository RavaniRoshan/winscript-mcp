# WinScript - Quick Reference Card

## For Users

### Add to AI Agent (Easiest)
👉 https://smithery.ai/server/winscript

### Install via PyPI
```bash
pip install winscript
winscript
```

### Run with Docker
```bash
docker run -v %USERPROFILE%/.winscript:~/.winscript ghcr.io/roshandamm/winscript-mcp:latest
```

### Run from Source
```bash
git clone https://github.com/roshandamm/winscript-mcp.git
cd winscript-mcp
python winscript-server.py
```

---

## For Developers

### Build Package
```bash
pip install build
python -m build
```

### Install Locally
```bash
pip install -e .
winscript
```

### Run Tests
```bash
pip install pytest
pytest tests/ -v
```

---

## Connect to MCP Clients

### Claude Desktop
`%APPDATA%\Claude\claude_desktop_config.json`:
```json
{"mcpServers": {"winscript": {"command": "winscript"}}}
```

### Cursor
`.cursor/mcp.json`:
```json
{"mcpServers": {"winscript": {"command": "python", "args": ["-m", "winscript.server"]}}}
```

---

## What You Get

**59 tools** across 15 categories:
- App Control (5)
- UI Interaction (6)
- COM Office (5)
- File System (7)
- Screen + Clipboard (4)
- PowerShell (1)
- App Adapters (15)
- Workflow Recorder (6)
- Semantic Intents (5)
- Audit + Memory (5)
- Execution Modes (2)
- State (1)

---

## Data Location
`~/.winscript/`
- `audit.db` - All actions (auto-purged at 30 days)
- `memory.db` - Window/file history
- `workflows/` - JSON workflow files

---

**Built by Roshan Ravani** · MIT License
