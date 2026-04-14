# WinScript MCP Server - Deployment Guide

Complete guide for deploying WinScript to any environment.

---

## Table of Contents

1. [Quick Comparison](#quick-comparison)
2. [Option 1: Smithery.ai (Cloud Hosted)](#option-1-smitheryai-cloud-hosted)
3. [Option 2: PyPI Package](#option-2-pypi-package)
4. [Option 3: Docker](#option-3-docker)
5. [Option 4: Direct from Source](#option-4-direct-from-source)
6. [Connecting MCP Clients](#connecting-mcp-clients)
7. [Production Considerations](#production-considerations)
8. [Troubleshooting](#troubleshooting)

---

## Quick Comparison

| Method | Best For | Setup Time | Maintenance |
|--------|----------|------------|-------------|
| **Smithery.ai** | AI agent users | 1 click | None (cloud) |
| **PyPI** | Python developers | 2 minutes | Manual updates |
| **Docker** | Sysadmins, isolation | 5 minutes | Docker updates |
| **Source** | Contributors, debugging | 10 minutes | Git pulls |

---

## Option 1: Smithery.ai (Cloud Hosted)

**Best for:** End users who want zero setup

### Steps:

1. Visit: https://smithery.ai/server/winscript
2. Click "Add to Claude" (or your MCP client)
3. Authenticate if prompted
4. Done! 59 tools appear in your AI client

### Pros:
- Zero installation
- No maintenance
- Works immediately
- Automatic updates

### Cons:
- Requires Smithery account
- Cloud-based (some prefer local-only)
- Limited to supported MCP clients

### Supported Clients:
- Claude Desktop
- Cursor
- Windsurf
- Any Smithery-connected MCP client

---

## Option 2: PyPI Package

**Best for:** Python developers, local control

### Installation:

```bash
# Install
pip install winscript

# Verify installation
winscript --help
# or
python -m winscript.server --help

# Run server
winscript
```

### Update:
```bash
pip install --upgrade winscript
```

### Uninstall:
```bash
pip uninstall winscript
```

### Connect to Claude Desktop:

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "winscript": {
      "command": "winscript"
    }
  }
}
```

Or with explicit Python path:

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

### Connect to Cursor:

Edit `.cursor/mcp.json` in your project:

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

---

## Option 3: Docker

**Best for:** Isolation, production, sysadmins

### Quick Start:

```bash
# Pull from GitHub Container Registry
docker run -d --name winscript \
  -v ${HOME}/.winscript:/root/.winscript \
  ghcr.io/roshandamm/winscript-mcp:latest
```

### Windows (PowerShell):

```powershell
docker run -d --name winscript `
  -v $env:USERPROFILE/.winscript:C:\Users\Container\.winscript `
  ghcr.io/roshandamm/winscript-mcp:latest
```

### Build from Source:

```bash
# Clone repo
git clone https://github.com/roshandamm/winscript-mcp.git
cd winscript-mcp

# Build image
docker build -t winscript:latest .

# Run
docker run -d --name winscript -v winscript_data:/root/.winscript winscript:latest
```

### Using Docker Compose:

```bash
# Start
docker-compose up -d

# View logs
docker-compose logs -f winscript

# Stop
docker-compose down
```

### Connect Docker container to MCP client:

Docker containers use stdio transport. Configure your MCP client:

```json
{
  "mcpServers": {
    "winscript": {
      "command": "docker",
      "args": ["exec", "-i", "winscript-mcp", "python", "-m", "winscript.server"]
    }
  }
}
```

### Persistent Data:

All WinScript data (audit logs, memory, workflows) is stored in `/root/.winscript`.

Mount a volume to preserve data:

```bash
docker run -v /path/on/host:/root/.winscript winscript:latest
```

Data structure:
```
~/.winscript/
├── audit.db          # Audit log (auto-purged at 30 days)
├── memory.db         # Window/file/action memory
└── workflows/        # JSON workflow files
```

---

## Option 4: Direct from Source

**Best for:** Contributors, debugging, custom builds

### Steps:

```bash
# 1. Clone repository
git clone https://github.com/roshandamm/winscript-mcp.git
cd winscript-mcp

# 2. Create virtual environment (recommended)
python -m venv venv

# Windows:
venv\Scripts\activate
# Linux/Mac:
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run server
python winscript-server.py
# or
python -m winscript.server
```

### One-Command Start:

```bash
# Windows
start.bat

# Linux/Mac
chmod +x start.sh
./start.sh
```

### Connect to MCP Client:

Point your MCP client to the server script:

```json
{
  "mcpServers": {
    "winscript": {
      "command": "python",
      "args": ["/absolute/path/to/winscript-mcp/winscript/server.py"]
    }
  }
}
```

---

## Connecting MCP Clients

### Claude Desktop

**Config location:** `%APPDATA%\Claude\claude_desktop_config.json`

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

Restart Claude Desktop after editing.

### Cursor

**Config location:** `.cursor/mcp.json` (project-level) or user settings

```json
{
  "mcpServers": {
    "winscript": {
      "command": "python",
      "args": ["-m", "winscript.server"],
      "env": {}
    }
  }
}
```

### Windsurf

**Config location:** `.codeium/windsurf/mcp_config.json`

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

### Custom MCP Client

Any MCP client that supports stdio transport can connect:

```python
from mcp import ClientSession, StdioServerParameters

params = StdioServerParameters(
    command="python",
    args=["-m", "winscript.server"]
)

async with ClientSession(params) as session:
    # List available tools
    tools = await session.list_tools()
    
    # Call a tool
    result = await session.call_tool("open_app", {"name": "notepad"})
```

---

## Production Considerations

### Security

1. **Execution Mode**: Start in safe mode for read-only access
   ```json
   {
     "mcpServers": {
       "winscript": {
         "command": "python",
         "args": ["-m", "winscript.server"],
         "env": {"WINSCRIPT_MODE": "safe"}
       }
     }
   }
   ```

2. **Admin Privileges**: Don't run as Administrator unless automating admin-level apps

3. **Audit Logs**: Review `~/.winscript/audit.db` regularly

### Performance

- **Memory**: ~50-100MB RAM during operation
- **CPU**: Minimal when idle, spikes during UI automation
- **Disk**: ~10MB for code + audit logs (auto-purged at 30 days)

### Reliability

1. **COM Process Cleanup**: Server automatically kills orphaned Excel/Outlook processes on startup

2. **Retry Guard**: Hard stop after 5 consecutive identical failures prevents infinite loops

3. **Error Handling**: All tools return error strings instead of crashing

### Monitoring

Check server health:

```python
# From an MCP client
get_audit_log(10)  # Recent actions
get_failure_report()  # Failure rates
```

---

## Troubleshooting

### "Module not found" errors

```bash
# Reinstall dependencies
pip install -r requirements.txt
```

### "No window found" errors

1. Ensure the app is actually running
2. Try different window title variations
3. Use `get_ui_tree()` to discover element names
4. Fall back to `take_screenshot()` for visual automation

### COM automation fails

1. Ensure Excel/Outlook is installed and licensed
2. Check that no other process is holding COM objects
3. Server auto-kills orphaned processes on startup

### OCR fallback not working

```bash
# Install Tesseract
# Windows: https://github.com/tesseract-ocr/tesseract/wiki
pip install pytesseract
```

### Docker container won't start

```bash
# Check logs
docker logs winscript-mcp

# Remove and recreate
docker rm -f winscript-mcp
docker run -d --name winscript-mcp ...
```

### MCP client doesn't see tools

1. Restart the MCP client completely
2. Check server is actually running (no errors in console)
3. Verify config JSON is valid
4. Check absolute paths are correct

---

## Environment Variables

| Variable | Purpose | Default |
|----------|---------|---------|
| `WINSCRIPT_MODE` | Set execution mode | `standard` |
| `PYTHONUNBUFFERED` | Real-time logs | `0` |
| `WINSCRIPT_DATA_DIR` | Override data directory | `~/.winscript` |

---

## Custom Builds

For contributing or customizing:

```bash
# Clone
git clone https://github.com/roshandamm/winscript-mcp.git
cd winscript-mcp

# Edit code...

# Test
python -m pytest tests/ -v

# Build package
pip install build
python -m build

# Install local build
pip install dist/winscript-*.whl
```

---

## Next Steps

- Read the full documentation in README.md
- Explore examples in `examples/` directory
- Check truth.md for architecture decisions
- Contribute: https://github.com/roshandamm/winscript-mcp

---

**Built by Roshan Ravani** · MIT License · April 2026
