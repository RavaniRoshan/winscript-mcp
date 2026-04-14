# WinScript - Claude Desktop Extension Guide

How to build, install, and submit WinScript to appear in Claude Desktop's built-in Extensions directory.

---

## What is a Claude Desktop Extension?

A **Claude Desktop Extension** (`.mcpb` file) is a packaged MCP server that users can install with one click in Claude Desktop. It's like a browser extension for Claude.

**The flow:**
```
Developer builds .mcpb → Submits to Anthropic → After review → Appears in Claude's Extensions directory → Users install with 1 click
```

**This is how Desktop Commander and other popular connectors appear in Claude Desktop.**

---

## Quick Start

### 1. Build the Extension

**Windows:**
```bash
# Double-click build-extension.bat
# or
build-extension.bat
```

**Linux/Mac:**
```bash
chmod +x build-extension.sh
./build-extension.sh
```

**Output:** `winscript.mcpb`

### 2. Test Locally

```bash
# Double-click the .mcpb file
# Claude Desktop opens → review → install

# Or drag into Claude Desktop:
# Settings → Extensions → Drag winscript.mcpb onto window
```

### 3. Submit to Anthropic

1. Go to: **https://console.anthropic.com** (or submission form)
2. Submit your `.mcpb` file
3. Wait for review (security + quality check)
4. After approval → appears in Claude's Extensions directory

---

## How Users Install WinScript

### After Approval (In Claude Desktop)

1. Open Claude Desktop
2. Go to **Settings → Extensions**
3. Search for **"WinScript"**
4. Click **Install**
5. Wait ~30 seconds for 59 tools to load
6. Done!

### Manual Install (Before Approval)

1. Download `winscript.mcpb` from GitHub releases
2. Double-click the file
3. Claude Desktop opens it
4. Review permissions → **Install**
5. Done!

---

## Extension Structure

```
winscript.mcpb (ZIP archive)
├── manifest.json          ← Extension metadata and config
├── icon.png              ← Extension icon (512x512)
└── server/               ← MCP server code
    ├── winscript-server.py   ← Entry point
    ├── requirements.txt      ← Python dependencies
    └── winscript/            ← Full server codebase
        ├── __init__.py
        ├── server.py
        ├── tools/
        ├── core/
        └── adapters/
```

---

## manifest.json Explained

```json
{
  "manifest_version": "0.3",
  "name": "winscript",
  "display_name": "WinScript",
  "version": "0.1.0",
  "description": "AppleScript for Windows. Control any Windows app from Claude.",
  "author": {
    "name": "Roshan Ravani",
    "email": "roshan@example.com"
  },
  "server": {
    "type": "python",
    "entry_point": "server/winscript-server.py",
    "mcp_config": {
      "command": "${server.python}",
      "args": ["${__dirname}/server/winscript-server.py"]
    }
  },
  "compatibility": {
    "platforms": ["win32"],
    "runtimes": {
      "python": ">=3.10"
    }
  }
}
```

**Key fields:**
- `server.type`: `"python"` tells Claude to use Python runtime
- `server.entry_point`: Path to server launcher
- `${__dirname}`: Resolves to extension install directory
- `${server.python}`: Uses Claude's bundled Python

---

## Requirements for Submission

Anthropic reviews extensions for:

### ✅ Must Have
- Valid `manifest.json` with all required fields
- Working `.mcpb` package
- Clear description of functionality
- Privacy policy (if handling user data)
- MIT/Apache/open source license

### ✅ WinScript Meets
- ✅ Valid manifest (created)
- ✅ Working server (59 tools tested)
- ✅ Clear description (written)
- ✅ MIT license (included)
- ✅ Privacy: all data stays local (user's machine)

### 📝 Recommended
- Icon (512x512 PNG) - **Need to create**
- Screenshots of functionality - **Optional**
- Long description with examples - **Included**
- Keywords for search - **Included**

---

## Build Process Details

### Step 1: Install mcpb CLI

```bash
npm install -g @anthropic-ai/mcpb
```

### Step 2: Prepare Bundle

The build script does this automatically:
1. Creates `extension-build/server/` directory
2. Copies `winscript/` package
3. Copies `winscript-server.py` launcher
4. Copies `manifest.json`
5. Copies `icon.png` (if exists)

### Step 3: Pack

```bash
cd extension-build
mcpb pack
# Output: winscript.mcpb
```

### Step 4: Test

```bash
# Double-click winscript.mcpb
# Claude Desktop should open it
# Review permissions → Install
# Check that 59 tools appear
```

### Step 5: Submit

Submit to Anthropic for inclusion in the curated directory.

---

## What Happens After Submission

1. **Review** (1-2 weeks): Anthropic checks security, quality, functionality
2. **Approval**: Extension added to curated directory
3. **Discovery**: Users can search "WinScript" in Claude Desktop → Settings → Extensions
4. **Install**: One-click install, no technical knowledge needed
5. **Updates**: Future `.mcpb` versions can be published for updates

---

## Comparison: Desktop Commander vs WinScript

| Feature | Desktop Commander | WinScript |
|---------|------------------|-----------|
| **Domain** | Terminal + filesystem | Full Windows automation |
| **Tools** | 26 | 59 |
| **Protocol** | MCP | MCP |
| **Extension** | `.mcpb` | `.mcpb` (ready to build) |
| **Install** | 1-click in Claude | 1-click after approval |
| **Runtime** | Node.js | Python |

**WinScript complements Desktop Commander:**
- Desktop Commander: file system + terminal
- WinScript: UI automation + Office apps + workflows

**Together: Complete Windows automation.**

---

## Troubleshooting

### mcpb pack fails

```bash
# Check manifest syntax
python -c "import json; json.load(open('manifest.json'))"

# Check mcpb version
mcpb --version
```

### Extension doesn't appear in Claude

1. Check Claude Desktop is updated (Extensions supported in recent versions)
2. Check `.mcpb` was created successfully
3. Try dragging `.mcpb` into Claude Desktop manually

### Server won't start after install

1. Check Python 3.10+ is available
2. Check `requirements.txt` dependencies install
3. Look at Claude Desktop logs for errors

---

## Next Steps

### Immediate
1. [ ] Create `icon.png` (512x512) - use the ASCII art logo
2. [ ] Test `.mcpb` build locally
3. [ ] Test `.mcpb` install in Claude Desktop
4. [ ] Submit to Anthropic

### Long-term
1. [ ] Get approved → appears in Extensions directory
2. [ ] Promote on social media
3. [ ] Gather user feedback
4. [ ] Publish updates via new `.mcpb` versions

---

**Built by Roshan Ravani** · MIT License · April 2026
