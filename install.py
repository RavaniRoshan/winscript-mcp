"""
WinScript MCP Server - Claude Desktop Auto-Installer

This script:
1. Installs WinScript dependencies
2. Auto-configures Claude Desktop's claude_desktop_config.json
3. Restarts Claude Desktop (optional)
4. WinScript appears as a connector with 59 tools

Run: python install.py
"""

import sys
import os
import json
import subprocess
import shutil
from pathlib import Path

def print_header():
    print()
    print("=" * 60)
    print("  WinScript MCP Server - Claude Desktop Installer")
    print()
    print("  AppleScript for Windows. Built for AI agents.")
    print("=" * 60)
    print()

def check_python():
    """Check Python 3.10+ is installed."""
    print("[1/5] Checking Python...")
    
    try:
        version = sys.version_info
        if version.major < 3 or (version.major == 3 and version.minor < 10):
            print(f"  ❌ Python 3.10+ required. You have {version.major}.{version.minor}")
            print("  Download: https://www.python.org/downloads/")
            return False
        
        print(f"  ✓ Python {version.major}.{version.minor}.{version.micro}")
        return True
    except Exception:
        print("  ❌ Python not found")
        return False

def install_dependencies():
    """Install pip dependencies."""
    print()
    print("[2/5] Installing dependencies...")
    
    try:
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"],
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if result.returncode == 0:
            print("  ✓ Dependencies installed")
            return True
        else:
            print(f"  ❌ Installation failed:")
            print(result.stderr)
            return False
    except subprocess.TimeoutExpired:
        print("  ❌ Installation timed out (5 min limit)")
        return False
    except Exception as e:
        print(f"  ❌ Error: {e}")
        return False

def get_claude_config_path():
    """Find Claude Desktop config location."""
    appdata = os.environ.get("APPDATA")
    if not appdata:
        return None
    
    return Path(appdata) / "Claude" / "claude_desktop_config.json"

def configure_claude_desktop():
    """Add WinScript to Claude Desktop config."""
    print()
    print("[3/5] Configuring Claude Desktop...")
    
    config_path = get_claude_config_path()
    if not config_path:
        print("  ⚠ Could not find APPDATA directory")
        print("  ⚠ Claude Desktop may not be installed")
        print("  ⚠ You can configure manually later")
        return False
    
    # Create directory if needed
    config_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Load existing config or create new
    config = {}
    if config_path.exists():
        try:
            config = json.loads(config_path.read_text(encoding="utf-8"))
        except Exception:
            print("  ⚠ Existing config is invalid, creating new")
    
    # Initialize mcpServers if needed
    if "mcpServers" not in config:
        config["mcpServers"] = {}
    
    # Add WinScript
    winscript_config = {
        "command": sys.executable,
        "args": ["-m", "winscript.server"],
        "env": {}
    }
    
    if config["mcpServers"].get("winscript") == winscript_config:
        print("  ✓ WinScript already configured")
        return True
    
    config["mcpServers"]["winscript"] = winscript_config
    
    # Backup existing config
    if config_path.exists():
        backup = config_path.with_suffix(".json.backup")
        shutil.copy2(config_path, backup)
        print(f"  ✓ Backed up config to {backup.name}")
    
    # Write new config
    config_path.write_text(
        json.dumps(config, indent=2),
        encoding="utf-8"
    )
    
    print(f"  ✓ Added WinScript to claude_desktop_config.json")
    print(f"    Location: {config_path}")
    
    return True

def check_claude_running():
    """Check if Claude Desktop is currently running."""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Claude.exe"],
            capture_output=True,
            text=True
        )
        return "Claude.exe" in result.stdout
    except Exception:
        return False

def offer_restart_claude():
    """Offer to restart Claude Desktop."""
    print()
    print("[4/5] Checking Claude Desktop...")
    
    if not check_claude_running():
        print("  ℹ Claude Desktop is not running")
        print("  ℹ Start it to see WinScript tools")
        return
    
    print("  ⚠ Claude Desktop is currently running")
    print()
    print("  To activate WinScript, you need to restart Claude Desktop.")
    print()
    
    try:
        choice = input("  Restart Claude Desktop now? (y/n): ").strip().lower()
        if choice in ["y", "yes"]:
            # Kill Claude
            subprocess.run(["taskkill", "/F", "/IM", "Claude.exe"], 
                         capture_output=True)
            print("  ✓ Claude Desktop closed")
            
            # Restart it
            claude_path = Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Claude" / "Claude.exe"
            if claude_path.exists():
                subprocess.Popen([str(claude_path)])
                print("  ✓ Claude Desktop restarted")
                print("  ℹ Wait ~30 seconds for tools to load")
            else:
                print("  ℹ Please restart Claude Desktop manually")
        else:
            print("  ℹ Please restart Claude Desktop manually when ready")
    except Exception:
        print("  ℹ Please restart Claude Desktop manually")

def verify_installation():
    """Verify WinScript can actually start."""
    print()
    print("[5/5] Verifying installation...")
    
    try:
        # Quick import test
        result = subprocess.run(
            [sys.executable, "-c", "from winscript.server import mcp; print('OK')"],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if "OK" in result.stdout:
            print("  ✓ WinScript MCP server loads successfully")
            return True
        else:
            print(f"  ⚠ Verification failed:")
            print(result.stderr)
            return False
    except Exception as e:
        print(f"  ⚠ Could not verify: {e}")
        return False

def print_success():
    """Print success message."""
    print()
    print("=" * 60)
    print("  ✓ Installation Complete!")
    print("=" * 60)
    print()
    print("  What happens next:")
    print()
    print("  1. Open Claude Desktop (or restart if already open)")
    print("  2. Wait ~30 seconds for tools to load")
    print("  3. Ask Claude to:")
    print()
    print("     \"Open Notepad and type Hello from WinScript\"")
    print()
    print("  4. 59 tools appear in Claude's Extensions panel:")
    print("     - App Control (open, close, focus apps)")
    print("     - UI Automation (click, type, read elements)")
    print("     - Excel/Outlook COM automation")
    print("     - File system operations")
    print("     - Screenshots + clipboard")
    print("     - Workflow recording + replay")
    print("     - Semantic intents")
    print("     - Audit logging + memory")
    print()
    print("  Where data is stored:")
    print("  %USERPROFILE%\\.winscript\\")
    print("  - audit.db (all actions, auto-purged at 30 days)")
    print("  - memory.db (window/file history)")
    print("  - workflows/ (JSON workflow files)")
    print()
    print("  To update:")
    print("  pip install --upgrade winscript")
    print()
    print("  For help:")
    print("  https://github.com/RavaniRoshan/winscript-mcp")
    print()
    print("=" * 60)
    print()

def main():
    """Main installation flow."""
    print_header()
    
    # Check Python
    if not check_python():
        print()
        print("  Please install Python 3.10+ and try again.")
        input("\n  Press Enter to exit...")
        sys.exit(1)
    
    # Install dependencies
    if not install_dependencies():
        print()
        print("  Installation failed. Check the errors above.")
        input("\n  Press Enter to exit...")
        sys.exit(1)
    
    # Configure Claude Desktop
    configure_claude_desktop()
    
    # Verify
    verify_installation()
    
    # Offer to restart Claude
    offer_restart_claude()
    
    # Success
    print_success()

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n  Installation cancelled by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n\n  ❌ Unexpected error: {e}")
        sys.exit(1)
