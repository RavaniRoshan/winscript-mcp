#!/usr/bin/env python3
"""
WinScript MCP Server - Standalone Launcher

This script allows running the WinScript MCP server without installing the package.
Users can simply download and run: python winscript-server.py

No pip install needed!
"""

import sys
import subprocess
from pathlib import Path

def check_dependencies():
    """Check if required dependencies are installed."""
    required = [
        "fastmcp",
        "pywinauto", 
        "pywin32",
        "uiautomation",
        "mss",
        "PIL",  # Pillow
    ]
    
    missing = []
    for pkg in required:
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    
    if missing:
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", "-r", 
                str(Path(__file__).parent / "requirements.txt")
            ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except subprocess.CalledProcessError:
            sys.exit(1)

def main():
    """Run the WinScript MCP server."""
    check_dependencies()
    
    # Add current directory to Python path
    current_dir = str(Path(__file__).parent)
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)
    
    # Import and run
    from winscript.server import main as server_main
    server_main()

if __name__ == "__main__":
    main()
