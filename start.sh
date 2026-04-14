#!/bin/bash
# WinScript MCP Server - Linux/WSL Quick Start
# Run this script to install and start WinScript without manual setup

echo ""
echo "========================================"
echo "  WinScript MCP Server Installer"
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Please install Python 3.10+ from your distribution's package manager."
    exit 1
fi

echo "[1/3] Checking Python version..."
python3 -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)"
if [ $? -ne 0 ]; then
    echo "ERROR: Python 3.10+ is required. You have:"
    python3 --version
    exit 1
fi
echo "OK"

echo ""
echo "[2/3] Creating virtual environment..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "Virtual environment created."
else
    echo "Virtual environment already exists."
fi

source venv/bin/activate

echo ""
echo "[3/3] Installing dependencies..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Failed to install dependencies."
    echo "Try running: pip install --upgrade pip"
    exit 1
fi
echo "OK"

echo ""
echo "Starting WinScript MCP Server..."
echo ""
echo "Server is starting. Connect your MCP client to it."
echo "Press Ctrl+C to stop the server."
echo ""

python -m winscript.server
