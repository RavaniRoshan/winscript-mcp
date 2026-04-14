@echo off
REM WinScript MCP Server - Windows Quick Start
REM Run this script to install and start WinScript without manual setup

echo.
echo ========================================
echo   WinScript MCP Server Installer
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python 3.10+ is not installed.
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Checking Python version...
python -c "import sys; exit(0 if sys.version_info >= (3, 10) else 1)"
if %errorlevel% neq 0 (
    echo ERROR: Python 3.10+ is required. You have:
    python --version
    pause
    exit /b 1
)
echo OK

echo.
echo [2/3] Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    echo Try running: pip install --upgrade pip
    pause
    exit /b 1
)
echo OK

echo.
echo [3/3] Starting WinScript MCP Server...
echo.
echo Server is starting. Connect your MCP client to it.
echo Press Ctrl+C to stop the server.
echo.

python -m winscript.server
