@echo off
REM WinScript MCP Server - Quick Start Menu
REM Run this to see installation/start options

echo.
echo ========================================
echo   WinScript MCP Server
echo   AppleScript for Windows
echo ========================================
echo.
echo  What do you want to do?
echo.
echo  1. Install for Claude Desktop (recommended)
echo  2. Start server only
echo  3. Exit
echo.

choice /C 123 /M "Select option"

if errorlevel 3 goto end
if errorlevel 2 goto server
if errorlevel 1 goto install

:install
echo.
echo Running Claude Desktop installer...
echo.
python install.py
goto end

:server
echo.
echo [1/2] Checking dependencies...
python -c "from winscript.server import mcp" 2>nul
if %errorlevel% neq 0 (
    echo Installing dependencies...
    pip install -r requirements.txt
)

echo.
echo [2/2] Starting WinScript MCP Server...
echo.
echo Server is running. Connect your MCP client.
echo Press Ctrl+C to stop.
echo.

python -m winscript.server

:end
echo.
echo Done.
pause
