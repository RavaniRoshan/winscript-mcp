@echo off
REM WinScript MCP Server - One-Click Installer for Claude Desktop
REM Double-click this file to install and configure WinScript

echo.
echo ========================================
echo   WinScript MCP Server Installer
echo   for Claude Desktop
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python 3.10+ is not installed.
    echo.
    echo Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Run installer
python install.py

if %errorlevel% neq 0 (
    echo.
    echo Installation failed. Check the errors above.
    pause
    exit /b 1
)

pause
