@echo off
REM Build WinScript as a Claude Desktop Extension (.mcpb file)
REM 
REM Prerequisites: npm install -g @anthropic-ai/mcpb
REM
REM Usage: Double-click or run: build-extension.bat

echo.
echo ==========================================
echo   WinScript - Claude Desktop Extension Builder
echo ==========================================
echo.

REM Check if mcpb is installed
where mcpb >nul 2>&1
if %errorlevel% neq 0 (
    echo mcpb CLI not found. Installing...
    call npm install -g @anthropic-ai/mcpb
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install mcpb.
        echo Make sure Node.js and npm are installed.
        pause
        exit /b 1
    )
    echo ✓ mcpb installed
)

REM Check if manifest.json exists
if not exist "manifest.json" (
    echo ERROR: manifest.json not found
    pause
    exit /b 1
)

echo 1/4 Preparing server bundle...
echo.

REM Create build directory
if exist "extension-build" rmdir /s /q extension-build
mkdir extension-build\server

REM Copy server code
xcopy /e /i /y winscript extension-build\server\winscript >nul
copy /y winscript-server.py extension-build\server\ >nul
copy /y requirements.txt extension-build\server\ >nul

REM Copy manifest
copy /y manifest.json extension-build\ >nul

REM Copy icon if exists
if exist "icon.png" (
    copy /y icon.png extension-build\ >nul
    echo ✓ Icon included
) else (
    echo ⚠ No icon.png found (optional)
)

echo ✓ Server bundle prepared
echo.

echo 2/4 Packing extension...
cd extension-build
call mcpb pack
cd ..

if %errorlevel% neq 0 (
    echo ERROR: mcpb pack failed
    pause
    exit /b 1
)

echo ✓ Extension packed
echo.

if exist "extension-build\winscript.mcpb" (
    echo 3/4 Moving to output directory...
    move /y extension-build\winscript.mcpb . >nul
    echo ✓ winscript.mcpb created
    echo.
    
    echo 4/4 Build complete!
    echo.
    echo ==========================================
    echo   ✓ WinScript Extension Built
    echo ==========================================
    echo.
    echo   File: winscript.mcpb
    echo.
    echo   To install:
    echo   1. Double-click winscript.mcpb
    echo   2. Claude Desktop opens it
    echo   3. Review permissions → Install
    echo.
    echo   To distribute:
    echo   1. Submit to Anthropic
    echo   2. After review → appears in Claude's Extensions directory
    echo.
    echo ==========================================
    echo.
    
    REM Clean up
    rmdir /s /q extension-build
    
    pause
) else (
    echo ERROR: Build failed - .mcpb file not created
    pause
    exit /b 1
)
