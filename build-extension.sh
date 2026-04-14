#!/bin/bash
# Build WinScript as a Claude Desktop Extension (.mcpb file)
# 
# Prerequisites: npm install -g @anthropic-ai/mcpb
#
# Usage: ./build-extension.sh
# Output: winscript.mcpb

set -e

echo ""
echo "=========================================="
echo "  WinScript - Claude Desktop Extension Builder"
echo "=========================================="
echo ""

# Check if mcpb is installed
if ! command -v mcpb &> /dev/null; then
    echo "❌ mcpb CLI not found. Installing..."
    npm install -g @anthropic-ai/mcpb
    echo "✓ mcpb installed"
fi

# Check if manifest.json exists
if [ ! -f "manifest.json" ]; then
    echo "❌ manifest.json not found in current directory"
    exit 1
fi

echo "📦 Building WinScript extension..."
echo ""

# Create server directory structure
echo "1/4 Preparing server bundle..."
mkdir -p extension-build/server

# Copy server code
cp -r winscript extension-build/server/
cp winscript-server.py extension-build/server/
cp requirements.txt extension-build/server/

# Copy manifest and icon
cp manifest.json extension-build/
if [ -f "icon.png" ]; then
    cp icon.png extension-build/
    echo "✓ Icon included"
else
    echo "⚠ No icon.png found (optional)"
fi

echo "✓ Server bundle prepared"
echo ""

# Build the .mcpb file
echo "2/4 Packing extension..."
cd extension-build
mcpb pack
cd ..

echo "✓ Extension packed"
echo ""

# Verify output
if [ -f "extension-build/winscript.mcpb" ]; then
    echo "3/4 Moving to output directory..."
    mv extension-build/winscript.mcpb ./winscript.mcpb
    echo "✓ winscript.mcpb created"
    echo ""
    
    FILESIZE=$(du -h winscript.mcpb | cut -f1)
    echo "4/4 Build complete!"
    echo ""
    echo "=========================================="
    echo "  ✓ WinScript Extension Built"
    echo "=========================================="
    echo ""
    echo "  File: winscript.mcpb ($FILESIZE)"
    echo ""
    echo "  To install:"
    echo "  1. Double-click winscript.mcpb"
    echo "  2. Claude Desktop opens it"
    echo "  3. Review permissions → Install"
    echo ""
    echo "  To distribute:"
    echo "  1. Submit to Anthropic: https://console.anthropic.com"
    echo "  2. After review → appears in Claude's Extensions directory"
    echo ""
    echo "=========================================="
    echo ""
    
    # Clean up build directory
    rm -rf extension-build
else
    echo "❌ Build failed - .mcpb file not created"
    exit 1
fi
