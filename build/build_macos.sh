#!/bin/bash
# ============================================================
#  Build RetreatPlacer standalone macOS application
#  Run from the project folder: ./build_macos.sh
# ============================================================

set -e

echo ""
echo "========================================"
echo " Building RetreatPlacer for macOS"
echo "========================================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 not found."
    echo "Install it with: brew install python3"
    echo "Or download from https://python.org"
    exit 1
fi

echo "Python: $(python3 --version)"

# Create/activate venv if not already in one
if [ -z "$VIRTUAL_ENV" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
    source .venv/bin/activate
fi

# Install dependencies
echo "Installing dependencies..."
pip install -r requirements.txt
pip install pyinstaller

# Find customtkinter path
CTK_PATH=$(python3 -c "import customtkinter; import os; print(os.path.dirname(customtkinter.__file__))")
echo "CustomTkinter path: $CTK_PATH"

# Determine architecture flag
ARCH=$(uname -m)
echo "Building for architecture: $ARCH"

# Build
echo ""
echo "Running PyInstaller..."
pyinstaller \
    --name "RetreatPlacer" \
    --onedir \
    --windowed \
    --add-data "$CTK_PATH:customtkinter/" \
    --hidden-import ortools \
    --hidden-import ortools.sat \
    --hidden-import ortools.sat.python \
    --hidden-import ortools.sat.python.cp_model \
    --hidden-import openpyxl \
    --hidden-import pandas \
    --hidden-import customtkinter \
    --collect-all ortools \
    --collect-all customtkinter \
    --osx-bundle-identifier "com.retreatplacer.app" \
    ../src/RetreatPlacerUI.py

if [ $? -ne 0 ]; then
    echo ""
    echo "BUILD FAILED. Check errors above."
    exit 1
fi

echo ""
echo "========================================"
echo " BUILD SUCCESSFUL"
echo " Output: dist/RetreatPlacer.app"
echo "========================================"
echo ""

# Show size
if [ -d "dist/RetreatPlacer.app" ]; then
    SIZE=$(du -sh "dist/RetreatPlacer.app" | cut -f1)
    echo " Size: $SIZE"
fi

# Create a zip for easy distribution
echo ""
echo "Creating distributable zip..."
cd dist
zip -r "RetreatPlacer-macOS-${ARCH}.zip" RetreatPlacer.app
echo " Created: dist/RetreatPlacer-macOS-${ARCH}.zip"
cd ..

echo ""
echo "Done! Share the .zip file with users."
echo "First launch: right-click the app → Open (to bypass Gatekeeper)"
