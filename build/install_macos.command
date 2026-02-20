#!/bin/bash
# ============================================================
#  RetreatPlacer — macOS Install & Run Script
#  Lives in: build/install_macos.command
#  Launches: src/RetreatPlacerUI.py
#
#  Double-click this file in Finder to set up and launch.
#  First run takes 2-3 minutes (installs Python packages).
#  Subsequent runs start in seconds.
# ============================================================

# Resolve project root (one level up from build/)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

echo ""
echo " =========================================="
echo "  RetreatPlacer - Room Assignment Solver"
echo " =========================================="
echo ""

# ─── Check for Python 3 ───────────────────────────────────────
if command -v python3 &> /dev/null; then
    PYTHON=python3
elif command -v python &> /dev/null; then
    PYTHON=python
else
    echo " Python 3 is not installed."
    echo ""
    echo " The easiest way to install it on macOS:"
    echo ""
    echo "   Option A — Official installer (recommended):"
    echo "     Opening https://python.org/downloads ..."
    open "https://www.python.org/downloads/"
    echo ""
    echo "   Option B — Homebrew (if you have it):"
    echo "     brew install python3"
    echo ""
    echo " After installing, close this window and double-click"
    echo " this file again."
    echo ""
    read -p " Press Enter to close..."
    exit 1
fi

echo " Python found: $($PYTHON --version)"
echo ""

# ─── Set up virtual environment in project root ───────────────
VENV_DIR="$PROJECT_ROOT/.retreat_venv"

if [ ! -f "$VENV_DIR/bin/python" ]; then
    echo " First run — setting up environment..."
    echo " This takes 2-3 minutes. Please wait."
    echo ""
    
    echo " [1/3] Creating virtual environment..."
    $PYTHON -m venv "$VENV_DIR"
    
    echo " [2/3] Installing solver engine (OR-Tools)..."
    "$VENV_DIR/bin/pip" install --quiet "ortools>=9.8"
    
    echo " [3/3] Installing remaining packages..."
    "$VENV_DIR/bin/pip" install --quiet "openpyxl>=3.1" "pandas>=2.0" "customtkinter>=5.2"
    
    echo ""
    echo " Setup complete!"
else
    echo " Environment already set up."
fi

echo ""
echo " Launching RetreatPlacer..."
echo ""

# ─── Launch the app ───────────────────────────────────────────
"$VENV_DIR/bin/python" "$PROJECT_ROOT/src/RetreatPlacerUI.py"

# If it crashes, keep the terminal open so the user can see the error
if [ $? -ne 0 ]; then
    echo ""
    echo " Something went wrong. Error details above."
    read -p " Press Enter to close..."
fi
