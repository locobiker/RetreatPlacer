@echo off
REM ============================================================
REM  Build RetreatPlacer standalone Windows executable
REM  Run this from the project folder: build_windows.bat
REM ============================================================

echo.
echo ========================================
echo  Building RetreatPlacer for Windows
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.10+ from python.org
    pause
    exit /b 1
)

REM Create/activate venv if not already in one
if not defined VIRTUAL_ENV (
    echo Creating virtual environment...
    python -m venv .venv
    call .venv\Scripts\activate.bat
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
pip install pyinstaller

REM Find customtkinter location for --add-data
for /f "tokens=*" %%i in ('python -c "import customtkinter; import os; print(os.path.dirname(customtkinter.__file__))"') do set CTK_PATH=%%i

echo CustomTkinter path: %CTK_PATH%

REM Build with PyInstaller
echo.
echo Running PyInstaller...
pyinstaller ^
    --name "RetreatPlacer" ^
    --onefile ^
    --windowed ^
    --add-data "%CTK_PATH%;customtkinter/" ^
    --hidden-import ortools ^
    --hidden-import ortools.sat ^
    --hidden-import ortools.sat.python ^
    --hidden-import ortools.sat.python.cp_model ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    --hidden-import customtkinter ^
    --collect-all ortools ^
    --collect-all customtkinter ^
    ..\src\RetreatPlacerUI.py

if errorlevel 1 (
    echo.
    echo BUILD FAILED. Check errors above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  BUILD SUCCESSFUL
echo  Output: dist\RetreatPlacer.exe
echo ========================================
echo.

REM Verify the exe exists
if exist "dist\RetreatPlacer.exe" (
    for %%A in ("dist\RetreatPlacer.exe") do echo  Size: %%~zA bytes
)

pause
