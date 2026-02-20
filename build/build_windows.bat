@echo off
REM ============================================================
REM  Build RetreatPlacer standalone Windows executable
REM  Lives in: build/build_windows.bat
REM  Builds:   src/RetreatPlacerUI.py (which imports src/RetreatPlacer.py)
REM  Output:   dist/RetreatPlacer.exe
REM
REM  Run from the build/ folder: build_windows.bat
REM ============================================================

REM Resolve project root (one level up from build/)
set "PROJECT_ROOT=%~dp0.."

echo.
echo ========================================
echo  Building RetreatPlacer for Windows
echo ========================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 goto :nopython
goto :haspython

:nopython
echo ERROR: Python not found. Install Python 3.10+ from python.org
pause
exit /b 1

:haspython

REM Create/activate venv if not already in one
if defined VIRTUAL_ENV goto :skipvenv
echo Creating virtual environment...
python -m venv "%PROJECT_ROOT%\.build_venv"
call "%PROJECT_ROOT%\.build_venv\Scripts\activate.bat"

:skipvenv

REM Install dependencies
echo Installing dependencies...
pip install -r "%~dp0requirements.txt"
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
    --distpath "%PROJECT_ROOT%\dist" ^
    --workpath "%PROJECT_ROOT%\build\pyinstaller_work" ^
    --specpath "%PROJECT_ROOT%\build" ^
    --add-data "%CTK_PATH%;customtkinter/" ^
    --add-data "%PROJECT_ROOT%\src\RetreatPlacer.py;." ^
    --hidden-import ortools ^
    --hidden-import ortools.sat ^
    --hidden-import ortools.sat.python ^
    --hidden-import ortools.sat.python.cp_model ^
    --hidden-import openpyxl ^
    --hidden-import pandas ^
    --hidden-import customtkinter ^
    --collect-all ortools ^
    --collect-all customtkinter ^
    "%PROJECT_ROOT%\src\RetreatPlacerUI.py"

if errorlevel 1 goto :buildfailed
goto :buildsuccess

:buildfailed
echo.
echo BUILD FAILED. Check errors above.
pause
exit /b 1

:buildsuccess
echo.
echo ========================================
echo  BUILD SUCCESSFUL
echo  Output: dist\RetreatPlacer.exe
echo ========================================
echo.

if exist "%PROJECT_ROOT%\dist\RetreatPlacer.exe" (
    for %%A in ("%PROJECT_ROOT%\dist\RetreatPlacer.exe") do echo  Size: %%~zA bytes
)

pause
