@echo off
REM ============================================================
REM  RetreatPlacer - Windows Install and Run Script
REM  Lives in: build/install_windows.bat
REM  Launches: src/RetreatPlacerUI.py
REM
REM  Double-click this file to set up and launch RetreatPlacer.
REM  First run takes 2-3 minutes (installs Python packages).
REM  Subsequent runs start in seconds.
REM ============================================================

title RetreatPlacer Setup

REM Resolve project root (one level up from build/)
set "PROJECT_ROOT=%~dp0.."

echo.
echo  ==========================================
echo   RetreatPlacer - Room Assignment Solver
echo  ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 goto :nopython
goto :haspython

:nopython
echo  Python is not installed.
echo.
echo  Opening the Python download page...
echo  IMPORTANT: Check "Add Python to PATH" during installation!
echo.
start https://www.python.org/downloads/
echo  After installing Python, close this window and double-click
echo  this file again.
echo.
pause
exit /b 1

:haspython
echo  Python found:
python --version
echo.

REM Set up virtual environment in project root
set "VENV_DIR=%PROJECT_ROOT%\.retreat_venv"

if exist "%VENV_DIR%\Scripts\python.exe" goto :venvready

echo  First run - setting up environment...
echo  This takes 2-3 minutes. Please wait.
echo.

echo  [1/3] Creating virtual environment...
python -m venv "%VENV_DIR%"
if errorlevel 1 goto :venverror

echo  [2/3] Installing solver engine - OR-Tools...
"%VENV_DIR%\Scripts\pip.exe" install --quiet "ortools>=9.8"
if errorlevel 1 goto :installerror

echo  [3/3] Installing remaining packages...
"%VENV_DIR%\Scripts\pip.exe" install --quiet "openpyxl>=3.1" "pandas>=2.0" "customtkinter>=5.2"
if errorlevel 1 goto :installerror

echo.
echo  Setup complete!
goto :launch

:venvready
echo  Environment already set up.
goto :launch

:venverror
echo.
echo  ERROR: Failed to create virtual environment.
echo  Make sure Python is installed correctly with pip.
pause
exit /b 1

:installerror
echo.
echo  ERROR: Failed to install packages.
echo  Check your internet connection and try again.
pause
exit /b 1

:launch
echo.
echo  Launching RetreatPlacer...
echo.

"%VENV_DIR%\Scripts\pythonw.exe" "%PROJECT_ROOT%\src\RetreatPlacerUI.py"
if errorlevel 1 goto :launcherror
goto :end

:launcherror
echo.
echo  Something went wrong. Trying with console output...
"%VENV_DIR%\Scripts\python.exe" "%PROJECT_ROOT%\src\RetreatPlacerUI.py"
pause

:end