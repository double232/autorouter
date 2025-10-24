@echo off
REM Trial Orders Automation - GUI Launcher
REM This script launches the Trial Orders Automation GUI

echo ================================================
echo   Trial Orders Automation - Starting GUI
echo ================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.9 or higher from python.org
    echo.
    pause
    exit /b 1
)

REM Check if required packages are installed
python -c "import msal" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install requirements
        pause
        exit /b 1
    )
    echo.
)

REM Launch the GUI
echo Launching GUI...
echo.
python gui.py

REM If GUI exits with error, show message
if errorlevel 1 (
    echo.
    echo GUI exited with an error
    pause
)
