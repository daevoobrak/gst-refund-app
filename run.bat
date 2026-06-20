@echo off
setlocal EnableDelayedExpansion

title GST Refund Working Generator

echo =========================================
echo  GST Refund Working Generator
echo =========================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo.
    echo Please install Python 3.8 or later from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, tick the box
    echo   "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

:: Show Python version
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo Found: %PYVER%
echo.

:: Create virtual environment if it doesn't exist
if not exist ".venv\" (
    echo Setting up virtual environment (first run only)...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Could not create virtual environment.
        pause
        exit /b 1
    )
)

:: Activate venv
call .venv\Scripts\activate.bat

:: Install / upgrade dependencies
echo Installing dependencies...
pip install --quiet --upgrade flask "pdfplumber==0.10.4" openpyxl 2>nul
if errorlevel 1 (
    echo Retrying with verbose output...
    pip install flask "pdfplumber==0.10.4" openpyxl
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies.
        pause
        exit /b 1
    )
)

echo.
echo Starting GST Refund Working Generator...
echo Open your browser at:  http://localhost:5050
echo.
echo Press Ctrl+C to stop the server.
echo.

:: Open browser after a short delay (run in background)
start "" cmd /c "timeout /t 2 >nul && start http://localhost:5050"

:: Start Flask
python app.py

pause
