@echo off
cd /d "%~dp0"
echo ========================================
echo  Attendance Dashboard - Quick Start
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed!
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found!
echo.
echo Installing required libraries...
echo.

python -m pip install pandas openpyxl xlsxwriter xlrd

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to install libraries
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Installation Complete!
echo  Launching Attendance Dashboard...
echo ========================================
echo.

python attendance_dashboard_app.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to launch application
    echo.
    pause
    exit /b 1
)
