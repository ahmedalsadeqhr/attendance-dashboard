@echo off
cd /d "%~dp0"
echo ========================================
echo  Creating Standalone .EXE Application
echo ========================================
echo.

REM Check if PyInstaller is installed
python -m pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller not found. Installing...
    python -m pip install pyinstaller
)

REM Also ensure xlrd is installed for .xls file support
python -m pip install xlrd >nul 2>&1

echo.
echo Creating standalone .EXE file...
echo This may take 2-3 minutes...
echo.

REM Use python -m PyInstaller to avoid PATH issues
python -m PyInstaller --onefile --windowed --name "AttendanceDashboard" attendance_dashboard_app.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Failed to create .EXE file
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  SUCCESS!
echo ========================================
echo.
echo Your standalone application is ready!
echo.
echo Location: dist\AttendanceDashboard.exe
echo.
echo You can now:
echo  1. Copy AttendanceDashboard.exe to any Windows computer
echo  2. Double-click to run (no installation needed)
echo  3. Distribute to users
echo.
echo The .EXE file is completely standalone!
echo.
pause

REM Open the dist folder
explorer dist
