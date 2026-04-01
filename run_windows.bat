@echo off
echo ========================================
echo   Assignment Project Generator — Setup
echo ========================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please download it from https://python.org and re-run this script.
    pause
    exit /b 1
)

echo [1/2] Installing dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo [2/2] Launching app...
echo.
python assignment_generator.py

pause
