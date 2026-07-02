@echo off
setlocal enabledelayedexpansion
title Gaming Log Parser - Auto Installer

echo ========================================
echo   Gaming Log Parser v2.4
echo   Auto-Installer
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH
    echo.
    echo Please install Python 3.8 or higher from:
    echo https://www.python.org/downloads/
    echo.
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo [OK] Python is installed
python --version

REM Install required packages
echo.
echo Installing required packages...
echo.

pip install openpyxl pillow --break-system-packages >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Installation with --break-system-packages failed
    echo Trying without flag...
    pip install openpyxl pillow
)

echo [OK] Packages installed
echo.

REM Launch the program
echo Launching Gaming Log Parser...
echo.
python gaming_log_parser.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Program exited with error code: %errorlevel%
    pause
)
