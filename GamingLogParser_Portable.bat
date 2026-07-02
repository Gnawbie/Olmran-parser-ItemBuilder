@echo off
setlocal enabledelayedexpansion
title Gaming Log Parser

:: ════════════════════════════════════════════════════════════════════
::  Gaming Log Parser - Portable Launcher
::  Launches without any visible console window
:: ════════════════════════════════════════════════════════════════════

:: Find Python
set "PYTHON="
set "PYTHONW="

for %%C in (python python3 py) do (
    if not defined PYTHON (
        where %%C >nul 2>&1
        if !errorlevel! == 0 set "PYTHON=%%C"
    )
)

:: Search common install locations if not in PATH
if not defined PYTHON (
    for %%D in (
        "%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
        "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    ) do (
        if not defined PYTHON if exist "%%~D" set "PYTHON=%%~D"
    )
)

:: If Python not found, show error
if not defined PYTHON (
    echo.
    echo ERROR: Python not found!
    echo.
    echo Please run: GamingLogParser_Portable_DEBUG.bat
    echo to install Python automatically.
    echo.
    pause
    exit /b 1
)

:: Find pythonw (GUI version)
for %%F in ("%PYTHON%") do (
    set "PYTHON_DIR=%%~dpF"
    if exist "!PYTHON_DIR!pythonw.exe" set "PYTHONW=!PYTHON_DIR!pythonw.exe"
)

:: Quick dependency check
%PYTHON% -c "import tkinter, openpyxl" >nul 2>&1
if !errorlevel! neq 0 (
    echo.
    echo ERROR: Missing dependencies!
    echo.
    echo Please run: GamingLogParser_Portable_DEBUG.bat
    echo to install required packages.
    echo.
    pause
    exit /b 1
)

:: Create a VBScript to launch without console window
set "VBS_SCRIPT=%TEMP%\launch_glp.vbs"
set "SCRIPT_PATH=%~dp0gaming_log_parser.py"

if defined PYTHONW (
    set "PYTHON_EXE=%PYTHONW%"
) else (
    set "PYTHON_EXE=%PYTHON%"
)

:: Write VBS launcher script
(
echo Set WshShell = CreateObject("WScript.Shell"^)
echo WshShell.Run """%PYTHON_EXE%"" ""%SCRIPT_PATH%""", 0, False
) > "%VBS_SCRIPT%"

:: Launch via VBScript (this prevents any console window)
cscript //nologo "%VBS_SCRIPT%"

:: Clean up VBS script
del "%VBS_SCRIPT%" 2>nul

:: Exit launcher
exit
