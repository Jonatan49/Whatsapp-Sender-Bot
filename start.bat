@echo off
title WhatsApp Sender Bot
color 0A
cd /d "%~dp0"

echo.
echo ========================================
echo        WhatsApp Sender Bot
echo ========================================
echo.

REM Use py launcher which is more reliable on Windows
py --version >nul 2>&1
if errorlevel 1 (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found!
        echo.
        echo Install Python from: https://python.org/downloads
        echo Make sure to check "Add to PATH" during installation!
        echo.
        start https://www.python.org/downloads/
        pause
        exit /b 1
    )
    set PYTHON_CMD=python
) else (
    set PYTHON_CMD=py
)

echo [OK] Python found
%PYTHON_CMD% --version
echo.

if not exist "app.py" (
    echo [ERROR] app.py not found!
    echo Run this from the WhatsApp-Sender-Bot folder.
    pause
    exit /b 1
)

REM Check for problematic Lib folder
if exist "Lib" (
    echo [WARNING] Found 'Lib' folder that may cause issues.
    echo Renaming to 'Lib_backup'...
    ren "Lib" "Lib_backup" 2>nul
)

echo [*] Installing packages...
%PYTHON_CMD% -m pip install --upgrade pip -q 2>nul
%PYTHON_CMD% -m pip install PyQt5 selenium webdriver-manager pandas openpyxl pyperclip -q 2>nul

echo [OK] Ready!
echo.
echo ========================================
echo     Starting Application...
echo ========================================
echo.

%PYTHON_CMD% app.py

echo.
echo ========================================
echo Press any key to close...
pause >nul
