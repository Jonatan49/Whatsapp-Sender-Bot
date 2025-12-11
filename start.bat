@echo off
setlocal EnableDelayedExpansion
title WhatsApp Sender Bot
color 0A

echo.
echo ========================================
echo        WhatsApp Sender Bot
echo           One-Click Setup
echo ========================================
echo.

REM Check if Python is installed
where python >nul 2>&1
if errorlevel 1 (
    echo [!] Python not found!
    echo.
    echo Please install Python manually:
    echo 1. Go to https://python.org/downloads
    echo 2. Download Python 3.12
    echo 3. Run installer - IMPORTANT: Check "Add to PATH"
    echo 4. Run this file again
    echo.

    REM Try to open Python download page
    start https://www.python.org/downloads/

    echo Opening Python download page...
    echo.
    pause
    exit /b 1
)

echo [OK] Python found:
python --version
echo.

REM Check if we're in the right folder
if not exist "app.py" (
    echo [ERROR] app.py not found!
    echo.
    echo Make sure you run start.bat from the WhatsApp-Sender-Bot folder.
    echo.
    pause
    exit /b 1
)

echo [*] Installing/updating packages...
echo     Please wait...
echo.

REM Upgrade pip first
python -m pip install --upgrade pip --quiet 2>nul

REM Install requirements
if exist "requirements.txt" (
    python -m pip install -r requirements.txt --quiet 2>nul
) else (
    python -m pip install PyQt5 selenium webdriver-manager pandas openpyxl pyperclip --quiet 2>nul
)

echo [OK] Packages ready!
echo.
echo ========================================
echo     Starting WhatsApp Sender Bot...
echo ========================================
echo.

python app.py

echo.
echo ========================================
if errorlevel 1 (
    echo [!] App closed with error
) else (
    echo App closed normally
)
echo Press any key to exit...
pause >nul
