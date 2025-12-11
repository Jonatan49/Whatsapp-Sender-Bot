@echo off
chcp 65001 >nul 2>&1
title WhatsApp Sender Bot

echo.
echo ========================================
echo        WhatsApp Sender Bot
echo        (C) Yonatan Cohen
echo ========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed!
    echo Please install Python from https://python.org
    echo.
    pause
    exit /b 1
)

echo [*] Checking dependencies...

REM Install dependencies
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    echo Try running: pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

echo [OK] Dependencies ready!
echo.
echo [*] Starting application...
echo.

python app.py

if errorlevel 1 (
    echo.
    echo [ERROR] Application closed with error
    pause
)
