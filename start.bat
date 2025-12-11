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
python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python not found - Installing automatically...
    echo.

    REM Download Python installer using PowerShell
    echo [*] Downloading Python 3.12...
    powershell -Command "& {Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.0/python-3.12.0-amd64.exe' -OutFile '%TEMP%\python_installer.exe'}" 2>nul

    if not exist "%TEMP%\python_installer.exe" (
        echo.
        echo [ERROR] Failed to download Python
        echo Please download manually from https://python.org
        echo.
        pause
        exit /b 1
    )

    echo [*] Installing Python - this may take a minute...
    echo     DO NOT close this window!
    echo.

    REM Install Python silently with pip and add to PATH
    "%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1

    REM Wait for installation
    timeout /t 10 /nobreak >nul

    REM Refresh PATH
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"

    REM Verify installation
    python --version >nul 2>&1
    if errorlevel 1 (
        echo.
        echo [!] Python installed but PATH not updated
        echo [!] Please RESTART your computer and run this again
        echo.
        pause
        exit /b 1
    )

    echo [OK] Python installed successfully!
    echo.

    REM Cleanup
    del "%TEMP%\python_installer.exe" >nul 2>&1
)

echo [OK] Python found:
python --version
echo.

REM Check if requirements.txt exists
if not exist "requirements.txt" (
    echo [ERROR] requirements.txt not found!
    echo Make sure you are running from the correct folder.
    echo.
    pause
    exit /b 1
)

REM Check if app.py exists
if not exist "app.py" (
    echo [ERROR] app.py not found!
    echo Make sure you are running from the correct folder.
    echo.
    pause
    exit /b 1
)

echo [*] Installing required packages...
echo     This may take a few minutes on first run...
echo.

python -m pip install --upgrade pip -q 2>nul
python -m pip install -r requirements.txt -q 2>nul

if errorlevel 1 (
    echo [!] Retrying installation...
    python -m pip install PyQt5 selenium webdriver-manager pandas openpyxl pyperclip -q 2>nul
)

echo.
echo [OK] Setup complete!
echo.
echo ========================================
echo     Starting WhatsApp Sender Bot...
echo ========================================
echo.

REM Run the app and capture errors
python app.py 2>&1

REM If we get here, the app closed
echo.
echo ========================================
echo App closed. Press any key to exit...
pause >nul
