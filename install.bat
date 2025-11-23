@echo off
REM Installation script for WhatsApp Sender Bot Pro (Windows)

echo ==================================
echo WhatsApp Sender Bot Pro v2.0.0
echo Installation Script (Windows)
echo ==================================
echo.

REM Check Python version
echo Checking Python version...
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python 3.8 or higher.
    pause
    exit /b 1
)

REM Create virtual environment
echo.
echo Creating virtual environment...
python -m venv venv

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Upgrade pip
echo.
echo Upgrading pip...
python -m pip install --upgrade pip

REM Install dependencies
echo.
echo Installing dependencies...
pip install -r requirements.txt

REM Copy environment file
echo.
if not exist .env (
    echo Creating .env file...
    copy .env.example .env
    echo + .env file created. Please edit it with your settings.
) else (
    echo + .env file already exists.
)

REM Initialize database
echo.
echo Initializing database...
python -c "from src.core.database import init_db; init_db()"

echo.
echo ==================================
echo Installation Complete!
echo ==================================
echo.
echo Next steps:
echo 1. Edit .env file with your settings
echo 2. Run: venv\Scripts\activate.bat
echo 3. Run: python main.py
echo.
echo Default login credentials:
echo   Username: admin
echo   Password: admin123
echo.
echo WARNING: Change the default password after first login!
echo.
pause
