@echo off
echo ==========================================
echo WhatsApp Bot - One-Click Start
echo ==========================================
echo.
echo Please wait while I set everything up...
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed!
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

REM Delete old venv if exists
if exist "venv\" (
    echo Removing old installation...
    rmdir /s /q venv
)

REM Create venv
echo Creating virtual environment...
python -m venv venv

REM Activate venv
call venv\Scripts\activate.bat

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip --quiet

REM Install packages directly (skip requirements.txt)
echo Installing packages (this takes 2-3 minutes)...
echo Please be patient...
pip install --quiet PyQt5 selenium webdriver-manager pandas openpyxl bcrypt python-dotenv cryptography SQLAlchemy colorlog PyYAML pyperclip tenacity python-dateutil

if errorlevel 1 (
    echo.
    echo ERROR: Installation failed!
    pause
    exit /b 1
)

REM Fix the Boolean import bug
echo Fixing code...
powershell -Command "(Get-Content src\models\campaign.py) -replace 'from sqlalchemy import Column, String, Integer, ForeignKey, Float, JSON, Enum as SQLEnum$', 'from sqlalchemy import Column, String, Integer, ForeignKey, Float, JSON, Enum as SQLEnum, Boolean' | Set-Content src\models\campaign.py"
powershell -Command "(Get-Content src\models\message.py) -replace 'from sqlalchemy import Column, String, Integer, ForeignKey, Float, Enum as SQLEnum, Text$', 'from sqlalchemy import Column, String, Integer, ForeignKey, Float, Enum as SQLEnum, Text, Boolean' | Set-Content src\models\message.py"

REM Create .env if not exists
if not exist ".env" (
    copy .env.example .env >nul 2>&1
)

REM Initialize database
echo Setting up database...
python -c "from src.core.database import init_db; init_db()" >nul 2>&1

echo.
echo ==========================================
echo Setup Complete! Starting application...
echo ==========================================
echo.
echo Login with:
echo   Username: admin
echo   Password: admin123
echo.
echo IMPORTANT: Change your password after login!
echo.
pause

REM Start the app
python main.py

pause
