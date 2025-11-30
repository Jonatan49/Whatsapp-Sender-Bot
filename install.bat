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
    echo ERROR: Python is not installed or not in PATH!
    echo.
    echo Please:
    echo 1. Download Python from: https://www.python.org/downloads/
    echo 2. Run the installer
    echo 3. CHECK THE BOX: "Add Python to PATH"
    echo 4. Complete installation
    echo 5. Restart this script
    echo.
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Create virtual environment
echo Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment!
    echo.
    pause
    exit /b 1
)
echo Virtual environment created successfully.
echo.

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment!
    echo.
    pause
    exit /b 1
)
echo Virtual environment activated.
echo.

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo WARNING: Failed to upgrade pip, but continuing...
    echo.
)

REM Install dependencies
echo.
echo Installing dependencies (this may take 2-3 minutes)...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies!
    echo.
    echo This might be due to:
    echo - Internet connection issues
    echo - Missing build tools
    echo.
    pause
    exit /b 1
)
echo Dependencies installed successfully.
echo.

REM Copy environment file
if not exist .env (
    echo Creating .env file...
    copy .env.example .env
    if errorlevel 1 (
        echo WARNING: Could not copy .env file
    ) else (
        echo .env file created successfully.
    )
) else (
    echo .env file already exists.
)
echo.

REM Initialize database
echo Initializing database...
python -c "from src.core.database import init_db; init_db()"
if errorlevel 1 (
    echo ERROR: Failed to initialize database!
    echo.
    pause
    exit /b 1
)
echo Database initialized successfully.
echo.

echo ==================================
echo Installation Complete!
echo ==================================
echo.
echo Next steps:
echo 1. Double-click start.bat to run the application
echo 2. Login with: admin / admin123
echo 3. Change your password in Settings!
echo.
echo Default login credentials:
echo   Username: admin
echo   Password: admin123
echo.
echo WARNING: Change the default password after first login!
echo.
pause
