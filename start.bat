@echo off
REM Simple start script for WhatsApp Sender Bot Pro

echo ======================================
echo WhatsApp Sender Bot Pro v2.0
echo ======================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

echo Python found:
python --version
echo.

REM Check if virtual environment exists
if not exist "venv\" (
    echo WARNING: Virtual environment not found!
    echo.
    echo Running installation now...
    echo.
    call install.bat
    if errorlevel 1 (
        echo.
        echo Installation failed! Please check the errors above.
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

REM Start application
echo Starting WhatsApp Bot...
echo.
python main.py

REM If we get here, the program exited - pause to see any errors
echo.
echo.
echo ======================================
echo Program ended. Check for errors above.
echo ======================================
pause
