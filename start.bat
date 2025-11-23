@echo off
REM Simple start script for WhatsApp Sender Bot Pro

echo ======================================
echo WhatsApp Sender Bot Pro v2.0
echo ======================================
echo.

REM Check if virtual environment exists
if not exist "venv\" (
    echo X Virtual environment not found!
    echo.
    echo Please run the installation first:
    echo   install.bat
    echo.
    pause
    exit /b 1
)

REM Activate virtual environment
echo Starting application...
call venv\Scripts\activate.bat

REM Start application
python main.py
