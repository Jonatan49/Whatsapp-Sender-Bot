@echo off
echo ==========================================
echo Fixing Installation
echo ==========================================
echo.

echo Step 1: Deleting old virtual environment...
if exist "venv\" (
    rmdir /s /q venv
    echo Old venv deleted.
) else (
    echo No old venv found.
)
echo.

echo Step 2: Running full installation...
echo.
call install.bat

echo.
echo ==========================================
echo Done! Now you can run start.bat
echo ==========================================
pause
