@echo off
chcp 65001 >nul
title בוט שליחת הודעות וואצאפ

echo.
echo ╔════════════════════════════════════════════╗
echo ║     בוט שליחת הודעות וואצאפ              ║
echo ║         © Yonatan Cohen                   ║
echo ╚════════════════════════════════════════════╝
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python לא מותקן במחשב
    echo [!] אנא התקן Python מ-https://python.org
    pause
    exit /b 1
)

:: Check and install dependencies
echo [*] בודק תלויות...
pip show PyQt5 >nul 2>&1
if errorlevel 1 (
    echo [*] מתקין תלויות נדרשות - זה עלול לקחת דקה...
    pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo [!] שגיאה בהתקנת התלויות
        pause
        exit /b 1
    )
    echo [+] התלויות הותקנו בהצלחה!
)

echo [*] מפעיל את התוכנה...
echo.
python app.py

if errorlevel 1 (
    echo.
    echo [!] התוכנה נסגרה עם שגיאה
    pause
)
