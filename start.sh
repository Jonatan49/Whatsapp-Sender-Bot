#!/bin/bash

echo ""
echo "╔════════════════════════════════════════════╗"
echo "║     בוט שליחת הודעות וואצאפ              ║"
echo "║         © Yonatan Cohen                   ║"
echo "╚════════════════════════════════════════════╝"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "[!] Python3 לא מותקן"
    echo "[!] אנא התקן Python3"
    exit 1
fi

# Navigate to script directory
cd "$(dirname "$0")"

# Check and install dependencies
echo "[*] בודק תלויות..."
if ! python3 -c "import PyQt5" &> /dev/null; then
    echo "[*] מתקין תלויות נדרשות..."
    pip3 install -r requirements.txt --quiet
    if [ $? -ne 0 ]; then
        echo "[!] שגיאה בהתקנת התלויות"
        exit 1
    fi
    echo "[+] התלויות הותקנו בהצלחה!"
fi

echo "[*] מפעיל את התוכנה..."
echo ""
python3 app.py
