#!/bin/bash
# Simple start script for WhatsApp Sender Bot Pro

echo "======================================"
echo "WhatsApp Sender Bot Pro v2.0"
echo "======================================"
echo ""

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "❌ Virtual environment not found!"
    echo ""
    echo "Please run the installation first:"
    echo "  ./install.sh"
    echo ""
    exit 1
fi

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source venv/bin/activate

# Start application
echo "🚀 Starting application..."
echo ""
python main.py
