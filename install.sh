#!/bin/bash
# Installation script for WhatsApp Sender Bot Pro

set -e  # Exit on error

echo "=================================="
echo "WhatsApp Sender Bot Pro v2.0.0"
echo "Installation Script"
echo "=================================="
echo ""

# Check Python version
echo "Checking Python version..."
python3 --version || { echo "Error: Python 3 not found. Please install Python 3.8 or higher."; exit 1; }

# Create virtual environment
echo ""
echo "Creating virtual environment..."
python3 -m venv venv

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Upgrade pip
echo ""
echo "Upgrading pip..."
pip install --upgrade pip

# Install dependencies
echo ""
echo "Installing dependencies..."
pip install -r requirements.txt

# Copy environment file
echo ""
if [ ! -f .env ]; then
    echo "Creating .env file..."
    cp .env.example .env
    echo "✓ .env file created. Please edit it with your settings."
else
    echo "✓ .env file already exists."
fi

# Initialize database
echo ""
echo "Initializing database..."
python -c "from src.core.database import init_db; init_db()"

echo ""
echo "=================================="
echo "Installation Complete!"
echo "=================================="
echo ""
echo "Next steps:"
echo "1. Edit .env file with your settings"
echo "2. Run: source venv/bin/activate"
echo "3. Run: python main.py"
echo ""
echo "Default login credentials:"
echo "  Username: admin"
echo "  Password: admin123"
echo ""
echo "⚠️  IMPORTANT: Change the default password after first login!"
echo ""
