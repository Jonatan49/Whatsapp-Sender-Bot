# WhatsApp Sender Bot Pro v2.0

## 📁 Project Files Explained

### **Files You Need (Don't Delete!)**

#### 🚀 **Start the Application**
- **`start.sh`** (Linux/Mac) - Double-click to start the bot
- **`start.bat`** (Windows) - Double-click to start the bot
- **`main.py`** - Main application file

#### 🔧 **Installation**
- **`install.sh`** (Linux/Mac) - Run once to install
- **`install.bat`** (Windows) - Run once to install
- **`requirements.txt`** - List of dependencies
- **`setup.py`** - Installation configuration

#### ⚙️ **Configuration**
- **`.env.example`** - Template for settings (copy to `.env`)
- **`config/config.yaml`** - Main configuration file

#### 📚 **Documentation**
- **`README.md`** - Complete user manual
- **`USER_GUIDE.md`** - Step-by-step tutorials
- **`UPGRADE_SUMMARY.md`** - What changed from v1.0
- **`THIS_FILE.md`** - This guide

#### 💻 **Source Code** (Don't modify unless you know what you're doing)
- **`src/`** - All the application code
  - `core/` - Main bot engine
  - `models/` - Database structure
  - `services/` - Business logic
  - `utils/` - Helper functions

---

### **Folders Created Automatically (Don't Delete!)**

- **`data/`** - Your database and saved data
- **`logs/`** - Application logs
- **`venv/`** - Python virtual environment (created after installation)

---

### **Hidden Files (Don't worry about these)**

- **`.git/`** - Version control (don't touch)
- **`.gitignore`** - Git configuration
- **`.env`** - Your personal settings (created from `.env.example`)

---

## 🎯 Simple Usage Guide

### **First Time Setup** (Do this once)

#### Windows:
1. Double-click **`install.bat`**
2. Wait for installation to complete
3. Copy `.env.example` to `.env` and edit if needed

#### Mac/Linux:
1. Double-click **`install.sh`** (or run in terminal: `./install.sh`)
2. Wait for installation to complete
3. Copy `.env.example` to `.env` and edit if needed

---

### **Running the Application** (Every time you want to use it)

#### Windows:
1. Double-click **`start.bat`**
2. Login with username: `admin`, password: `admin123`
3. **IMPORTANT:** Change password in Settings!

#### Mac/Linux:
1. Double-click **`start.sh`** (or run in terminal: `./start.sh`)
2. Login with username: `admin`, password: `admin123`
3. **IMPORTANT:** Change password in Settings!

---

## ❓ Common Questions

### "Which file do I click to start?"
- **Windows:** `start.bat`
- **Mac/Linux:** `start.sh`

### "I need to configure settings"
1. Copy `.env.example` to `.env`
2. Edit `.env` with a text editor
3. Or edit `config/config.yaml`

### "Where is my data saved?"
- Database: `data/whatsapp_bot.db`
- Logs: `logs/` folder

### "Can I delete old files?"
Yes! All old files have been removed. Only keep:
- Everything listed above under "Files You Need"
- The `src/` folder
- The `config/` folder

### "What if I break something?"
The code is in Git, so you can always restore:
```bash
git checkout .
```

---

## 📖 Need Help?

1. **Quick Start:** Read `README.md`
2. **Detailed Guide:** Read `USER_GUIDE.md`
3. **What Changed:** Read `UPGRADE_SUMMARY.md`

---

## 🎯 Quick Reference

| Task | Windows | Mac/Linux |
|------|---------|-----------|
| **Install** | Double-click `install.bat` | Double-click `install.sh` |
| **Start** | Double-click `start.bat` | Double-click `start.sh` |
| **Configure** | Edit `.env` or `config/config.yaml` | Edit `.env` or `config/config.yaml` |
| **Check Logs** | Open `logs/` folder | Open `logs/` folder |

---

**That's it! Keep it simple.** 🎉

If you see files not mentioned here, they were created by the system (like `__pycache__/`) and can be ignored.
