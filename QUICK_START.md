# ⚡ Quick Start - WhatsApp Sender Bot Pro

## 🎯 For People Who Don't Code

### Step 1️⃣: Install (Do Once)

**Windows:**
1. Double-click **`install.bat`**
2. Wait 2-3 minutes
3. Done! ✅

**Mac/Linux:**
1. Double-click **`install.sh`**
2. Wait 2-3 minutes
3. Done! ✅

---

### Step 2️⃣: Start the Bot (Every Time)

**Windows:**
1. Double-click **`start.bat`**

**Mac/Linux:**
1. Double-click **`start.sh`**

---

### Step 3️⃣: Login

When the window opens:
- **Username:** `admin`
- **Password:** `admin123`

⚠️ **IMPORTANT:** After login, go to Settings → Change Password!

---

### Step 4️⃣: Connect WhatsApp

First time only:
1. Click "Start Campaign" or "Send Test"
2. Chrome/Edge will open with WhatsApp Web
3. On your phone: Open WhatsApp → Settings → Linked Devices → Link a Device
4. Scan the QR code on computer

---

### Step 5️⃣: Send Messages

#### Import Contacts
1. Click **"Contacts"** tab
2. Click **"Import Contacts"**
3. Choose your Excel/CSV file
4. Click **"Process Contacts"**

#### Send Messages
1. Click **"Campaigns"** tab
2. Click **"New Campaign"**
3. Type your message
4. Click **"Start Campaign"**
5. Watch it go! 🚀

---

## 📋 Excel File Format

Your contact file should look like:

```
Phone Number    Name
+972501234567   John
0507654321      Jane
972509876543    Bob
```

**Supported formats:**
- `+972501234567` (with +)
- `0501234567` (Israeli local)
- `972501234567` (without +)

---

## ❓ Problems?

### "Can't find install.bat"
You're in the wrong folder. Open the **`Whatsapp-Sender-Bot`** folder.

### "Python not found"
Install Python: https://www.python.org/downloads/
**Check** "Add Python to PATH" when installing!

### "Login not working"
Username: `admin`
Password: `admin123`
(exactly as written)

### "WhatsApp QR code timeout"
1. Close browser
2. Restart the bot
3. Scan faster this time 😊

---

## 📂 Important Files

**Only care about these:**

| File | What is it? |
|------|-------------|
| `install.bat` / `install.sh` | Setup (run once) |
| `start.bat` / `start.sh` | Start bot (run every time) |
| `README.md` | Full instructions |
| `FILE_GUIDE.md` | What each file does |

**Don't touch these:**
- `src/` folder (code)
- `config/` folder (settings)
- `data/` folder (your database)
- `logs/` folder (activity logs)

---

## 🎬 Video Tutorial (Imagine)

1. **Install:** Double-click `install.bat` → Wait
2. **Start:** Double-click `start.bat` → Login screen appears
3. **Login:** Type `admin` / `admin123` → Click Login
4. **Import:** Click Contacts → Import → Choose Excel file
5. **Send:** Click Campaigns → New Campaign → Type message → Start

**Total time:** 5 minutes! ⏱️

---

## 💡 Pro Tips

1. **Always test first:** Use "Send Test" before sending to everyone
2. **Start small:** Try 5-10 contacts first
3. **Change password:** Settings → Change Password (use something secure)
4. **Save messages:** Use Templates to save messages you use often
5. **Check logs:** If something goes wrong, look in `logs/` folder

---

## 🆘 Still Stuck?

1. Read **`README.md`** (detailed guide)
2. Read **`USER_GUIDE.md`** (step-by-step tutorials)
3. Check **`logs/main.log`** (error messages)

---

**You got this! 🎉**
