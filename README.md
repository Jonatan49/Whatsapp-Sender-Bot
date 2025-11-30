# WhatsApp Sender Bot Pro v2.0.0

![Version](https://img.shields.io/badge/version-2.0.0-blue)
![Python](https://img.shields.io/badge/python-3.8+-green)
![License](https://img.shields.io/badge/license-MIT-orange)

Professional WhatsApp bulk messaging bot with advanced features, built with Python, PyQt5, and Selenium.

## 🚀 What's New in v2.0

### ✨ Major Upgrades

- **Modular Architecture**: Complete rewrite with clean separation of concerns
- **Secure Authentication**: BCrypt password hashing, account lockout, session management
- **Database Layer**: SQLAlchemy ORM with SQLite for persistent storage
- **Advanced Logging**: Colored console output, file rotation, campaign-specific logs
- **Retry Logic**: Automatic retry with exponential backoff for failed operations
- **Anti-Detection**: Enhanced human-like behavior simulation
- **Message Templates**: Variable substitution and reusable templates
- **Campaign Management**: Track campaigns, statistics, and analytics
- **Configuration System**: YAML config files + environment variables
- **Professional Error Handling**: Comprehensive exception handling and recovery

### 🏗️ Architecture

```
whatsapp-sender-bot-pro/
├── src/
│   ├── core/              # Core components
│   │   ├── bot.py         # Enhanced WhatsApp bot with retry logic
│   │   ├── config.py      # Configuration management
│   │   └── database.py    # Database connection and session management
│   ├── models/            # Database models
│   │   ├── user.py        # User authentication
│   │   ├── contact.py     # Contact management
│   │   ├── campaign.py    # Campaign tracking
│   │   ├── message.py     # Message tracking
│   │   └── template.py    # Message templates
│   ├── services/          # Business logic
│   │   ├── auth_service.py      # Authentication service
│   │   ├── contact_service.py   # Contact management
│   │   ├── campaign_service.py  # Campaign management
│   │   └── message_sender.py    # Message sending logic
│   ├── ui/                # User interface (PyQt5)
│   └── utils/             # Utilities
│       ├── logger.py      # Advanced logging
│       └── translations.py # Multi-language support
├── config/
│   └── config.yaml        # Application configuration
├── data/                  # Database and data files
├── logs/                  # Log files
├── tests/                 # Unit and integration tests
├── .env                   # Environment variables (create from .env.example)
├── requirements.txt       # Python dependencies
├── setup.py              # Installation script
└── main.py               # Application entry point
```

## 📋 Features

### Core Features

- ✅ **Bulk Messaging**: Send messages to multiple contacts
- ✅ **Contact Management**: Import from Excel, CSV, JSON
- ✅ **Phone Validation**: Israeli phone number validation (+972)
- ✅ **Message Templates**: Reusable templates with variable substitution
- ✅ **Campaign Tracking**: Track sent, failed, pending messages
- ✅ **Real-time Logging**: Colored console + file logging with rotation
- ✅ **Multi-language**: English and Hebrew support

### Advanced Features

- ✅ **Retry Logic**: Automatic retry with exponential backoff
- ✅ **Anti-Detection**: Human-like typing, random delays, viewport randomization
- ✅ **Link Preview**: Optional link preview in messages
- ✅ **Pause/Resume**: Campaign flow control
- ✅ **Statistics**: Success rate, completion percentage, duration tracking
- ✅ **Backup System**: Automatic campaign backups
- ✅ **Secure Authentication**: BCrypt password hashing, account lockout
- ✅ **Database Persistence**: All data stored in SQLite database

### Security Features

- ✅ **Password Hashing**: BCrypt with configurable rounds
- ✅ **Account Lockout**: Automatic lockout after failed attempts
- ✅ **Session Management**: Secure session handling
- ✅ **Encrypted Logs**: Optional log encryption
- ✅ **No Hardcoded Credentials**: Credentials stored securely in database

## 🔧 Installation

### Prerequisites

- Python 3.8 or higher
- Google Chrome or Microsoft Edge browser
- pip (Python package manager)

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/whatsapp-sender-bot-pro.git
cd whatsapp-sender-bot-pro
```

### Step 2: Create Virtual Environment

```bash
python -m venv venv

# On Windows
venv\Scripts\activate

# On Linux/Mac
source venv/bin/activate
```

### Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 4: Configure Environment

```bash
# Copy example environment file
cp .env.example .env

# Edit .env with your settings
nano .env  # or use your preferred editor
```

### Step 5: Initialize Database

```bash
python -c "from src.core.database import init_db; init_db()"
```

### Step 6: Run Application

```bash
python main.py
```

## 🎯 Quick Start

### Default Login Credentials

**⚠️ IMPORTANT: Change these immediately after first login!**

- Username: `admin`
- Password: `admin123`

### First Time Setup

1. **Login**: Use default credentials
2. **Change Password**: Go to Settings → Change Password
3. **Configure Browser**: Select Chrome or Edge in Settings
4. **Import Contacts**: Import from Excel/CSV or add manually
5. **Create Campaign**: Create your first messaging campaign
6. **Start Sending**: Review and start your campaign

## 📚 Usage Guide

### Importing Contacts

#### From Excel/CSV

```python
from src.services.contact_service import ContactService

service = ContactService()
contacts, errors = service.import_contacts_from_file('contacts.xlsx')
print(f"Imported {len(contacts)} contacts")
```

**Excel/CSV Format:**
```
Column A: Phone Number
Column B: Name (optional)
```

#### From JSON

```json
[
  {"phone": "+972501234567", "name": "John Doe"},
  {"phone": "+972507654321", "name": "Jane Smith"}
]
```

### Creating Message Templates

```python
from src.models.template import MessageTemplate

template = MessageTemplate(
    name="Welcome Message",
    content="Hello {name}, welcome to {company}!",
    user_id=1
)
template.save_variables()  # Extracts variables: ['name', 'company']

# Render template
message = template.render({"name": "John", "company": "ACME Corp"})
# Result: "Hello John, welcome to ACME Corp!"
```

### Creating and Running a Campaign

```python
from src.services.campaign_service import CampaignService
from src.services.message_sender import MessageSender

# Create campaign
campaign_service = CampaignService()
campaign = campaign_service.create_campaign(
    name="Summer Sale 2024",
    message_content="Hi {name}, check out our summer sale!",
    user_id=1,
    contact_ids=[1, 2, 3, 4, 5],
    min_delay=5,
    max_delay=10,
    pause_after=50,
    pause_duration=120
)

# Start sending
sender = MessageSender()
thread = sender.start_campaign(campaign.id)
```

### Phone Number Validation

```python
from src.models.contact import Contact

# Normalize phone numbers
phone1 = Contact.normalize_phone("0501234567")  # Returns: +972501234567
phone2 = Contact.normalize_phone("+972501234567")  # Returns: +972501234567

# Validate Israeli numbers
is_valid = Contact.validate_israeli_phone("+972501234567")  # Returns: True
```

## ⚙️ Configuration

### Environment Variables (.env)

```bash
# Application
APP_NAME=WhatsApp Sender Bot Pro
APP_VERSION=2.0.0

# Database
DATABASE_URL=sqlite:///data/whatsapp_bot.db

# Security
SECRET_KEY=your-secret-key-here
SESSION_TIMEOUT=3600

# WhatsApp
WHATSAPP_DEFAULT_COUNTRY_CODE=+972
WHATSAPP_QR_TIMEOUT=120

# Browser
DEFAULT_BROWSER=chrome
BROWSER_HEADLESS=false

# Delays (seconds)
MIN_DELAY=5
MAX_DELAY=10
PAUSE_AFTER=100
PAUSE_DURATION=120

# Logging
LOG_LEVEL=INFO
LOG_MAX_BYTES=10485760
LOG_BACKUP_COUNT=5
```

### YAML Configuration (config/config.yaml)

See `config/config.yaml` for full configuration options including:
- Browser settings
- WhatsApp settings
- Delay configurations
- Anti-detection features
- Database settings
- Logging options
- Security settings
- UI preferences

## 🔒 Security Best Practices

1. **Change Default Password**: Immediately change the default admin password
2. **Secure .env File**: Never commit `.env` file to version control
3. **Use Strong Passwords**: Minimum 8 characters with mixed case, numbers, symbols
4. **Regular Backups**: Enable automatic database backups
5. **Encrypt Sensitive Data**: Enable log and database encryption
6. **Limit Login Attempts**: Configure max login attempts and lockout duration
7. **Review Logs**: Regularly review logs for suspicious activity

## 📊 Database Schema

### Users Table
- id, username, password_hash, email, full_name
- is_active, is_admin, last_login
- failed_login_attempts, locked_until
- language, theme, created_at, updated_at

### Contacts Table
- id, phone_number, name, country_code
- is_valid, is_blocked, notes, source
- messages_sent, messages_failed, last_contacted

### Campaigns Table
- id, name, description, status, user_id
- message_content, template_id
- min_delay, max_delay, pause_after, pause_duration
- total_contacts, messages_sent, messages_failed, messages_pending
- scheduled_at, started_at, completed_at, total_duration

### Messages Table
- id, campaign_id, contact_id, content, status
- phone_number, sent_at, delivery_time
- error_message, retry_count, link_preview

### Message Templates Table
- id, name, content, description, category
- user_id, variables (JSON), usage_count

## 🧪 Testing

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=src --cov-report=html

# Run specific test
pytest tests/test_contact_service.py
```

## 📝 Logging

### Log Levels

- **DEBUG**: Detailed debugging information
- **INFO**: General informational messages
- **WARNING**: Warning messages
- **ERROR**: Error messages
- **CRITICAL**: Critical error messages

### Log Locations

- **Console**: Colored output to terminal
- **Main Log**: `logs/main.log`
- **Campaign Logs**: `logs/campaigns/{campaign_id}_{timestamp}/`
- **Service Logs**: `logs/auth_service.log`, `logs/contact_service.log`, etc.

### Log Rotation

Logs automatically rotate when they reach 10MB (configurable). The last 5 log files are kept.

## 🚨 Important Legal Notice

**⚠️ WARNING: Use Responsibly**

This software is intended for legitimate business communication purposes only. Users are responsible for compliance with:

- **WhatsApp Terms of Service**: Bulk messaging may violate WhatsApp's ToS
- **Anti-Spam Laws**: CAN-SPAM Act, GDPR, TCPA, etc.
- **Local Regulations**: Check your local laws regarding automated messaging

**Possible Consequences:**
- WhatsApp account ban (temporary or permanent)
- Legal action for spam violations
- Fines and penalties

**Best Practices:**
- Only message people who have opted in
- Provide clear opt-out mechanism
- Respect rate limits
- Use for transactional messages, not marketing (unless explicitly allowed)
- Keep records of consent

## 🐛 Troubleshooting

### Browser Issues

**Problem**: Chrome/Edge driver not found

**Solution**:
```bash
pip install --upgrade webdriver-manager
```

**Problem**: Browser won't open

**Solution**: Check browser version matches driver version, or let webdriver-manager auto-download

### Database Issues

**Problem**: Database locked

**Solution**: Close all connections, restart application

**Problem**: Tables not created

**Solution**:
```bash
python -c "from src.core.database import init_db; db = init_db(); db.create_tables()"
```

### WhatsApp Issues

**Problem**: QR code timeout

**Solution**: Increase `WHATSAPP_QR_TIMEOUT` in .env

**Problem**: Message not sending

**Solution**: Check internet connection, WhatsApp Web status, phone number format

## 🤝 Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👏 Credits

- **Original Author**: Yonatan Cohen
- **Upgraded By**: AI Assistant
- **Libraries Used**: PyQt5, Selenium, SQLAlchemy, pandas, bcrypt

## 📞 Support

For issues, questions, or suggestions:

- **GitHub Issues**: https://github.com/yourusername/whatsapp-sender-bot-pro/issues
- **Email**: support@whatsappbot.local

## 🗺️ Roadmap

### Upcoming Features

- [ ] Web-based dashboard (FastAPI + React)
- [ ] API endpoints for external integration
- [ ] Scheduling campaigns for future dates
- [ ] Contact segmentation and tags
- [ ] Advanced analytics and reporting
- [ ] Multi-user support with roles
- [ ] Two-factor authentication
- [ ] WhatsApp Business API integration
- [ ] Media file support (images, videos, documents)
- [ ] Chatbot integration
- [ ] Webhook support
- [ ] Docker containerization
- [ ] Cloud deployment guides

## 📈 Changelog

### v2.0.0 (Current)

- ✅ Complete modular rewrite
- ✅ Secure authentication with BCrypt
- ✅ Database layer with SQLAlchemy
- ✅ Advanced logging system
- ✅ Retry logic and anti-detection
- ✅ Message templates
- ✅ Campaign management
- ✅ Configuration system

### v1.0.0 (Legacy)

- Basic bulk messaging
- Simple file import
- Hardcoded authentication
- Single-file architecture

---

**Made with ❤️ by Yonatan Cohen**
