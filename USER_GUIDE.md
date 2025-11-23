# WhatsApp Sender Bot Pro - User Guide

## Table of Contents

1. [Getting Started](#getting-started)
2. [Authentication](#authentication)
3. [Managing Contacts](#managing-contacts)
4. [Creating Campaigns](#creating-campaigns)
5. [Message Templates](#message-templates)
6. [Settings](#settings)
7. [Analytics](#analytics)
8. [Troubleshooting](#troubleshooting)

## Getting Started

### First Launch

1. **Run the application**:
   ```bash
   python main.py
   ```

2. **Login with default credentials**:
   - Username: `admin`
   - Password: `admin123`

3. **⚠️ Change your password immediately**:
   - Go to Settings → Change Password
   - Use a strong password (min 8 characters)

### Understanding the Interface

The application is divided into several sections:

- **Dashboard**: Overview of campaigns and statistics
- **Contacts**: Manage phone numbers and contact lists
- **Campaigns**: Create and manage messaging campaigns
- **Templates**: Save and reuse message templates
- **Settings**: Configure application settings
- **Analytics**: View campaign statistics and reports

## Authentication

### Changing Your Password

1. Navigate to **Settings** → **Change Password**
2. Enter your current password
3. Enter your new password (min 8 characters)
4. Confirm your new password
5. Click **Change Password**

### Account Security

- **Failed Login Attempts**: After 5 failed attempts, your account will be locked for 5 minutes
- **Session Timeout**: Sessions expire after 1 hour of inactivity (configurable)
- **Password Requirements**: Minimum 8 characters, mixed case recommended

## Managing Contacts

### Adding Contacts Manually

1. Go to **Contacts** tab
2. Click **Add Contact**
3. Enter phone number in international format (e.g., +972501234567)
4. Optionally add contact name
5. Click **Save**

### Importing Contacts from File

#### Excel/CSV Format

Create a file with the following format:

| Phone Number    | Name (Optional) |
|-----------------|-----------------|
| +972501234567   | John Doe        |
| 0507654321      | Jane Smith      |
| 972509876543    | Bob Johnson     |

**Supported formats:**
- +972501234567 (International)
- 0501234567 (Local Israeli)
- 972501234567 (Without +)

#### Import Steps

1. Go to **Contacts** tab
2. Click **Import Contacts**
3. Select your file (Excel .xlsx, CSV .csv, or JSON .json)
4. Click **Open**
5. Review imported contacts
6. Click **Process Contacts** to validate and save

### Contact Validation

The system automatically validates Israeli phone numbers:

✅ **Valid Formats**:
- Must start with +972
- Must be followed by 5 (mobile numbers)
- Total length: 13-14 characters

❌ **Invalid Numbers**:
- Landline numbers (not starting with 5)
- Incorrect length
- Missing country code

### Managing Contact Groups

1. Go to **Contacts** → **Groups**
2. Click **Create Group**
3. Enter group name and description
4. Select contacts to add
5. Click **Save**

## Creating Campaigns

### Step-by-Step Campaign Creation

#### 1. Prepare Your Message

1. Go to **Campaigns** tab
2. Click **New Campaign**
3. Enter campaign name (e.g., "Summer Sale 2024")
4. Enter description (optional)

#### 2. Compose Your Message

**Plain Text**:
```
Hello! Check out our summer sale.
Up to 50% off on all items!
Visit: www.example.com
```

**With Formatting**:
- **Bold**: `*text*`
- _Italic_: `_text_`
- ~Strikethrough~: `~text~`
- `Monospace`: ` ```text``` `

#### 3. Using Variables (with Templates)

```
Hello {name},

Your appointment is scheduled for {date} at {time}.

Thanks,
{company}
```

Variables will be replaced with actual values from your contact list.

#### 4. Select Recipients

- **All Contacts**: Send to everyone in your contact list
- **Specific Group**: Send to a contact group
- **Custom Selection**: Choose individual contacts

#### 5. Configure Timing

**Delay Settings**:
- **Min Delay**: Minimum seconds between messages (default: 5)
- **Max Delay**: Maximum seconds between messages (default: 10)
- **Random Delay**: System picks random value between min and max

**Pause Settings**:
- **Pause After**: Automatically pause after N messages (default: 100)
- **Pause Duration**: How long to pause in seconds (default: 120)

#### 6. Start Campaign

1. Review all settings
2. Click **Send Test Message** to verify
3. Click **Start Campaign**
4. Scan QR code if prompted (first time only)
5. Monitor progress in real-time

### Campaign Controls

While a campaign is running:

- **Pause**: Temporarily pause sending (click Resume to continue)
- **Stop**: Completely stop the campaign (cannot be resumed)
- **View Log**: See real-time message status

### Understanding Campaign Status

- **Draft**: Campaign created but not started
- **Running**: Currently sending messages
- **Paused**: Temporarily paused by user
- **Completed**: All messages sent successfully
- **Failed**: Critical error occurred
- **Cancelled**: Stopped by user

## Message Templates

### Creating a Template

1. Go to **Templates** tab
2. Click **New Template**
3. Enter template name (e.g., "Welcome Message")
4. Enter message content with variables:
   ```
   Hello {name},

   Welcome to {company}!

   Your account ID is: {account_id}
   ```
5. Select category (optional): General, Marketing, Reminder, etc.
6. Click **Save Template**

### Using Variables

Variables are placeholders in the format `{variable_name}`:

**Supported Variable Names**:
- Letters, numbers, underscore
- Must start with letter or underscore
- Case-sensitive

**Examples**:
- `{name}` - Contact name
- `{date}` - Appointment date
- `{amount}` - Invoice amount
- `{order_id}` - Order number

### Using a Template in Campaign

1. Create new campaign
2. Click **Load Template**
3. Select your template
4. Provide values for variables
5. Preview the message
6. Start campaign

## Settings

### Browser Settings

**Browser Selection**:
- **Chrome**: Google Chrome (recommended)
- **Edge**: Microsoft Edge

**Browser Version**:
- Leave empty for auto-detection
- Enter specific version if needed (e.g., "120.0.6099.109")

**Options**:
- **Headless Mode**: Run browser in background (no visible window)
- **Detach Mode**: Keep browser open after script ends

### Delay & Timing Settings

**Message Delays**:
- **Min Delay**: 5-10 seconds recommended
- **Max Delay**: 10-15 seconds recommended
- **Why?**: Appears more human-like, avoids WhatsApp rate limits

**Batch Pauses**:
- **Pause After**: 50-100 messages recommended
- **Pause Duration**: 60-120 seconds recommended
- **Why?**: Prevents account restrictions

### Anti-Detection Settings

**Human Behavior Simulation**:
- ✅ Random typing speed
- ✅ Random delays
- ✅ Random scrolling
- ✅ Viewport randomization

**Safety Recommendations**:
- Don't send more than 500 messages per day
- Use realistic delays
- Don't run campaigns 24/7
- Take breaks between campaigns

### Theme Settings

Choose from three themes:
- **Light**: Clean, bright interface
- **Dark**: Easy on the eyes, modern look
- **WhatsApp**: WhatsApp-inspired green theme

### Language Settings

- **English**: Full English interface
- **Hebrew**: עברית מלאה

## Analytics

### Campaign Statistics

View detailed statistics for each campaign:

- **Total Contacts**: Number of recipients
- **Messages Sent**: Successfully delivered
- **Messages Failed**: Failed deliveries
- **Messages Pending**: Not yet sent
- **Success Rate**: Percentage of successful deliveries
- **Completion**: Overall progress percentage
- **Duration**: Total time taken
- **Average Delay**: Average time between messages

### Exporting Reports

1. Go to **Analytics** tab
2. Select campaign
3. Click **Export Report**
4. Choose format:
   - PDF: Formatted report
   - Excel: Data for analysis
   - CSV: Simple data export
   - JSON: Machine-readable format

### Understanding Reports

**Success Rate**:
- 95-100%: Excellent
- 90-95%: Good
- 85-90%: Acceptable
- Below 85%: Investigate issues

**Common Failure Reasons**:
- Invalid phone number
- Contact blocked you
- Network issues
- WhatsApp account restricted

## Troubleshooting

### Common Issues

#### Browser Won't Open

**Symptoms**: Error message about Chrome/Edge driver

**Solution**:
```bash
pip install --upgrade webdriver-manager
```

#### QR Code Timeout

**Symptoms**: QR code expires before you can scan it

**Solution**:
1. Increase timeout in `.env`:
   ```
   WHATSAPP_QR_TIMEOUT=180
   ```
2. Restart application

#### Messages Not Sending

**Possible Causes**:
1. **Invalid phone number**: Check format (+972501234567)
2. **WhatsApp not connected**: Rescan QR code
3. **Account restricted**: WhatsApp blocked your number
4. **Internet issues**: Check connection

**Debugging Steps**:
1. Check logs in `logs/` folder
2. Try test message first
3. Verify phone number format
4. Check WhatsApp Web in browser manually

#### Database Locked Error

**Symptoms**: "Database is locked" error

**Solution**:
1. Close all application instances
2. Restart application
3. If persists, delete `data/whatsapp_bot.db-journal`

#### Account Locked

**Symptoms**: "Account is locked" message

**Solution**:
1. Wait 5 minutes (default lockout duration)
2. Or reset in database:
   ```bash
   python -c "from src.services.auth_service import AuthService; from src.core.database import get_db; db = get_db(); from src.models.user import User; from datetime import datetime; user = db.get_session().query(User).filter_by(username='admin').first(); user.locked_until = None; user.failed_login_attempts = 0; db.get_session().commit()"
   ```

### Getting Help

1. **Check Logs**: Review `logs/main.log` for errors
2. **Read FAQ**: Check README.md FAQ section
3. **GitHub Issues**: Report bugs at GitHub repository
4. **Email Support**: support@whatsappbot.local

### Best Practices

1. **Always Test First**: Use "Send Test" before starting campaign
2. **Start Small**: Test with 5-10 contacts first
3. **Use Realistic Delays**: Don't set delays too low
4. **Monitor Campaigns**: Watch logs for issues
5. **Backup Database**: Regular backups of `data/whatsapp_bot.db`
6. **Keep Software Updated**: Check for updates regularly
7. **Respect Privacy**: Only message people who opted in
8. **Follow Laws**: Comply with anti-spam regulations

### Performance Tips

1. **Database Optimization**: Vacuum database monthly:
   ```bash
   sqlite3 data/whatsapp_bot.db "VACUUM;"
   ```

2. **Log Cleanup**: Remove old logs:
   ```bash
   find logs/ -name "*.log" -mtime +30 -delete
   ```

3. **Contact List Management**: Remove inactive contacts regularly

4. **Campaign Archival**: Archive completed campaigns to separate database

---

**Need more help?** Contact support or check the GitHub repository for updates and community discussions.
