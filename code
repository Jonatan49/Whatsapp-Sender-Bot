import sys, time, datetime, random, os, re, shutil, json
import pandas as pd

# נסה לייבא את pyperclip, ואם לא זמין – הודע למשתמש להתקין אותו
try:
    import pyperclip
except ImportError:
    pyperclip = None
    print("Error: pyperclip is not installed. Please install it using 'pip install pyperclip' and restart the application.")

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QStackedWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit, QMessageBox, QComboBox,
                             QRadioButton, QSpinBox, QGroupBox, QGridLayout, QDialog, QListWidget, QDialogButtonBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# -------------------------------------------------------------------
# הגדרות רב לשוניות (עברית)
# -------------------------------------------------------------------
LANG = "he"
translations = {
    "login_title": {"en": "Login", "he": "כניסה"},
    "username": {"en": "Username", "he": "שם משתמש"},
    "password": {"en": "Password", "he": "סיסמה"},
    "login_button": {"en": "Login", "he": "כניסה"},
    "language": {"en": "Language", "he": "שפה"},
    "main_title": {"en": "WhatsApp Sender Bot", "he": "בוט שליחת הודעות וואצאפ"},
    "upload_contacts": {"en": "Upload Contacts", "he": "העלה אנשי קשר"},
    "manual_numbers": {"en": "Paste numbers manually", "he": "הדבק מספרים ידנית"},
    "process_contacts": {"en": "Process Contacts", "he": "עיבוד אנשי קשר"},
    "test_number": {"en": "Test Number", "he": "מספר בדיקה"},
    "message": {"en": "Message", "he": "הודעה"},
    "send_test": {"en": "Send Test Message", "he": "שלח הודעת בדיקה"},
    "send_all": {"en": "Send Messages to All", "he": "שלח הודעות לכל"},
    "stop": {"en": "Stop", "he": "עצור"},
    "pause": {"en": "Pause", "he": "השהה"},
    "continue_": {"en": "Continue", "he": "המשך"},
    "min_delay": {"en": "Min Delay (sec)", "he": "מינימום השהיה (שניות)"},
    "max_delay": {"en": "Max Delay (sec)", "he": "מקסימום השהיה (שניות)"},
    "pause_after": {"en": "Pause After (# messages)", "he": "עצור אחרי (מספר הודעות)"},
    "pause_duration": {"en": "Pause Duration (sec)", "he": "משך הפסקה (שניות)"},
    "send_mode": {"en": "Send Mode", "he": "מצב שליחה"},
    "text_message": {"en": "Text Message", "he": "הודעה טקסטואלית"},
    "open_chat": {"en": "Open Chat Only", "he": "פתיחת צ'אט בלבד"},
    "timers_delays": {"en": "Timers and Delays", "he": "טיימרים והשהיות"},
    "bold": {"en": "Bold", "he": "הדגשה"},
    "italic": {"en": "Italic", "he": "נטוי"},
    "strike": {"en": "Strike", "he": "קו חוצה"},
    "mono": {"en": "Mono", "he": "מונוספייס"},
    "messages_sent": {"en": "Messages Sent:", "he": "הודעות שנשלחו:"},
    "current_data": {"en": "Current Data Folder:", "he": "נתיב נתונים נוכחי:"},
    "choose_data": {"en": "Choose Data Folder", "he": "בחר נתיב נתונים"},
    "browser": {"en": "Browser", "he": "דפדפן"},
    "chrome": {"en": "Chrome", "he": "כרום"},
    "edge": {"en": "Microsoft Edge", "he": "מיקרוסופט אדג'"},
    "browser_version": {"en": "Browser Version:", "he": "גרסת דפדפן:"},
    "load_title_numbers": {"en": "Load participants from title", "he": "טען מספרים מה-title"},
    "theme": {"en": "Theme", "he": "ערכת נושא"},
    "theme_light": {"en": "Light", "he": "בהיר"},
    "theme_dark": {"en": "Dark", "he": "כהה"},
    "theme_whatsapp": {"en": "WhatsApp", "he": "וואצאפ"},
    "message_library": {"en": "Message Library", "he": "ספריית הודעות"},
    "load_library": {"en": "Load Library", "he": "טען ספריה"},
    "save_library": {"en": "Save Current Message", "he": "שמור הודעה לספריה"},
    "summary_report": {"en": "Summary Report", "he": "דו\"ח סיכום"},
    "detect_link": {"en": "Detect Link", "he": "זיהוי קישור"},
}
def t(key):
    return translations.get(key, {}).get(LANG, key)

# -------------------------------------------------------------------
# ערכות נושא (Themes)
# -------------------------------------------------------------------
light_theme = """ 
QWidget { background-color: #f5f5f5; font-family: "Segoe UI"; color: #333; }
QGroupBox { border: 1px solid #ccc; border-radius: 5px; background-color: #fff; margin-top: 8px; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top center; padding: 2px 8px; color: #444; }
QPushButton { background-color: #E6F7FF; border: 1px solid #B3E5FC; border-radius: 4px; padding: 6px 12px; font-size: 13px; color: #333; }
QPushButton:hover { background-color: #CCEEFF; }
QLineEdit, QTextEdit { background-color: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 4px; font-size: 13px; color: #333; }
QLabel { font-size: 13px; color: #444; }
"""
dark_theme = """
QWidget { background-color: #1e1e1e; font-family: "Roboto"; color: #ffffff; }
QGroupBox { border: 1px solid #444; border-radius: 5px; background-color: #2b2b2b; margin-top: 8px; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top center; padding: 2px 8px; color: #00FF88; font-weight: bold; }
QPushButton { background-color: #2D3E50; border: none; border-radius: 4px; padding: 8px 16px; font-size: 13px; color: #fff; }
QPushButton:hover { background-color: #3A506B; }
QLineEdit, QTextEdit { background-color: #2b2b2b; border: 1px solid #444; border-radius: 4px; padding: 4px; font-size: 13px; color: #eeeeee; }
QLabel { font-size: 13px; color: #cccccc; }
"""
whatsapp_theme = """
QWidget { background-color: #e5ddd5; font-family: "Segoe UI", sans-serif; color: #202124; }
QGroupBox { border: 1px solid #ccc; border-radius: 5px; background-color: rgba(255, 255, 255, 0.85); margin-top: 10px; padding: 8px; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top center; padding: 4px 12px; background-color: #dcf8c6; color: #075e54; border-radius: 4px; font-weight: bold; }
QPushButton { background-color: #25d366; border: none; border-radius: 5px; padding: 8px 16px; font-size: 13px; color: #fff; }
QPushButton:hover { background-color: #20c158; }
QLineEdit, QTextEdit { background-color: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 6px; font-size: 13px; color: #202124; }
QLabel { font-size: 13px; color: #202124; }
"""

# -------------------------------------------------------------------
# מחלקת הבוט
# -------------------------------------------------------------------
class WhatsAppBot:
    def __init__(self, log_func, parent=None):
        self.log_func = log_func
        self.driver = None
        self.parent = parent

    def log(self, message):
        self.log_func(message)

    def init_driver(self):
        try:
            browser_choice = "Chrome"
            if self.parent and hasattr(self.parent, "browser_choice"):
                browser_choice = self.parent.browser_choice

            if browser_choice == "Chrome":
                from selenium.webdriver.chrome.service import Service as ChromeService
                from webdriver_manager.chrome import ChromeDriverManager
                options = webdriver.ChromeOptions()
                options.add_experimental_option("detach", True)
                if self.parent and hasattr(self.parent, "browser_version") and self.parent.browser_version:
                    version = self.parent.browser_version
                    self.log(f"Using specified Chrome version: {version}")
                    driver_path = ChromeDriverManager(version=version).install()
                else:
                    driver_path = ChromeDriverManager().install()
                service = ChromeService(driver_path)
                self.driver = webdriver.Chrome(service=service, options=options)
                self.log("הדפדפן הופעל עם כרום.")
            else:
                from selenium.webdriver.edge.service import Service as EdgeService
                from webdriver_manager.microsoft import EdgeChromiumDriverManager
                options = webdriver.EdgeOptions()
                options.add_experimental_option("detach", True)
                if self.parent and hasattr(self.parent, "browser_version") and self.parent.browser_version:
                    version = self.parent.browser_version
                    self.log(f"Using specified Edge version: {version}")
                    driver_path = EdgeChromiumDriverManager(version=version).install()
                else:
                    driver_path = EdgeChromiumDriverManager().install()
                service = EdgeService(executable_path=driver_path)
                self.driver = webdriver.Edge(service=service, options=options)
                self.log("הדפדפן הופעל עם מיקרוסופט אדג'.")
        except Exception as e:
            self.log(f"שגיאה באתחול הדפדפן ({browser_choice}): {e}")
            raise

    def open_whatsapp(self):
        if self.driver is None:
            self.init_driver()
        try:
            self.driver.get("https://web.whatsapp.com/")
            self.log("נפתח WhatsApp Web. סרוק את קוד ה-QR במידת הצורך.")
            if self.parent:
                QMessageBox.information(
                    self.parent,
                    "סרוק קוד QR",
                    "אנא סרוק את קוד ה-QR בחלון הדפדפן.\nלחץ OK כאשר סיימת."
                )
            self.log("WhatsApp Web מוכן.")
        except Exception as e:
            self.log(f"שגיאה בפתיחת WhatsApp Web: {e}")
            raise

    def send_message(self, number, message, mode="text"):
        try:
            # סינון תווים מחוץ ל-BMP
            message = ''.join(ch for ch in message if ord(ch) <= 0xFFFF)
            from selenium.webdriver.common.keys import Keys
            if self.driver is None:
                self.log("הדפדפן לא מאתחל. אתחל מחדש...")
                self.init_driver()
            try:
                if len(self.driver.window_handles) == 0:
                    self.log("חלון הדפדפן נסגר. אתחל מחדש את WhatsApp Web...")
                    self.open_whatsapp()
            except Exception as e:
                self.log(f"הדפדפן נסגר או אינו נגיש. אתחל מחדש. {e}")
                self.init_driver()
                self.open_whatsapp()
            self.driver.get(f"https://web.whatsapp.com/send?phone={number}")
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
            )
            time.sleep(2)
            text_box = self.driver.find_element(By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")
            
            if self.parent.link_preview_requested and pyperclip:
                # העתקת ההודעה ללוח והדבקה
                pyperclip.copy(message)
                text_box.send_keys(Keys.CONTROL, 'v')
                self.log("הודעה הודבקה מהלוח (כולל קישור) להפקת תצוגה מקדימה.")
                time.sleep(3)  # המתנה לטעינת ה-Preview
                text_box.send_keys(Keys.ENTER)
            else:
                # המרת \n ל-Shift+Enter כדי לשמור על הודעה רב־שורתית בהודעה אחת
                multiline_message = message.replace('\n', Keys.SHIFT + Keys.ENTER)
                text_box.send_keys(multiline_message)
                text_box.send_keys(Keys.ENTER)
            self.log(f"הודעה נשלחה ל-{number}.")
            return True
        except Exception as e:
            self.log(f"שגיאה בשליחת הודעה ל-{number}: {e}")
            return False

    def quit(self):
        if self.driver:
            self.driver.quit()
            self.log("הדפדפן נסגר.")

# -------------------------------------------------------------------
# Thread לשליחת הודעות
# -------------------------------------------------------------------
class MessageSenderThread(QThread):
    update_log = pyqtSignal(str)
    finished_sending = pyqtSignal()
    update_count = pyqtSignal(int)
    def __init__(self, contacts, message, mode, min_delay, max_delay, pause_after, pause_duration, bot: WhatsAppBot):
        super().__init__()
        self.contacts = contacts
        self.message = message
        self.mode = mode
        self.min_delay = min_delay
        self.max_delay = max_delay
        self.pause_after = pause_after
        self.pause_duration = pause_duration
        self.bot = bot
        self._stop_flag = False
        self._pause_flag = False
        self.sent_count = 0
        self.fail_count = 0
        self.start_time = time.time()
    def run(self):
        for i, number in enumerate(self.contacts, start=1):
            if self._stop_flag:
                self.update_log.emit("המשלוח נעצר על ידי המשתמש (עצור).")
                break
            while self._pause_flag and not self._stop_flag:
                time.sleep(0.5)
            if self._stop_flag:
                self.update_log.emit("המשלוח נעצר בזמן ההשהיה.")
                break
            timestamp = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            success = self.bot.send_message(number, self.message, self.mode)
            if not success:
                self.fail_count += 1
                self.update_log.emit(f"[{timestamp}] כשלון בשליחת הודעה ל-{number}.")
            else:
                self.update_log.emit(f"[{timestamp}] הודעה נשלחה ל-{number}.")
            self.sent_count += 1
            self.update_count.emit(self.sent_count)
            if self.pause_after > 0 and self.sent_count % self.pause_after == 0:
                self.update_log.emit(f"השהיה של {self.pause_duration} שניות אחרי {self.sent_count} הודעות.")
                time.sleep(self.pause_duration)
            delay = random.uniform(self.min_delay, self.max_delay)
            self.update_log.emit(f"המתנה של {round(delay, 2)} שניות לפני הודעה הבאה...")
            for _ in range(int(delay * 2)):
                if self._stop_flag:
                    break
                while self._pause_flag and not self._stop_flag:
                    time.sleep(0.5)
                if self._stop_flag:
                    break
                time.sleep(0.5)
            if self._stop_flag:
                self.update_log.emit("המשלוח נעצר בזמן ההמתנה.")
                break
        total_time = time.time() - self.start_time
        self.update_log.emit(f"סיום המעגל: נשלחו {self.sent_count} הודעות, {self.fail_count} כשלונות, זמן כולל: {int(total_time)} שניות.")
        self.finished_sending.emit()
    def stop(self):
        self._stop_flag = True
    def pause(self):
        self._pause_flag = True
    def resume(self):
        self._pause_flag = False

# -------------------------------------------------------------------
# חלון כניסה
# -------------------------------------------------------------------
class LoginWindow(QWidget):
    switch_window = pyqtSignal(str)
    def __init__(self):
        super().__init__()
        self.setWindowTitle(t("login_title"))
        self.setGeometry(100, 100, 300, 200)
        layout = QVBoxLayout()
        self.user_label = QLabel(t("username"))
        self.user_input = QLineEdit()
        self.pass_label = QLabel(t("password"))
        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.lang_label = QLabel(t("language"))
        self.lang_combo = QComboBox()
        self.lang_combo.addItem("עברית", "he")
        self.lang_combo.addItem("English", "en")
        self.login_button = QPushButton(t("login_button"))
        self.login_button.clicked.connect(self.handle_login)
        layout.addWidget(self.user_label)
        layout.addWidget(self.user_input)
        layout.addWidget(self.pass_label)
        layout.addWidget(self.pass_input)
        layout.addWidget(self.lang_label)
        layout.addWidget(self.lang_combo)
        layout.addWidget(self.login_button)
        self.setLayout(layout)
    def handle_login(self):
        username = self.user_input.text()
        password = self.pass_input.text()
        if username == "admin" and password == "bolbol":
            lang = self.lang_combo.currentData()
            self.switch_window.emit(lang)
        else:
            QMessageBox.critical(self, "Error", "שם משתמש או סיסמה לא נכונים.")

# -------------------------------------------------------------------
# חלון ראשי – פריסה אופקית עם 3 עמודות
# -------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self, language):
        super().__init__()
        global LANG
        LANG = language
        self.setWindowTitle(t("main_title"))
        self.setGeometry(100, 100, 1200, 700)
        self.contacts = []
        self.initial_contacts = []
        self.main_log_file = None
        self.job_log_file = None
        self.current_job_folder = None

        # אם לא הוזן נתיב – נתיב ברירת מחדל הוא הנתיב שבו נמצא הקובץ
        self.base_folder = os.path.dirname(os.path.abspath(__file__))
        self.browser_choice = "Chrome"
        self.browser_version = ""
        self.bot = WhatsAppBot(self.log_message, parent=self)
        self.message_thread = None
        self.paused = False
        # משתנה שיזהה אם המשתמש ביקש Preview לקישור
        self.link_preview_requested = False
        self.initUI()
        self.open_main_log()

    def open_main_log(self):
        daily_folder = os.path.join(self.base_folder, "בוט ווצאפ סופי", datetime.datetime.now().strftime("%d-%m-%Y"))
        if not os.path.exists(daily_folder):
            os.makedirs(daily_folder)
        log_filename = os.path.join(daily_folder, f"main_log_{datetime.datetime.now().strftime('%H-%M')}.txt")
        self.main_log_file = open(log_filename, "a", encoding="utf-8")
        self.log_message(f"נפתח קובץ לוג: {log_filename}")

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout()

        # עמודה שמאלית – הגדרות וספריית הודעות
        left_panel = QVBoxLayout()
        data_group = QGroupBox(t("current_data"))
        data_layout = QHBoxLayout()
        self.data_folder_lineedit = QLineEdit()
        self.data_folder_lineedit.setReadOnly(True)
        choose_data_button = QPushButton(t("choose_data"))
        choose_data_button.clicked.connect(self.choose_data_folder)
        data_layout.addWidget(choose_data_button)
        data_layout.addWidget(self.data_folder_lineedit)
        data_group.setLayout(data_layout)
        browser_group = QGroupBox(t("browser"))
        browser_layout = QVBoxLayout()
        top_browser_layout = QHBoxLayout()
        self.browser_combo = QComboBox()
        self.browser_combo.addItem(t("chrome"), "Chrome")
        self.browser_combo.addItem(t("edge"), "Edge")
        self.browser_combo.setCurrentIndex(0)
        self.browser_combo.currentIndexChanged.connect(self.browser_changed)
        top_browser_layout.addWidget(QLabel(t("browser") + ":"))
        top_browser_layout.addWidget(self.browser_combo)
        browser_layout.addLayout(top_browser_layout)
        version_layout = QHBoxLayout()
        self.browser_version_lineedit = QLineEdit()
        self.browser_version_lineedit.setPlaceholderText(t("browser_version"))
        version_layout.addWidget(QLabel(t("browser_version")))
        version_layout.addWidget(self.browser_version_lineedit)
        browser_layout.addLayout(version_layout)
        browser_group.setLayout(browser_layout)
        theme_group = QGroupBox(t("theme"))
        theme_layout = QHBoxLayout()
        self.theme_combo = QComboBox()
        self.theme_combo.addItem(t("theme_light"), "light")
        self.theme_combo.addItem(t("theme_dark"), "dark")
        self.theme_combo.addItem(t("theme_whatsapp"), "whatsapp")
        self.theme_combo.setCurrentIndex(2)
        self.apply_theme_button = QPushButton("החל עיצוב")
        self.apply_theme_button.clicked.connect(self.apply_theme)
        theme_layout.addWidget(QLabel(t("theme") + ":"))
        theme_layout.addWidget(self.theme_combo)
        theme_layout.addWidget(self.apply_theme_button)
        theme_group.setLayout(theme_layout)
        library_group = QGroupBox(t("message_library"))
        library_layout = QVBoxLayout()
        self.library_list = QListWidget()
        self.load_library_button = QPushButton(t("load_library"))
        self.load_library_button.clicked.connect(self.load_message_library)
        # לחיצה כפולה על פריט תדביק את ההודעה לשדה ההודעה
        self.library_list.itemDoubleClicked.connect(lambda item: self.message_input.setPlainText(item.text()))
        library_layout.addWidget(self.library_list)
        lib_buttons_layout = QHBoxLayout()
        lib_buttons_layout.addWidget(self.load_library_button)
        library_layout.addLayout(lib_buttons_layout)
        library_group.setLayout(library_layout)
        left_panel.addWidget(data_group)
        left_panel.addWidget(browser_group)
        left_panel.addWidget(theme_group)
        left_panel.addWidget(library_group)
        left_panel.addStretch(1)

        # עמודה מרכזית – ניהול אנשי קשר והודעות, כולל כפתור "זיהוי קישור"
        center_panel = QVBoxLayout()
        contacts_group = QGroupBox(t("process_contacts"))
        contacts_layout = QGridLayout()
        self.upload_button = QPushButton(t("upload_contacts"))
        self.upload_button.clicked.connect(self.upload_contacts)
        self.manual_text = QTextEdit()
        self.manual_text.setPlaceholderText(t("manual_numbers"))
        self.process_button = QPushButton(t("process_contacts"))
        self.process_button.clicked.connect(self.process_contacts)
        contacts_layout.addWidget(self.upload_button, 0, 0)
        contacts_layout.addWidget(self.manual_text, 1, 0)
        contacts_layout.addWidget(self.process_button, 2, 0)
        contacts_group.setLayout(contacts_layout)
        title_group = QGroupBox()
        title_layout = QHBoxLayout()
        self.load_title_button = QPushButton(t("load_title_numbers"))
        self.load_title_button.clicked.connect(self.load_group_participants_via_title)
        title_layout.addWidget(self.load_title_button)
        title_group.setLayout(title_layout)
        message_group = QGroupBox(t("message_sending"))
        message_layout = QGridLayout()
        self.test_number_input = QLineEdit()
        self.test_number_input.setPlaceholderText(t("test_number"))
        self.test_number_input.setText("+972547842680")
        message_box_layout = QVBoxLayout()
        self.message_input = QTextEdit()
        self.message_input.setPlaceholderText(t("message"))
        format_buttons_layout = QHBoxLayout()
        bold_button = QPushButton(t("bold"))
        italic_button = QPushButton(t("italic"))
        strike_button = QPushButton(t("strike"))
        mono_button = QPushButton(t("mono"))
        bold_button.clicked.connect(lambda: self.format_selection("*", "*"))
        italic_button.clicked.connect(lambda: self.format_selection("_", "_"))
        strike_button.clicked.connect(lambda: self.format_selection("~", "~"))
        mono_button.clicked.connect(lambda: self.format_selection("```", "```"))
        format_buttons_layout.addWidget(bold_button)
        format_buttons_layout.addWidget(italic_button)
        format_buttons_layout.addWidget(strike_button)
        format_buttons_layout.addWidget(mono_button)
        # כפתור לזיהוי קישורים
        self.detect_link_button = QPushButton(t("detect_link"))
        self.detect_link_button.clicked.connect(self.detect_links)
        format_buttons_layout.addWidget(self.detect_link_button)
        message_box_layout.addLayout(format_buttons_layout)
        message_box_layout.addWidget(self.message_input)
        self.send_test_button = QPushButton(t("send_test"))
        self.send_test_button.clicked.connect(self.send_test_message)
        self.send_all_button = QPushButton(t("send_all"))
        self.send_all_button.clicked.connect(self.send_all_messages)
        self.radio_text = QRadioButton(t("text_message"))
        self.radio_text.setChecked(True)
        self.radio_open = QRadioButton(t("open_chat"))
        message_layout.addWidget(QLabel(t("test_number")), 0, 0)
        message_layout.addWidget(self.test_number_input, 0, 1)
        message_layout.addWidget(QLabel(t("message")), 1, 0)
        message_layout.addLayout(message_box_layout, 1, 1, 1, 2)
        message_layout.addWidget(self.radio_text, 2, 0)
        message_layout.addWidget(self.radio_open, 2, 1)
        message_layout.addWidget(self.send_test_button, 3, 0)
        message_layout.addWidget(self.send_all_button, 3, 1)
        message_group.setLayout(message_layout)
        center_panel.addWidget(contacts_group)
        center_panel.addWidget(title_group)
        center_panel.addWidget(message_group)
        center_panel.addStretch(1)

        # עמודה ימנית – טיימרים, לוג, דו"ח סיכום, וגיבויים
        right_panel = QVBoxLayout()
        timer_group = QGroupBox(t("timers_delays"))
        timer_layout = QGridLayout()
        self.min_delay_spin = QSpinBox()
        self.min_delay_spin.setRange(1, 60)
        self.min_delay_spin.setValue(5)
        self.max_delay_spin = QSpinBox()
        self.max_delay_spin.setRange(1, 120)
        self.max_delay_spin.setValue(10)
        self.pause_after_spin = QSpinBox()
        self.pause_after_spin.setRange(0, 1000)
        self.pause_after_spin.setValue(100)
        self.pause_duration_spin = QSpinBox()
        self.pause_duration_spin.setRange(1, 600)
        self.pause_duration_spin.setValue(120)
        timer_layout.addWidget(QLabel(t("min_delay")), 0, 0)
        timer_layout.addWidget(self.min_delay_spin, 0, 1)
        timer_layout.addWidget(QLabel(t("max_delay")), 1, 0)
        timer_layout.addWidget(self.max_delay_spin, 1, 1)
        timer_layout.addWidget(QLabel(t("pause_after")), 2, 0)
        timer_layout.addWidget(self.pause_after_spin, 2, 1)
        timer_layout.addWidget(QLabel(t("pause_duration")), 3, 0)
        timer_layout.addWidget(self.pause_duration_spin, 3, 1)
        timer_group.setLayout(timer_layout)
        log_group = QGroupBox("לוג פעילות")
        log_layout = QVBoxLayout()
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        controls_layout = QHBoxLayout()
        self.pause_button = QPushButton(t("pause"))
        self.pause_button.clicked.connect(self.toggle_pause)
        self.stop_button = QPushButton(t("stop"))
        self.stop_button.clicked.connect(self.stop_sending)
        self.messages_sent_label = QLabel(f"{t('messages_sent')} 0")
        controls_layout.addWidget(self.pause_button)
        controls_layout.addWidget(self.stop_button)
        controls_layout.addWidget(self.messages_sent_label)
        log_layout.addWidget(self.log_area)
        log_layout.addLayout(controls_layout)
        log_group.setLayout(log_layout)
        summary_group = QGroupBox(t("summary_report"))
        summary_layout = QVBoxLayout()
        self.summary_text = QTextEdit()
        self.summary_text.setReadOnly(True)
        summary_layout.addWidget(self.summary_text)
        summary_group.setLayout(summary_layout)
        right_panel.addWidget(timer_group)
        right_panel.addWidget(log_group)
        right_panel.addWidget(summary_group)
        right_panel.addStretch(1)

        main_layout.addLayout(left_panel, 1)
        main_layout.addLayout(center_panel, 2)
        main_layout.addLayout(right_panel, 1)

        copyright_label = QLabel("© Yonatan Cohen")
        copyright_label.setAlignment(Qt.AlignCenter)
        copyright_label.setStyleSheet("color: #aaa; font-size: 10px;")
        overall_layout = QVBoxLayout()
        overall_layout.addLayout(main_layout)
        overall_layout.addWidget(copyright_label)
        central_widget.setLayout(overall_layout)

    def choose_data_folder(self):
        folder = QFileDialog.getExistingDirectory(self, t("choose_data"))
        if folder:
            self.base_folder = folder
            self.data_folder_lineedit.setText(folder)
            self.log_message(f"נתיב נתונים נבחר: {folder}")

    def browser_changed(self):
        self.browser_choice = self.browser_combo.currentData()
        self.log_message(f"דפדפן נבחר: {self.browser_choice}")
        self.browser_version = self.browser_version_lineedit.text().strip()

    def apply_theme(self):
        theme = self.theme_combo.currentData()
        if theme == "light":
            QApplication.instance().setStyleSheet(light_theme)
            self.log_message("ערכת נושא 'בהיר' הופעלה.")
        elif theme == "dark":
            QApplication.instance().setStyleSheet(dark_theme)
            self.log_message("ערכת נושא 'כהה' הופעלה.")
        elif theme == "whatsapp":
            QApplication.instance().setStyleSheet(whatsapp_theme)
            self.log_message("ערכת נושא 'וואצאפ' הופעלה.")

    def format_selection(self, prefix, suffix):
        cursor = self.message_input.textCursor()
        selected_text = cursor.selectedText()
        if selected_text:
            new_text = prefix + selected_text + suffix
            cursor.insertText(new_text)
        else:
            cursor.insertText(prefix + suffix)
        self.message_input.setFocus()

    def detect_links(self):
        text = self.message_input.toPlainText()
        pattern = r'(https?://[^\s]+)'
        links = re.findall(pattern, text)
        if links:
            reply = QMessageBox.question(
                self,
                "קישורים שנמצאו",
                f"נמצאו {len(links)} קישורים:\n" + "\n".join(links) + "\n\nהאם לכלול תצוגה מקדימה לקישור?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.link_preview_requested = True
                self.log_message("המשתמש בחר לכלול תצוגה מקדימה לקישור.", level="INFO")
            else:
                self.link_preview_requested = False
                self.log_message("המשתמש בחר שלא לכלול תצוגה מקדימה לקישור.", level="INFO")
        else:
            QMessageBox.information(self, "קישורים", "לא זוהו קישורים.")
            self.log_message("לא זוהו קישורים בהודעה.", level="INFO")
            self.link_preview_requested = False

    def load_group_participants_via_title(self):
        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
            span_element = WebDriverWait(self.bot.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'copyable-text') and @title]"))
            )
            title_content = span_element.get_attribute("title")
            self.log_message(f"תוכן title: {title_content}")
            pattern = r"\+[\d\s()-]+"
            found_numbers = re.findall(pattern, title_content)
            clean_numbers = []
            for fn in found_numbers:
                n = fn.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
                if n.startswith("+"):
                    clean_numbers.append(n)
            self.manual_text.clear()
            for cn in clean_numbers:
                self.manual_text.append(cn)
            self.log_message(f"טענו {len(clean_numbers)} מספרים מה-title.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"שגיאה בטעינת משתתפים מה-title: {e}")

    def upload_contacts(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Open Contacts File", "",
                                                  "Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json)",
                                                  options=options)
        if fileName:
            try:
                if fileName.endswith(".xlsx"):
                    df = pd.read_excel(fileName, engine="openpyxl")
                    new_contacts = df.iloc[:, 0].dropna().astype(str).tolist()
                elif fileName.endswith(".csv"):
                    df = pd.read_csv(fileName)
                    new_contacts = df.iloc[:, 0].dropna().astype(str).tolist()
                elif fileName.endswith(".json"):
                    with open(fileName, "r", encoding="utf-8") as f:
                        new_contacts = json.load(f)
                for num in new_contacts:
                    self.manual_text.append(num)
                self.log_message(f"טענו {len(new_contacts)} אנשי קשר מהקובץ.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"שגיאה בטעינת הקובץ: {e}")

    def process_contacts(self):
        raw_numbers = self.manual_text.toPlainText().splitlines()
        processed = []
        for num in raw_numbers:
            num = num.strip()
            if not num:
                continue
            if num.startswith("+972"):
                rest = num[4:]
                if not (rest.startswith("5") or rest.startswith("05")):
                    self.log_message(f"המספר {num} לא תקין (אינו מתחיל ב-05) – לא ייכלל.", level="WARNING")
                    continue
            elif num.startswith("972"):
                num = "+" + num
                rest = num[4:]
                if not (rest.startswith("5") or rest.startswith("05")):
                    self.log_message(f"המספר {num} לא תקין (אינו מתחיל ב-05) – לא ייכלל.", level="WARNING")
                    continue
            elif num.startswith("0"):
                num = "+972" + num[1:]
                rest = num[4:]
                if not (rest.startswith("5") or rest.startswith("05")):
                    self.log_message(f"המספר {num} לא תקין (אינו מתחיל ב-05) – לא ייכלל.", level="WARNING")
                    continue
            else:
                num = "+972" + num
            if len(num) not in [13, 14]:
                self.log_message(f"המספר {num} אינו באורך תקין – לא ייכלל.", level="WARNING")
                continue
            processed.append(num)
        unique_ordered = []
        for num in processed:
            if num not in unique_ordered:
                unique_ordered.append(num)
        self.contacts = unique_ordered
        self.initial_contacts = unique_ordered.copy()
        self.manual_text.clear()
        for num in self.contacts:
            self.manual_text.append(num)
        self.log_message(f"עיבוד אנשי קשר: {len(self.contacts)} מספרים תקינים.")

    def load_message_library(self):
        library_file = os.path.join(self.base_folder, "message_library.json")
        self.library_list.clear()
        if os.path.exists(library_file):
            with open(library_file, "r", encoding="utf-8") as f:
                try:
                    messages = json.load(f)
                    for msg in messages:
                        self.library_list.addItem(msg)
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"שגיאה בטעינת ספריית ההודעות: {e}")
        else:
            self.library_list.addItem("אין הודעות שמורות.")

    def save_current_message_to_library(self):
        library_file = os.path.join(self.base_folder, "message_library.json")
        current_msg = self.message_input.toPlainText().strip()
        if not current_msg:
            QMessageBox.warning(self, "Warning", "תיבת ההודעה ריקה.")
            return
        if os.path.exists(library_file):
            with open(library_file, "r", encoding="utf-8") as f:
                try:
                    messages = json.load(f)
                except Exception:
                    messages = []
        else:
            messages = []
        messages.append(current_msg)
        with open(library_file, "w", encoding="utf-8") as f:
            json.dump(messages, f, ensure_ascii=False, indent=2)
        self.log_message("הודעה נשמרה לספריית ההודעות.")
        self.load_message_library()

    def send_test_message(self):
        test_number = self.test_number_input.text().strip()
        mode = "text" if self.radio_text.isChecked() else "open"
        if mode == "text" and not self.message_input.toPlainText().strip():
            QMessageBox.warning(self, "Warning", "תוכן ההודעה ריק.")
            return
        if not test_number:
            QMessageBox.warning(self, "Warning", "אנא הזן מספר בדיקה.")
            return
        if not test_number.startswith("+"):
            if test_number.startswith("0"):
                test_number = "+972" + test_number[1:]
            else:
                test_number = "+972" + test_number
        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"שגיאה בפתיחת WhatsApp Web: {e}")
            return
        message = ''.join(ch for ch in self.message_input.toPlainText().strip() if ord(ch) <= 0xFFFF)
        success = self.bot.send_message(test_number, message, mode)
        if success:
            QMessageBox.information(self, "Success", f"הודעת בדיקה נשלחה ל-{test_number}.")
        else:
            QMessageBox.critical(self, "Error", f"כשלון בשליחת הודעת בדיקה ל-{test_number} (עיין בלוג).")

    def send_all_messages(self):
        if not self.contacts:
            QMessageBox.warning(self, "Warning", "אין אנשי קשר לשליחה.")
            return
        mode = "text" if self.radio_text.isChecked() else "open"
        if mode == "text" and not self.message_input.toPlainText().strip():
            QMessageBox.warning(self, "Warning", "תוכן ההודעה ריק.")
            return
        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"שגיאה בפתיחת WhatsApp Web: {e}")
            return
        # שמירת ההודעה לספריית ההודעות באופן אוטומטי
        current_msg = self.message_input.toPlainText().strip()
        if current_msg:
            self.save_current_message_to_library()
        daily_folder = os.path.join(self.base_folder, "בוט ווצאפ סופי", datetime.datetime.now().strftime("%d-%m-%Y"))
        if not os.path.exists(daily_folder):
            os.makedirs(daily_folder)
        cycle_folder = os.path.join(daily_folder, datetime.datetime.now().strftime("%H-%M"))
        os.makedirs(cycle_folder, exist_ok=True)
        self.current_job_folder = cycle_folder
        job_log_file_path = os.path.join(cycle_folder, "activity_log.txt")
        self.job_log_file = open(job_log_file_path, "a", encoding="utf-8")
        self.log_message(f"נוצר תיקיית משימה: {cycle_folder}")
        full_message = self.message_input.toPlainText().strip()
        if self.link_preview_requested and pyperclip:
            # במצב Preview, נעתיק את ההודעה ללוח ונדביק אותה
            pyperclip.copy(full_message)
            prepared_message = full_message  # נשמר למעקב
        else:
            from selenium.webdriver.common.keys import Keys
            prepared_message = full_message.replace('\n', Keys.SHIFT + Keys.ENTER)
        min_delay = self.min_delay_spin.value()
        max_delay = self.max_delay_spin.value()
        pause_after = self.pause_after_spin.value()
        pause_duration = self.pause_duration_spin.value()
        self.message_thread = MessageSenderThread(
            self.contacts, prepared_message, mode,
            min_delay, max_delay,
            pause_after, pause_duration,
            self.bot
        )
        self.message_thread.update_log.connect(self.log_message)
        self.message_thread.finished_sending.connect(self.handle_cycle_finished)
        self.message_thread.update_count.connect(self.update_sent_count)
        self.message_thread.start()
        self.manual_text.clear()
        self.contacts = []
        self.initial_contacts = []
        self.log_message("איפוס אנשי הקשר, מוכן לטעינה חדשה.")

    def handle_cycle_finished(self):
        self.sending_finished(silent=False)
        total_time = time.time() - self.message_thread.start_time if self.message_thread else 0
        summary = (f"מספר הודעות שנשלחו: {self.message_thread.sent_count}\n"
                   f"מספר כשלונות: {self.message_thread.fail_count}\n"
                   f"זמן כולל: {int(total_time)} שניות")
        summary_file = os.path.join(self.current_job_folder, "summary_report.txt")
        with open(summary_file, "w", encoding="utf-8") as f:
            f.write(summary)
        self.summary_text.setPlainText(summary)
        self.log_message("דו\"ח סיכום נוצר.")
        backup_folder = os.path.join(self.base_folder, "בוט ווצאפ סופי", "Backups", datetime.datetime.now().strftime("%d-%m-%Y"))
        if not os.path.exists(backup_folder):
            os.makedirs(backup_folder)
        backup_dest = os.path.join(backup_folder, os.path.basename(self.current_job_folder))
        try:
            shutil.copytree(self.current_job_folder, backup_dest)
            self.log_message(f"גיבוי אוטומטי נוצר: {backup_dest}")
        except Exception as e:
            self.log_message(f"שגיאה ביצירת גיבוי: {e}")

    def stop_sending(self):
        if self.message_thread:
            self.message_thread.stop()
            self.log_message("משתמש ביקש עצירה מלאה.")
            self.message_thread = None
            self.update_sent_count(0)
            self.manual_text.clear()
            self.contacts = []
            self.initial_contacts = []
            self.log_message("איפוס אנשי הקשר, מוכן לטעינה חדשה.")

    def toggle_pause(self):
        if not self.message_thread:
            return
        if not self.paused:
            self.message_thread.pause()
            self.pause_button.setText(t("continue_"))
            self.paused = True
            self.log_message("משתמש ביקש השהיה.")
        else:
            self.message_thread.resume()
            self.pause_button.setText(t("pause"))
            self.paused = False
            self.log_message("משתמש ביקש המשך.")
            
    def update_sent_count(self, count):
        self.messages_sent_label.setText(f"{t('messages_sent')} {count}")

    def sending_finished(self, silent=False):
        self.log_message("תהליך השליחה הסתיים.")
        pending = self.initial_contacts[self.message_thread.sent_count:] if self.message_thread else self.initial_contacts
        pending_file = os.path.join(self.current_job_folder, "pending_contacts.txt")
        with open(pending_file, "w", encoding="utf-8") as f:
            for num in pending:
                f.write(num + "\n")
        initial_file = os.path.join(self.current_job_folder, "initial_contacts.txt")
        with open(initial_file, "w", encoding="utf-8") as f:
            for num in self.initial_contacts:
                f.write(num + "\n")
        if self.job_log_file and not self.job_log_file.closed:
            self.job_log_file.close()
            self.job_log_file = None
        self.paused = False
        self.pause_button.setText(t("pause"))
        self.update_sent_count(0)
        self.message_thread = None
        if not silent:
            QMessageBox.information(self, "Job Finished", f"תהליך השליחה הסתיים.\nתיקיית המשימה:\n{self.current_job_folder}")

    def log_message(self, msg, level="INFO"):
        timestamp = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        full_msg = f"[{timestamp}] [{level}] {msg}"
        self.log_area.append(full_msg)
        if self.main_log_file and not self.main_log_file.closed:
            self.main_log_file.write(full_msg + "\n")
            self.main_log_file.flush()
        if self.job_log_file and not self.job_log_file.closed:
            self.job_log_file.write(full_msg + "\n")
            self.job_log_file.flush()

    def closeEvent(self, event):
        if self.message_thread:
            self.message_thread.stop()
            self.sending_finished(silent=True)
        if self.bot:
            self.bot.quit()
        if self.main_log_file and not self.main_log_file.closed:
            self.main_log_file.write("[PROGRAM EXIT] סגירת קובץ הלוג.\n")
            self.main_log_file.close()
        super().closeEvent(event)

class App(QStackedWidget):
    def __init__(self):
        super().__init__()
        self.login_window = LoginWindow()
        self.addWidget(self.login_window)
        self.login_window.switch_window.connect(self.show_main)
    def show_main(self, language):
        self.main_window = MainWindow(language)
        self.addWidget(self.main_window)
        self.setCurrentWidget(self.main_window)

def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(whatsapp_theme)
    window = App()
    window.setWindowIcon(QIcon("C:\\Users\\hamuz\\Desktop\\בוט לווצאפ סופי\\Future_logo.ico"))
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
