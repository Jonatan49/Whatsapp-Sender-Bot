#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
בוט שליחת הודעות וואצאפ - גרסה מתקדמת
© Yonatan Cohen
"""

import sys
import time
import datetime
import random
import os
import re
import json
import csv
# winsound is imported conditionally when needed (Windows only)

import pandas as pd

try:
    import pyperclip
except ImportError:
    pyperclip = None

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit, QMessageBox,
    QComboBox, QSpinBox, QGroupBox, QGridLayout, QListWidget, QListWidgetItem,
    QProgressBar, QTabWidget, QCheckBox, QTimeEdit, QDateTimeEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QSplitter, QFrame,
    QDialog, QDialogButtonBox, QInputDialog, QSystemTrayIcon, QMenu, QAction
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer, QTime, QDateTime, QSize
from PyQt5.QtGui import QFont, QColor, QIcon, QPixmap

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# -------------------------------------------------------------------
# עיצוב מודרני בסגנון וואצאפ
# -------------------------------------------------------------------
MAIN_STYLE = """
QMainWindow, QDialog {
    background-color: #0b141a;
}
QWidget {
    background-color: #0b141a;
    font-family: "Segoe UI", "Arial", sans-serif;
    color: #e9edef;
    font-size: 13px;
}
QTabWidget::pane {
    border: 2px solid #00a884;
    border-radius: 10px;
    background-color: #111b21;
}
QTabBar::tab {
    background-color: #1f2c33;
    color: #8696a0;
    padding: 10px 20px;
    margin-right: 2px;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}
QTabBar::tab:selected {
    background-color: #00a884;
    color: white;
    font-weight: bold;
}
QTabBar::tab:hover:!selected {
    background-color: #2a3942;
}
QGroupBox {
    border: 2px solid #00a884;
    border-radius: 10px;
    background-color: #111b21;
    margin-top: 15px;
    padding: 15px;
    padding-top: 25px;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 5px 20px;
    background-color: #00a884;
    color: white;
    border-radius: 5px;
    font-weight: bold;
    font-size: 14px;
}
QPushButton {
    background-color: #00a884;
    border: none;
    border-radius: 8px;
    padding: 10px 20px;
    font-size: 13px;
    font-weight: bold;
    color: white;
    min-height: 18px;
}
QPushButton:hover {
    background-color: #02c091;
}
QPushButton:pressed {
    background-color: #008c6f;
}
QPushButton:disabled {
    background-color: #3a4750;
    color: #8696a0;
}
QPushButton#stopButton, QPushButton#deleteBtn, QPushButton#clearBtn {
    background-color: #ea4335;
}
QPushButton#stopButton:hover, QPushButton#deleteBtn:hover, QPushButton#clearBtn:hover {
    background-color: #f55246;
}
QPushButton#pauseButton, QPushButton#warningBtn {
    background-color: #fbbc04;
    color: #1a1a1a;
}
QPushButton#pauseButton:hover, QPushButton#warningBtn:hover {
    background-color: #ffc929;
}
QPushButton#secondaryBtn {
    background-color: #2d3e50;
}
QPushButton#secondaryBtn:hover {
    background-color: #3a506b;
}
QLineEdit, QTextEdit {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 8px;
    font-size: 13px;
    color: #e9edef;
    selection-background-color: #00a884;
}
QLineEdit:focus, QTextEdit:focus {
    border: 2px solid #00a884;
}
QLabel {
    font-size: 13px;
    color: #8696a0;
    background-color: transparent;
}
QLabel#titleLabel {
    font-size: 28px;
    font-weight: bold;
    color: #00a884;
}
QLabel#subtitleLabel {
    font-size: 14px;
    color: #8696a0;
}
QLabel#statusLabel {
    font-size: 16px;
    font-weight: bold;
    color: #00a884;
    padding: 10px;
    background-color: #1f2c33;
    border-radius: 8px;
}
QLabel#statNumber {
    font-size: 32px;
    font-weight: bold;
    color: #00a884;
}
QLabel#statLabel {
    font-size: 12px;
    color: #8696a0;
}
QComboBox {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 8px;
    color: #e9edef;
    min-width: 100px;
}
QComboBox:hover {
    border: 2px solid #00a884;
}
QComboBox::drop-down {
    border: none;
    padding-right: 10px;
}
QComboBox QAbstractItemView {
    background-color: #1f2c33;
    color: #e9edef;
    selection-background-color: #00a884;
}
QSpinBox, QTimeEdit, QDateTimeEdit {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 8px;
    color: #e9edef;
    min-width: 80px;
}
QSpinBox:focus, QTimeEdit:focus, QDateTimeEdit:focus {
    border: 2px solid #00a884;
}
QListWidget {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 5px;
    color: #e9edef;
}
QListWidget::item {
    padding: 8px;
    border-radius: 5px;
    margin: 2px;
}
QListWidget::item:selected {
    background-color: #00a884;
    color: white;
}
QListWidget::item:hover:!selected {
    background-color: #2a3942;
}
QTableWidget {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    gridline-color: #2a3942;
    color: #e9edef;
}
QTableWidget::item {
    padding: 5px;
}
QTableWidget::item:selected {
    background-color: #00a884;
    color: white;
}
QHeaderView::section {
    background-color: #111b21;
    color: #00a884;
    padding: 8px;
    border: none;
    font-weight: bold;
}
QProgressBar {
    border: none;
    border-radius: 8px;
    background-color: #1f2c33;
    height: 25px;
    text-align: center;
    font-weight: bold;
}
QProgressBar::chunk {
    background-color: #00a884;
    border-radius: 8px;
}
QCheckBox {
    color: #e9edef;
    spacing: 8px;
}
QCheckBox::indicator {
    width: 20px;
    height: 20px;
    border-radius: 4px;
    border: 2px solid #2a3942;
    background-color: #1f2c33;
}
QCheckBox::indicator:checked {
    background-color: #00a884;
    border: 2px solid #00a884;
}
QScrollBar:vertical {
    background-color: #111b21;
    width: 10px;
    border-radius: 5px;
}
QScrollBar::handle:vertical {
    background-color: #2a3942;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::handle:vertical:hover {
    background-color: #00a884;
}
QScrollBar:horizontal {
    background-color: #111b21;
    height: 10px;
    border-radius: 5px;
}
QScrollBar::handle:horizontal {
    background-color: #2a3942;
    border-radius: 5px;
    min-width: 30px;
}
QSplitter::handle {
    background-color: #2a3942;
}
QMenu {
    background-color: #1f2c33;
    border: 1px solid #2a3942;
    border-radius: 8px;
    padding: 5px;
}
QMenu::item {
    padding: 8px 20px;
    border-radius: 4px;
}
QMenu::item:selected {
    background-color: #00a884;
}
"""


# -------------------------------------------------------------------
# נתיבי קבצים
# -------------------------------------------------------------------
def get_data_path():
    """מחזיר את נתיב תיקיית הנתונים"""
    base = os.path.dirname(os.path.abspath(__file__))
    data_path = os.path.join(base, "data")
    os.makedirs(data_path, exist_ok=True)
    return data_path


def get_file_path(filename):
    """מחזיר נתיב מלא לקובץ בתיקיית הנתונים"""
    return os.path.join(get_data_path(), filename)


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
                driver_path = ChromeDriverManager().install()
                service = ChromeService(driver_path)
                self.driver = webdriver.Chrome(service=service, options=options)
                self.log("דפדפן Chrome הופעל בהצלחה")
            else:
                from selenium.webdriver.edge.service import Service as EdgeService
                from webdriver_manager.microsoft import EdgeChromiumDriverManager
                options = webdriver.EdgeOptions()
                options.add_experimental_option("detach", True)
                driver_path = EdgeChromiumDriverManager().install()
                service = EdgeService(executable_path=driver_path)
                self.driver = webdriver.Edge(service=service, options=options)
                self.log("דפדפן Edge הופעל בהצלחה")
        except Exception as e:
            self.log(f"שגיאה באתחול הדפדפן: {e}")
            raise

    def open_whatsapp(self):
        if self.driver is None:
            self.init_driver()
        try:
            self.driver.get("https://web.whatsapp.com/")
            self.log("WhatsApp Web נפתח - סרוק את קוד ה-QR")
            if self.parent:
                QMessageBox.information(
                    self.parent,
                    "סריקת QR",
                    "סרוק את קוד ה-QR בחלון הדפדפן.\n\nלחץ OK כשסיימת."
                )
            self.log("WhatsApp Web מוכן לשימוש")
        except Exception as e:
            self.log(f"שגיאה בפתיחת WhatsApp Web: {e}")
            raise

    def send_message(self, number, message, mode="text", simulate=False):
        """שולח הודעה למספר. אם simulate=True, רק מדמה בלי לשלוח באמת"""
        try:
            message = ''.join(ch for ch in message if ord(ch) <= 0xFFFF)

            if simulate:
                self.log(f"[סימולציה] הודעה ל-{number}")
                time.sleep(0.5)
                return True

            if self.driver is None:
                self.log("מאתחל דפדפן מחדש...")
                self.init_driver()

            try:
                if len(self.driver.window_handles) == 0:
                    self.log("חלון הדפדפן נסגר - פותח מחדש...")
                    self.open_whatsapp()
            except Exception:
                self.init_driver()
                self.open_whatsapp()

            self.driver.get(f"https://web.whatsapp.com/send?phone={number}")
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
            )
            time.sleep(2)

            text_box = self.driver.find_element(By.XPATH, "//div[@contenteditable='true' and @data-tab='10']")

            if self.parent and self.parent.link_preview_requested and pyperclip:
                pyperclip.copy(message)
                text_box.send_keys(Keys.CONTROL, 'v')
                time.sleep(3)
                text_box.send_keys(Keys.ENTER)
            else:
                multiline_message = message.replace('\n', Keys.SHIFT + Keys.ENTER)
                text_box.send_keys(multiline_message)
                text_box.send_keys(Keys.ENTER)

            self.log(f"הודעה נשלחה ל-{number}")
            return True
        except Exception as e:
            self.log(f"שגיאה בשליחה ל-{number}: {e}")
            return False

    def quit(self):
        if self.driver:
            self.driver.quit()
            self.log("הדפדפן נסגר")


# -------------------------------------------------------------------
# Thread לשליחת הודעות
# -------------------------------------------------------------------
class MessageSenderThread(QThread):
    update_log = pyqtSignal(str)
    finished_sending = pyqtSignal()
    update_count = pyqtSignal(int, int)
    message_sent = pyqtSignal(str, bool)  # number, success

    def __init__(self, contacts, message, mode, min_delay, max_delay,
                 pause_after, pause_duration, bot, blacklist=None,
                 simulate=False, personalize_data=None):
        super().__init__()
        self.contacts = contacts
        self.message = message
        self.mode = mode
        self.min_delay = min_delay
        self.max_delay = max_delay
        self.pause_after = pause_after
        self.pause_duration = pause_duration
        self.bot = bot
        self.blacklist = blacklist or []
        self.simulate = simulate
        self.personalize_data = personalize_data or {}
        self._stop_flag = False
        self._pause_flag = False
        self.sent_count = 0
        self.fail_count = 0
        self.start_time = time.time()
        self.results = []

    def personalize_message(self, message, number):
        """מחליף משתנים בהודעה"""
        result = message

        # משתנים מובנים
        result = result.replace("{תאריך}", datetime.datetime.now().strftime("%d/%m/%Y"))
        result = result.replace("{שעה}", datetime.datetime.now().strftime("%H:%M"))
        result = result.replace("{יום}", self.get_hebrew_day())
        result = result.replace("{מספר}", number)

        # משתנים מותאמים אישית מהנתונים
        if number in self.personalize_data:
            data = self.personalize_data[number]
            for key, value in data.items():
                result = result.replace("{" + key + "}", str(value))

        return result

    def get_hebrew_day(self):
        days = ["שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת", "ראשון"]
        return days[datetime.datetime.now().weekday()]

    def run(self):
        total = len(self.contacts)
        for i, number in enumerate(self.contacts, start=1):
            if self._stop_flag:
                self.update_log.emit("השליחה נעצרה על ידי המשתמש")
                break

            # בדיקת רשימה שחורה
            if number in self.blacklist:
                self.update_log.emit(f"דילוג על {number} (ברשימה שחורה)")
                continue

            while self._pause_flag and not self._stop_flag:
                time.sleep(0.5)

            if self._stop_flag:
                break

            # התאמה אישית של ההודעה
            personalized_msg = self.personalize_message(self.message, number)

            timestamp = datetime.datetime.now().strftime('%H:%M:%S')
            prefix = "[סימולציה] " if self.simulate else ""

            success = self.bot.send_message(number, personalized_msg, self.mode, self.simulate)

            if not success:
                self.fail_count += 1
                self.update_log.emit(f"{prefix}[{timestamp}] נכשל: {number}")
                self.results.append({"number": number, "status": "failed", "time": timestamp})
            else:
                self.update_log.emit(f"{prefix}[{timestamp}] נשלח: {number}")
                self.results.append({"number": number, "status": "success", "time": timestamp})

            self.message_sent.emit(number, success)
            self.sent_count += 1
            self.update_count.emit(self.sent_count, total)

            if self.pause_after > 0 and self.sent_count % self.pause_after == 0:
                self.update_log.emit(f"הפסקה של {self.pause_duration} שניות...")
                time.sleep(self.pause_duration)

            if i < total and not self._stop_flag:
                delay = random.uniform(self.min_delay, self.max_delay)
                self.update_log.emit(f"ממתין {round(delay, 1)} שניות...")

                for _ in range(int(delay * 2)):
                    if self._stop_flag:
                        break
                    while self._pause_flag and not self._stop_flag:
                        time.sleep(0.5)
                    if self._stop_flag:
                        break
                    time.sleep(0.5)

        total_time = time.time() - self.start_time
        success_count = self.sent_count - self.fail_count
        self.update_log.emit(f"סיום! הצלחות: {success_count}, כשלונות: {self.fail_count} ({int(total_time)} שניות)")
        self.finished_sending.emit()

    def stop(self):
        self._stop_flag = True

    def pause(self):
        self._pause_flag = True

    def resume(self):
        self._pause_flag = False


# -------------------------------------------------------------------
# Thread לתזמון שליחה
# -------------------------------------------------------------------
class SchedulerThread(QThread):
    ready_to_send = pyqtSignal()
    time_remaining = pyqtSignal(str)

    def __init__(self, target_datetime):
        super().__init__()
        self.target_datetime = target_datetime
        self._stop_flag = False

    def run(self):
        while not self._stop_flag:
            now = QDateTime.currentDateTime()
            if now >= self.target_datetime:
                self.ready_to_send.emit()
                break

            secs = now.secsTo(self.target_datetime)
            hours = secs // 3600
            mins = (secs % 3600) // 60
            secs = secs % 60
            self.time_remaining.emit(f"{hours:02d}:{mins:02d}:{secs:02d}")
            time.sleep(1)

    def stop(self):
        self._stop_flag = True


# -------------------------------------------------------------------
# דיאלוג לעריכת תבנית
# -------------------------------------------------------------------
class TemplateDialog(QDialog):
    def __init__(self, parent=None, template_name="", template_text=""):
        super().__init__(parent)
        self.setWindowTitle("תבנית הודעה")
        self.setMinimumSize(500, 400)
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout()

        # שם התבנית
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("שם התבנית:"))
        self.name_input = QLineEdit(template_name)
        name_layout.addWidget(self.name_input)
        layout.addLayout(name_layout)

        # תוכן ההודעה
        layout.addWidget(QLabel("תוכן ההודעה:"))
        self.text_input = QTextEdit()
        self.text_input.setPlainText(template_text)
        layout.addWidget(self.text_input)

        # משתנים זמינים
        vars_group = QGroupBox("משתנים זמינים (לחץ להוספה)")
        vars_layout = QHBoxLayout()
        for var in ["{שם}", "{תאריך}", "{שעה}", "{יום}", "{מספר}"]:
            btn = QPushButton(var)
            btn.setObjectName("secondaryBtn")
            btn.clicked.connect(lambda checked, v=var: self.insert_variable(v))
            vars_layout.addWidget(btn)
        vars_group.setLayout(vars_layout)
        layout.addWidget(vars_group)

        # כפתורים
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def insert_variable(self, var):
        self.text_input.insertPlainText(var)

    def get_data(self):
        return self.name_input.text(), self.text_input.toPlainText()


# -------------------------------------------------------------------
# חלון ראשי
# -------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("בוט שליחת הודעות וואצאפ - גרסה מתקדמת")
        self.setGeometry(50, 50, 1300, 850)
        self.setLayoutDirection(Qt.RightToLeft)

        # משתנים
        self.contacts = []
        self.contacts_data = {}  # נתונים נוספים לכל איש קשר
        self.blacklist = []
        self.templates = {}
        self.history = []
        self.base_folder = os.path.dirname(os.path.abspath(__file__))
        self.browser_choice = "Chrome"
        self.bot = WhatsAppBot(self.log_message, parent=self)
        self.message_thread = None
        self.scheduler_thread = None
        self.paused = False
        self.link_preview_requested = False
        self.current_job_folder = None

        # סטטיסטיקות
        self.total_sent = 0
        self.total_failed = 0
        self.session_sent = 0
        self.session_failed = 0

        # טעינת נתונים שמורים
        self.load_saved_data()

        self.initUI()

    def load_saved_data(self):
        """טוען נתונים שמורים מקבצים"""
        # טעינת תבניות
        templates_file = get_file_path("templates.json")
        if os.path.exists(templates_file):
            with open(templates_file, "r", encoding="utf-8") as f:
                self.templates = json.load(f)

        # טעינת רשימה שחורה
        blacklist_file = get_file_path("blacklist.json")
        if os.path.exists(blacklist_file):
            with open(blacklist_file, "r", encoding="utf-8") as f:
                self.blacklist = json.load(f)

        # טעינת היסטוריה
        history_file = get_file_path("history.json")
        if os.path.exists(history_file):
            with open(history_file, "r", encoding="utf-8") as f:
                self.history = json.load(f)

        # טעינת סטטיסטיקות
        stats_file = get_file_path("stats.json")
        if os.path.exists(stats_file):
            with open(stats_file, "r", encoding="utf-8") as f:
                stats = json.load(f)
                self.total_sent = stats.get("total_sent", 0)
                self.total_failed = stats.get("total_failed", 0)

    def save_data(self):
        """שומר נתונים לקבצים"""
        # שמירת תבניות
        with open(get_file_path("templates.json"), "w", encoding="utf-8") as f:
            json.dump(self.templates, f, ensure_ascii=False, indent=2)

        # שמירת רשימה שחורה
        with open(get_file_path("blacklist.json"), "w", encoding="utf-8") as f:
            json.dump(self.blacklist, f, ensure_ascii=False, indent=2)

        # שמירת היסטוריה (שומר רק 1000 אחרונים)
        with open(get_file_path("history.json"), "w", encoding="utf-8") as f:
            json.dump(self.history[-1000:], f, ensure_ascii=False, indent=2)

        # שמירת סטטיסטיקות
        with open(get_file_path("stats.json"), "w", encoding="utf-8") as f:
            json.dump({
                "total_sent": self.total_sent,
                "total_failed": self.total_failed
            }, f)

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # כותרת
        header = QHBoxLayout()
        title_label = QLabel("בוט שליחת הודעות וואצאפ")
        title_label.setObjectName("titleLabel")
        header.addWidget(title_label)
        header.addStretch()

        # סטטיסטיקות מהירות
        self.quick_stats = QLabel(f"נשלחו היום: {self.session_sent} | סה\"כ: {self.total_sent}")
        self.quick_stats.setStyleSheet("color: #00a884; font-size: 14px;")
        header.addWidget(self.quick_stats)

        main_layout.addLayout(header)

        # טאבים
        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_main_tab(), "שליחה")
        self.tabs.addTab(self.create_contacts_tab(), "אנשי קשר")
        self.tabs.addTab(self.create_templates_tab(), "תבניות")
        self.tabs.addTab(self.create_blacklist_tab(), "רשימה שחורה")
        self.tabs.addTab(self.create_schedule_tab(), "תזמון")
        self.tabs.addTab(self.create_history_tab(), "היסטוריה")
        self.tabs.addTab(self.create_settings_tab(), "הגדרות")

        main_layout.addWidget(self.tabs)

        # פס תחתון
        footer = QHBoxLayout()
        footer_label = QLabel("© Yonatan Cohen | גרסה 2.0")
        footer_label.setStyleSheet("color: #4a5568; font-size: 11px;")
        footer.addWidget(footer_label)
        footer.addStretch()
        main_layout.addLayout(footer)

        central_widget.setLayout(main_layout)

    def create_main_tab(self):
        """טאב שליחה ראשי"""
        widget = QWidget()
        layout = QHBoxLayout()
        layout.setSpacing(15)

        # צד ימין - הודעה ושליחה
        right_side = QVBoxLayout()

        # קבוצת הודעה
        msg_group = QGroupBox("תוכן ההודעה")
        msg_layout = QVBoxLayout()

        # כפתורי עיצוב
        format_layout = QHBoxLayout()
        for text, prefix, suffix in [("מודגש", "*", "*"), ("נטוי", "_", "_"), ("קו חוצה", "~", "~")]:
            btn = QPushButton(text)
            btn.setMaximumWidth(80)
            btn.setObjectName("secondaryBtn")
            btn.clicked.connect(lambda c, p=prefix, s=suffix: self.format_selection(p, s))
            format_layout.addWidget(btn)

        self.link_btn = QPushButton("זיהוי קישור")
        self.link_btn.setMaximumWidth(100)
        self.link_btn.setObjectName("secondaryBtn")
        self.link_btn.clicked.connect(self.detect_links)
        format_layout.addWidget(self.link_btn)

        # בחירת תבנית
        self.template_combo = QComboBox()
        self.template_combo.addItem("-- בחר תבנית --")
        self.template_combo.currentTextChanged.connect(self.load_selected_template)
        self.update_template_combo()
        format_layout.addWidget(self.template_combo)

        format_layout.addStretch()
        msg_layout.addLayout(format_layout)

        self.message_input = QTextEdit()
        self.message_input.setPlaceholderText("כתוב את ההודעה כאן...\n\nמשתנים זמינים: {שם}, {תאריך}, {שעה}, {יום}")
        self.message_input.setMinimumHeight(180)
        msg_layout.addWidget(self.message_input)

        msg_group.setLayout(msg_layout)
        right_side.addWidget(msg_group)

        # כפתורי פעולה
        action_group = QGroupBox("פעולות")
        action_layout = QVBoxLayout()

        # שורה עליונה - בדיקה
        test_layout = QHBoxLayout()
        self.test_input = QLineEdit()
        self.test_input.setPlaceholderText("מספר לבדיקה")
        test_layout.addWidget(self.test_input)

        self.test_btn = QPushButton("שלח בדיקה")
        self.test_btn.clicked.connect(self.send_test_message)
        test_layout.addWidget(self.test_btn)

        self.simulate_check = QCheckBox("מצב סימולציה")
        self.simulate_check.setToolTip("בדיקה ללא שליחה אמיתית")
        test_layout.addWidget(self.simulate_check)

        action_layout.addLayout(test_layout)

        # שורה תחתונה - שליחה
        send_layout = QHBoxLayout()

        self.pause_btn = QPushButton("השהה")
        self.pause_btn.setObjectName("pauseButton")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.pause_btn.setEnabled(False)
        send_layout.addWidget(self.pause_btn)

        self.stop_btn = QPushButton("עצור")
        self.stop_btn.setObjectName("stopButton")
        self.stop_btn.clicked.connect(self.stop_sending)
        self.stop_btn.setEnabled(False)
        send_layout.addWidget(self.stop_btn)

        send_layout.addStretch()

        self.send_all_btn = QPushButton("שלח לכל אנשי הקשר")
        self.send_all_btn.setMinimumWidth(200)
        self.send_all_btn.clicked.connect(self.send_all_messages)
        send_layout.addWidget(self.send_all_btn)

        action_layout.addLayout(send_layout)
        action_group.setLayout(action_layout)
        right_side.addWidget(action_group)

        layout.addLayout(right_side, 2)

        # צד שמאל - לוג וסטטוס
        left_side = QVBoxLayout()

        # סטטיסטיקות סשן
        stats_group = QGroupBox("סטטיסטיקות")
        stats_layout = QHBoxLayout()

        for label_text, attr_name in [("אנשי קשר", "contacts_stat"), ("נשלחו", "sent_stat"), ("נכשלו", "failed_stat")]:
            stat_widget = QVBoxLayout()
            num_label = QLabel("0")
            num_label.setObjectName("statNumber")
            num_label.setAlignment(Qt.AlignCenter)
            setattr(self, attr_name, num_label)
            stat_widget.addWidget(num_label)
            text_label = QLabel(label_text)
            text_label.setObjectName("statLabel")
            text_label.setAlignment(Qt.AlignCenter)
            stat_widget.addWidget(text_label)
            stats_layout.addLayout(stat_widget)

        stats_group.setLayout(stats_layout)
        left_side.addWidget(stats_group)

        # לוג
        log_group = QGroupBox("יומן פעילות")
        log_layout = QVBoxLayout()

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        log_layout.addWidget(self.log_area)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%v / %m")
        log_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("מוכן לשליחה")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignCenter)
        log_layout.addWidget(self.status_label)

        log_group.setLayout(log_layout)
        left_side.addWidget(log_group)

        layout.addLayout(left_side, 1)

        widget.setLayout(layout)
        return widget

    def create_contacts_tab(self):
        """טאב ניהול אנשי קשר"""
        widget = QWidget()
        layout = QVBoxLayout()

        # כפתורי פעולה
        buttons_layout = QHBoxLayout()

        self.upload_btn = QPushButton("העלה קובץ")
        self.upload_btn.clicked.connect(self.upload_contacts)
        buttons_layout.addWidget(self.upload_btn)

        self.load_group_btn = QPushButton("טען מקבוצת וואצאפ")
        self.load_group_btn.clicked.connect(self.load_group_participants)
        buttons_layout.addWidget(self.load_group_btn)

        self.export_btn = QPushButton("ייצא לאקסל")
        self.export_btn.setObjectName("secondaryBtn")
        self.export_btn.clicked.connect(self.export_contacts)
        buttons_layout.addWidget(self.export_btn)

        self.clear_contacts_btn = QPushButton("נקה הכל")
        self.clear_contacts_btn.setObjectName("clearBtn")
        self.clear_contacts_btn.clicked.connect(self.clear_contacts)
        buttons_layout.addWidget(self.clear_contacts_btn)

        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

        # תיבת הזנה ידנית
        manual_group = QGroupBox("הזנה ידנית")
        manual_layout = QVBoxLayout()

        self.manual_text = QTextEdit()
        self.manual_text.setPlaceholderText("הדבק מספרים כאן (כל מספר בשורה נפרדת)\n\nניתן גם להדביק עם שמות בפורמט:\nמספר,שם\n0541234567,יוסי")
        self.manual_text.setMaximumHeight(120)
        manual_layout.addWidget(self.manual_text)

        process_layout = QHBoxLayout()
        self.process_btn = QPushButton("עבד והוסף")
        self.process_btn.clicked.connect(self.process_contacts)
        process_layout.addWidget(self.process_btn)
        process_layout.addStretch()
        manual_layout.addLayout(process_layout)

        manual_group.setLayout(manual_layout)
        layout.addWidget(manual_group)

        # טבלת אנשי קשר
        table_group = QGroupBox("רשימת אנשי קשר")
        table_layout = QVBoxLayout()

        self.contacts_table = QTableWidget()
        self.contacts_table.setColumnCount(3)
        self.contacts_table.setHorizontalHeaderLabels(["מספר", "שם", "פעולות"])
        self.contacts_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.contacts_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.contacts_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.contacts_table.setColumnWidth(2, 150)
        table_layout.addWidget(self.contacts_table)

        self.contacts_count_label = QLabel("0 אנשי קשר")
        self.contacts_count_label.setAlignment(Qt.AlignCenter)
        table_layout.addWidget(self.contacts_count_label)

        table_group.setLayout(table_layout)
        layout.addWidget(table_group)

        widget.setLayout(layout)
        return widget

    def create_templates_tab(self):
        """טאב תבניות הודעות"""
        widget = QWidget()
        layout = QHBoxLayout()

        # רשימת תבניות
        list_layout = QVBoxLayout()

        list_layout.addWidget(QLabel("תבניות שמורות:"))
        self.templates_list = QListWidget()
        self.templates_list.itemDoubleClicked.connect(self.edit_template)
        self.update_templates_list()
        list_layout.addWidget(self.templates_list)

        # כפתורים
        btn_layout = QHBoxLayout()
        add_btn = QPushButton("הוסף תבנית")
        add_btn.clicked.connect(self.add_template)
        btn_layout.addWidget(add_btn)

        edit_btn = QPushButton("ערוך")
        edit_btn.setObjectName("secondaryBtn")
        edit_btn.clicked.connect(self.edit_template)
        btn_layout.addWidget(edit_btn)

        del_btn = QPushButton("מחק")
        del_btn.setObjectName("deleteBtn")
        del_btn.clicked.connect(self.delete_template)
        btn_layout.addWidget(del_btn)

        list_layout.addLayout(btn_layout)
        layout.addLayout(list_layout)

        # תצוגה מקדימה
        preview_layout = QVBoxLayout()
        preview_layout.addWidget(QLabel("תצוגה מקדימה:"))

        self.template_preview = QTextEdit()
        self.template_preview.setReadOnly(True)
        self.templates_list.currentItemChanged.connect(self.show_template_preview)
        preview_layout.addWidget(self.template_preview)

        use_btn = QPushButton("השתמש בתבנית")
        use_btn.clicked.connect(self.use_selected_template)
        preview_layout.addWidget(use_btn)

        layout.addLayout(preview_layout)

        widget.setLayout(layout)
        return widget

    def create_blacklist_tab(self):
        """טאב רשימה שחורה"""
        widget = QWidget()
        layout = QVBoxLayout()

        layout.addWidget(QLabel("מספרים ברשימה השחורה לא יקבלו הודעות:"))

        # הוספה ידנית
        add_layout = QHBoxLayout()
        self.blacklist_input = QLineEdit()
        self.blacklist_input.setPlaceholderText("הזן מספר להוספה")
        add_layout.addWidget(self.blacklist_input)

        add_btn = QPushButton("הוסף")
        add_btn.clicked.connect(self.add_to_blacklist)
        add_layout.addWidget(add_btn)

        layout.addLayout(add_layout)

        # רשימה
        self.blacklist_widget = QListWidget()
        self.update_blacklist_display()
        layout.addWidget(self.blacklist_widget)

        # כפתורים
        btn_layout = QHBoxLayout()
        remove_btn = QPushButton("הסר נבחר")
        remove_btn.setObjectName("deleteBtn")
        remove_btn.clicked.connect(self.remove_from_blacklist)
        btn_layout.addWidget(remove_btn)

        clear_btn = QPushButton("נקה הכל")
        clear_btn.setObjectName("clearBtn")
        clear_btn.clicked.connect(self.clear_blacklist)
        btn_layout.addWidget(clear_btn)

        btn_layout.addStretch()

        import_btn = QPushButton("ייבא מקובץ")
        import_btn.setObjectName("secondaryBtn")
        import_btn.clicked.connect(self.import_blacklist)
        btn_layout.addWidget(import_btn)

        layout.addLayout(btn_layout)

        widget.setLayout(layout)
        return widget

    def create_schedule_tab(self):
        """טאב תזמון"""
        widget = QWidget()
        layout = QVBoxLayout()

        schedule_group = QGroupBox("תזמון שליחה")
        schedule_layout = QVBoxLayout()

        schedule_layout.addWidget(QLabel("בחר תאריך ושעה לשליחה:"))

        datetime_layout = QHBoxLayout()
        self.schedule_datetime = QDateTimeEdit()
        self.schedule_datetime.setDateTime(QDateTime.currentDateTime().addSecs(3600))
        self.schedule_datetime.setCalendarPopup(True)
        datetime_layout.addWidget(self.schedule_datetime)
        datetime_layout.addStretch()
        schedule_layout.addLayout(datetime_layout)

        self.schedule_btn = QPushButton("הפעל תזמון")
        self.schedule_btn.clicked.connect(self.toggle_schedule)
        schedule_layout.addWidget(self.schedule_btn)

        self.schedule_status = QLabel("לא פעיל")
        self.schedule_status.setObjectName("statusLabel")
        self.schedule_status.setAlignment(Qt.AlignCenter)
        schedule_layout.addWidget(self.schedule_status)

        self.countdown_label = QLabel("")
        self.countdown_label.setAlignment(Qt.AlignCenter)
        self.countdown_label.setStyleSheet("font-size: 32px; color: #00a884;")
        schedule_layout.addWidget(self.countdown_label)

        schedule_group.setLayout(schedule_layout)
        layout.addWidget(schedule_group)

        layout.addStretch()

        widget.setLayout(layout)
        return widget

    def create_history_tab(self):
        """טאב היסטוריה"""
        widget = QWidget()
        layout = QVBoxLayout()

        # טבלת היסטוריה
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(4)
        self.history_table.setHorizontalHeaderLabels(["תאריך", "מספר", "סטטוס", "הודעה"])
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.history_table)

        # כפתורים
        btn_layout = QHBoxLayout()

        export_btn = QPushButton("ייצא לאקסל")
        export_btn.clicked.connect(self.export_history)
        btn_layout.addWidget(export_btn)

        clear_btn = QPushButton("נקה היסטוריה")
        clear_btn.setObjectName("clearBtn")
        clear_btn.clicked.connect(self.clear_history)
        btn_layout.addWidget(clear_btn)

        btn_layout.addStretch()

        self.history_count = QLabel(f"{len(self.history)} רשומות")
        btn_layout.addWidget(self.history_count)

        # עדכון הטבלה אחרי יצירת כל הרכיבים
        self.update_history_table()

        layout.addLayout(btn_layout)

        widget.setLayout(layout)
        return widget

    def create_settings_tab(self):
        """טאב הגדרות"""
        widget = QWidget()
        layout = QVBoxLayout()

        # הגדרות דפדפן
        browser_group = QGroupBox("דפדפן")
        browser_layout = QHBoxLayout()
        browser_layout.addWidget(QLabel("בחר דפדפן:"))
        self.browser_combo = QComboBox()
        self.browser_combo.addItem("Chrome", "Chrome")
        self.browser_combo.addItem("Edge", "Edge")
        self.browser_combo.currentIndexChanged.connect(self.browser_changed)
        browser_layout.addWidget(self.browser_combo)
        browser_layout.addStretch()
        browser_group.setLayout(browser_layout)
        layout.addWidget(browser_group)

        # הגדרות השהיה
        delay_group = QGroupBox("השהיות בין הודעות")
        delay_layout = QGridLayout()

        delay_layout.addWidget(QLabel("השהיה מינימלית (שניות):"), 0, 0)
        self.min_delay_spin = QSpinBox()
        self.min_delay_spin.setRange(1, 60)
        self.min_delay_spin.setValue(5)
        delay_layout.addWidget(self.min_delay_spin, 0, 1)

        delay_layout.addWidget(QLabel("השהיה מקסימלית (שניות):"), 1, 0)
        self.max_delay_spin = QSpinBox()
        self.max_delay_spin.setRange(1, 120)
        self.max_delay_spin.setValue(10)
        delay_layout.addWidget(self.max_delay_spin, 1, 1)

        delay_layout.addWidget(QLabel("הפסקה אחרי (מספר הודעות):"), 2, 0)
        self.pause_after_spin = QSpinBox()
        self.pause_after_spin.setRange(0, 500)
        self.pause_after_spin.setValue(30)
        delay_layout.addWidget(self.pause_after_spin, 2, 1)

        delay_layout.addWidget(QLabel("משך הפסקה (שניות):"), 3, 0)
        self.pause_duration_spin = QSpinBox()
        self.pause_duration_spin.setRange(1, 600)
        self.pause_duration_spin.setValue(60)
        delay_layout.addWidget(self.pause_duration_spin, 3, 1)

        delay_group.setLayout(delay_layout)
        layout.addWidget(delay_group)

        # הגדרות נוספות
        options_group = QGroupBox("אפשרויות נוספות")
        options_layout = QVBoxLayout()

        self.sound_check = QCheckBox("השמע צליל בסיום שליחה")
        self.sound_check.setChecked(True)
        options_layout.addWidget(self.sound_check)

        self.confirm_check = QCheckBox("בקש אישור לפני שליחה המונית")
        self.confirm_check.setChecked(True)
        options_layout.addWidget(self.confirm_check)

        self.save_history_check = QCheckBox("שמור היסטוריית שליחות")
        self.save_history_check.setChecked(True)
        options_layout.addWidget(self.save_history_check)

        options_group.setLayout(options_layout)
        layout.addWidget(options_group)

        layout.addStretch()

        widget.setLayout(layout)
        return widget

    # -------------------------------------------------------------------
    # פונקציות עזר
    # -------------------------------------------------------------------
    def format_selection(self, prefix, suffix):
        cursor = self.message_input.textCursor()
        selected = cursor.selectedText()
        if selected:
            cursor.insertText(prefix + selected + suffix)
        else:
            cursor.insertText(prefix + suffix)
        self.message_input.setFocus()

    def detect_links(self):
        text = self.message_input.toPlainText()
        links = re.findall(r'(https?://[^\s]+)', text)
        if links:
            reply = QMessageBox.question(
                self, "קישורים זוהו",
                f"נמצאו {len(links)} קישורים.\n\nהאם לכלול תצוגה מקדימה?",
                QMessageBox.Yes | QMessageBox.No
            )
            self.link_preview_requested = (reply == QMessageBox.Yes)
            self.log_message(f"תצוגה מקדימה: {'כן' if self.link_preview_requested else 'לא'}")
        else:
            QMessageBox.information(self, "קישורים", "לא זוהו קישורים")

    def browser_changed(self):
        self.browser_choice = self.browser_combo.currentData()
        self.log_message(f"דפדפן: {self.browser_choice}")

    def log_message(self, msg):
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        self.log_area.append(f"[{timestamp}] {msg}")

    def update_stats(self):
        self.contacts_stat.setText(str(len(self.contacts)))
        self.sent_stat.setText(str(self.session_sent))
        self.failed_stat.setText(str(self.session_failed))
        self.quick_stats.setText(f"נשלחו היום: {self.session_sent} | סה\"כ: {self.total_sent}")

    # -------------------------------------------------------------------
    # ניהול אנשי קשר
    # -------------------------------------------------------------------
    def upload_contacts(self):
        fileName, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ", "",
            "Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json)"
        )
        if fileName:
            try:
                new_contacts = []
                new_data = {}

                if fileName.endswith(".xlsx"):
                    df = pd.read_excel(fileName, engine="openpyxl")
                    for _, row in df.iterrows():
                        number = str(row.iloc[0]).strip()
                        if len(row) > 1:
                            name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                            new_data[number] = {"שם": name}
                        new_contacts.append(number)
                elif fileName.endswith(".csv"):
                    df = pd.read_csv(fileName)
                    for _, row in df.iterrows():
                        number = str(row.iloc[0]).strip()
                        if len(row) > 1:
                            name = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                            new_data[number] = {"שם": name}
                        new_contacts.append(number)
                elif fileName.endswith(".json"):
                    with open(fileName, "r", encoding="utf-8") as f:
                        data = json.load(f)
                        if isinstance(data, list):
                            new_contacts = data
                        elif isinstance(data, dict):
                            new_contacts = list(data.keys())
                            new_data = data

                # הוסף לתיבה לעיבוד
                for num in new_contacts:
                    if num in new_data and "שם" in new_data[num]:
                        self.manual_text.append(f"{num},{new_data[num]['שם']}")
                    else:
                        self.manual_text.append(num)

                self.log_message(f"נטענו {len(new_contacts)} אנשי קשר")
            except Exception as e:
                QMessageBox.critical(self, "שגיאה", f"שגיאה בטעינה:\n{e}")

    def load_group_participants(self):
        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()

            QMessageBox.information(self, "הוראות", "פתח קבוצה בוואצאפ ולחץ OK")

            span = WebDriverWait(self.bot.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'copyable-text') and @title]"))
            )
            title = span.get_attribute("title")

            numbers = re.findall(r"\+[\d\s()-]+", title)
            clean = []
            for n in numbers:
                n = re.sub(r"[\s()-]", "", n)
                if n.startswith("+"):
                    clean.append(n)

            self.manual_text.clear()
            for n in clean:
                self.manual_text.append(n)

            self.log_message(f"נטענו {len(clean)} מספרים מהקבוצה")
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה:\n{e}")

    def process_contacts(self):
        raw = self.manual_text.toPlainText().splitlines()
        processed = []

        for line in raw:
            line = line.strip()
            if not line:
                continue

            # בדיקה אם יש שם
            parts = line.split(",")
            num = parts[0].strip()
            name = parts[1].strip() if len(parts) > 1 else ""

            # נרמול
            if num.startswith("+972"):
                pass
            elif num.startswith("972"):
                num = "+" + num
            elif num.startswith("0"):
                num = "+972" + num[1:]
            else:
                num = "+972" + num

            # בדיקת תקינות
            digits_only = re.sub(r"\D", "", num)
            if len(digits_only) in [12, 13] and num not in [c[0] for c in processed]:
                processed.append((num, name))
                if name:
                    self.contacts_data[num] = {"שם": name}

        self.contacts = [p[0] for p in processed]
        self.update_contacts_table()
        self.update_stats()
        self.log_message(f"עובדו {len(self.contacts)} מספרים")

    def update_contacts_table(self):
        self.contacts_table.setRowCount(len(self.contacts))
        for i, num in enumerate(self.contacts):
            self.contacts_table.setItem(i, 0, QTableWidgetItem(num))
            name = self.contacts_data.get(num, {}).get("שם", "")
            self.contacts_table.setItem(i, 1, QTableWidgetItem(name))

            # כפתורי פעולה
            btn_widget = QWidget()
            btn_layout = QHBoxLayout()
            btn_layout.setContentsMargins(2, 2, 2, 2)

            blacklist_btn = QPushButton("חסום")
            blacklist_btn.setMaximumWidth(60)
            blacklist_btn.clicked.connect(lambda c, n=num: self.add_to_blacklist_direct(n))
            btn_layout.addWidget(blacklist_btn)

            remove_btn = QPushButton("הסר")
            remove_btn.setMaximumWidth(50)
            remove_btn.clicked.connect(lambda c, n=num: self.remove_contact(n))
            btn_layout.addWidget(remove_btn)

            btn_widget.setLayout(btn_layout)
            self.contacts_table.setCellWidget(i, 2, btn_widget)

        self.contacts_count_label.setText(f"{len(self.contacts)} אנשי קשר")
        self.manual_text.clear()

    def remove_contact(self, number):
        if number in self.contacts:
            self.contacts.remove(number)
            self.update_contacts_table()
            self.update_stats()

    def clear_contacts(self):
        if QMessageBox.question(self, "אישור", "למחוק את כל אנשי הקשר?") == QMessageBox.Yes:
            self.contacts = []
            self.contacts_data = {}
            self.update_contacts_table()
            self.update_stats()

    def export_contacts(self):
        if not self.contacts:
            QMessageBox.warning(self, "שגיאה", "אין אנשי קשר לייצוא")
            return

        fileName, _ = QFileDialog.getSaveFileName(
            self, "שמור קובץ", "contacts.xlsx", "Excel Files (*.xlsx)"
        )
        if fileName:
            data = []
            for num in self.contacts:
                name = self.contacts_data.get(num, {}).get("שם", "")
                data.append({"מספר": num, "שם": name})
            df = pd.DataFrame(data)
            df.to_excel(fileName, index=False)
            self.log_message(f"אנשי קשר יוצאו ל-{fileName}")

    # -------------------------------------------------------------------
    # ניהול תבניות
    # -------------------------------------------------------------------
    def update_templates_list(self):
        self.templates_list.clear()
        for name in self.templates.keys():
            self.templates_list.addItem(name)

    def update_template_combo(self):
        self.template_combo.clear()
        self.template_combo.addItem("-- בחר תבנית --")
        for name in self.templates.keys():
            self.template_combo.addItem(name)

    def add_template(self):
        dialog = TemplateDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            name, text = dialog.get_data()
            if name and text:
                self.templates[name] = text
                self.update_templates_list()
                self.update_template_combo()
                self.save_data()
                self.log_message(f"תבנית '{name}' נשמרה")

    def edit_template(self, item=None):
        if item is None:
            item = self.templates_list.currentItem()
        if item is None:
            return

        name = item.text()
        text = self.templates.get(name, "")
        dialog = TemplateDialog(self, name, text)
        if dialog.exec_() == QDialog.Accepted:
            new_name, new_text = dialog.get_data()
            if new_name != name:
                del self.templates[name]
            self.templates[new_name] = new_text
            self.update_templates_list()
            self.update_template_combo()
            self.save_data()

    def delete_template(self):
        item = self.templates_list.currentItem()
        if item:
            name = item.text()
            if QMessageBox.question(self, "מחיקה", f"למחוק את התבנית '{name}'?") == QMessageBox.Yes:
                del self.templates[name]
                self.update_templates_list()
                self.update_template_combo()
                self.save_data()

    def show_template_preview(self, current, previous):
        if current:
            text = self.templates.get(current.text(), "")
            self.template_preview.setPlainText(text)

    def use_selected_template(self):
        item = self.templates_list.currentItem()
        if item:
            text = self.templates.get(item.text(), "")
            self.message_input.setPlainText(text)
            self.tabs.setCurrentIndex(0)

    def load_selected_template(self, text):
        if text and text != "-- בחר תבנית --":
            template_text = self.templates.get(text, "")
            self.message_input.setPlainText(template_text)

    # -------------------------------------------------------------------
    # ניהול רשימה שחורה
    # -------------------------------------------------------------------
    def add_to_blacklist(self):
        number = self.blacklist_input.text().strip()
        if number:
            # נרמול
            if not number.startswith("+"):
                if number.startswith("0"):
                    number = "+972" + number[1:]
                else:
                    number = "+972" + number

            if number not in self.blacklist:
                self.blacklist.append(number)
                self.update_blacklist_display()
                self.save_data()
            self.blacklist_input.clear()

    def add_to_blacklist_direct(self, number):
        if number not in self.blacklist:
            self.blacklist.append(number)
            self.update_blacklist_display()
            self.save_data()
            self.log_message(f"{number} נוסף לרשימה השחורה")

    def remove_from_blacklist(self):
        item = self.blacklist_widget.currentItem()
        if item:
            number = item.text()
            self.blacklist.remove(number)
            self.update_blacklist_display()
            self.save_data()

    def clear_blacklist(self):
        if QMessageBox.question(self, "אישור", "לנקות את הרשימה השחורה?") == QMessageBox.Yes:
            self.blacklist = []
            self.update_blacklist_display()
            self.save_data()

    def import_blacklist(self):
        fileName, _ = QFileDialog.getOpenFileName(self, "בחר קובץ", "", "Text Files (*.txt);;All Files (*)")
        if fileName:
            with open(fileName, "r") as f:
                for line in f:
                    num = line.strip()
                    if num and num not in self.blacklist:
                        self.blacklist.append(num)
            self.update_blacklist_display()
            self.save_data()

    def update_blacklist_display(self):
        self.blacklist_widget.clear()
        for num in self.blacklist:
            self.blacklist_widget.addItem(num)

    # -------------------------------------------------------------------
    # תזמון
    # -------------------------------------------------------------------
    def toggle_schedule(self):
        if self.scheduler_thread and self.scheduler_thread.isRunning():
            self.scheduler_thread.stop()
            self.scheduler_thread.wait()
            self.scheduler_thread = None
            self.schedule_btn.setText("הפעל תזמון")
            self.schedule_status.setText("לא פעיל")
            self.countdown_label.setText("")
        else:
            target = self.schedule_datetime.dateTime()
            if target <= QDateTime.currentDateTime():
                QMessageBox.warning(self, "שגיאה", "בחר זמן עתידי")
                return

            self.scheduler_thread = SchedulerThread(target)
            self.scheduler_thread.ready_to_send.connect(self.scheduled_send)
            self.scheduler_thread.time_remaining.connect(self.update_countdown)
            self.scheduler_thread.start()

            self.schedule_btn.setText("בטל תזמון")
            self.schedule_status.setText(f"מתוזמן ל-{target.toString('dd/MM/yyyy HH:mm')}")

    def update_countdown(self, text):
        self.countdown_label.setText(text)

    def scheduled_send(self):
        self.schedule_btn.setText("הפעל תזמון")
        self.schedule_status.setText("מתחיל שליחה...")
        self.countdown_label.setText("")
        self.send_all_messages()

    # -------------------------------------------------------------------
    # היסטוריה
    # -------------------------------------------------------------------
    def update_history_table(self):
        self.history_table.setRowCount(len(self.history))
        for i, record in enumerate(reversed(self.history)):
            self.history_table.setItem(i, 0, QTableWidgetItem(record.get("date", "")))
            self.history_table.setItem(i, 1, QTableWidgetItem(record.get("number", "")))
            status = "הצלחה" if record.get("status") == "success" else "כישלון"
            self.history_table.setItem(i, 2, QTableWidgetItem(status))
            self.history_table.setItem(i, 3, QTableWidgetItem(record.get("message", "")[:50]))

        self.history_count.setText(f"{len(self.history)} רשומות")

    def add_to_history(self, number, success, message):
        if self.save_history_check.isChecked():
            self.history.append({
                "date": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                "number": number,
                "status": "success" if success else "failed",
                "message": message[:100]
            })
            self.update_history_table()

    def export_history(self):
        if not self.history:
            QMessageBox.warning(self, "שגיאה", "אין היסטוריה לייצוא")
            return

        fileName, _ = QFileDialog.getSaveFileName(
            self, "שמור קובץ", "history.xlsx", "Excel Files (*.xlsx)"
        )
        if fileName:
            df = pd.DataFrame(self.history)
            df.to_excel(fileName, index=False)
            self.log_message(f"היסטוריה יוצאה ל-{fileName}")

    def clear_history(self):
        if QMessageBox.question(self, "אישור", "לנקות את ההיסטוריה?") == QMessageBox.Yes:
            self.history = []
            self.update_history_table()
            self.save_data()

    # -------------------------------------------------------------------
    # שליחת הודעות
    # -------------------------------------------------------------------
    def send_test_message(self):
        number = self.test_input.text().strip()
        message = self.message_input.toPlainText().strip()

        if not number:
            QMessageBox.warning(self, "שגיאה", "הזן מספר לבדיקה")
            return
        if not message:
            QMessageBox.warning(self, "שגיאה", "כתוב הודעה")
            return

        # נרמול
        if not number.startswith("+"):
            if number.startswith("0"):
                number = "+972" + number[1:]
            else:
                number = "+972" + number

        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", str(e))
            return

        simulate = self.simulate_check.isChecked()
        success = self.bot.send_message(number, message, "text", simulate)

        if success:
            QMessageBox.information(self, "הצלחה", f"{'[סימולציה] ' if simulate else ''}הודעה נשלחה!")
        else:
            QMessageBox.critical(self, "שגיאה", "כשלון בשליחה")

    def send_all_messages(self):
        if not self.contacts:
            QMessageBox.warning(self, "שגיאה", "אין אנשי קשר")
            return

        message = self.message_input.toPlainText().strip()
        if not message:
            QMessageBox.warning(self, "שגיאה", "כתוב הודעה")
            return

        # אישור
        if self.confirm_check.isChecked():
            reply = QMessageBox.question(
                self, "אישור שליחה",
                f"לשלוח הודעה ל-{len(self.contacts)} אנשי קשר?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", str(e))
            return

        simulate = self.simulate_check.isChecked()

        # יצירת Thread
        self.message_thread = MessageSenderThread(
            self.contacts, message, "text",
            self.min_delay_spin.value(),
            self.max_delay_spin.value(),
            self.pause_after_spin.value(),
            self.pause_duration_spin.value(),
            self.bot,
            blacklist=self.blacklist,
            simulate=simulate,
            personalize_data=self.contacts_data
        )

        self.message_thread.update_log.connect(self.log_message)
        self.message_thread.finished_sending.connect(self.handle_finished)
        self.message_thread.update_count.connect(self.update_progress)
        self.message_thread.message_sent.connect(
            lambda num, success: self.add_to_history(num, success, message)
        )

        self.message_thread.start()

        # עדכון ממשק
        self.send_all_btn.setEnabled(False)
        self.pause_btn.setEnabled(True)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("שולח..." + (" [סימולציה]" if simulate else ""))
        self.progress_bar.setMaximum(len(self.contacts))

    def update_progress(self, sent, total):
        self.progress_bar.setValue(sent)
        self.session_sent = sent - self.message_thread.fail_count
        self.session_failed = self.message_thread.fail_count
        self.update_stats()
        self.status_label.setText(f"נשלחו {sent} מתוך {total}")

    def handle_finished(self):
        self.send_all_btn.setEnabled(True)
        self.pause_btn.setEnabled(False)
        self.stop_btn.setEnabled(False)
        self.status_label.setText("הושלם!")

        # עדכון סטטיסטיקות
        if self.message_thread:
            self.total_sent += self.message_thread.sent_count - self.message_thread.fail_count
            self.total_failed += self.message_thread.fail_count

        self.update_stats()
        self.save_data()

        # צליל
        if self.sound_check.isChecked() and sys.platform == 'win32':
            try:
                import winsound
                winsound.MessageBeep()
            except:
                pass

        # איפוס
        self.contacts = []
        self.update_contacts_table()
        self.paused = False
        self.pause_btn.setText("השהה")

        QMessageBox.information(self, "סיום", "השליחה הושלמה!")

    def toggle_pause(self):
        if not self.message_thread:
            return

        if not self.paused:
            self.message_thread.pause()
            self.pause_btn.setText("המשך")
            self.paused = True
            self.status_label.setText("מושהה")
        else:
            self.message_thread.resume()
            self.pause_btn.setText("השהה")
            self.paused = False
            self.status_label.setText("שולח...")

    def stop_sending(self):
        if self.message_thread:
            self.message_thread.stop()
            self.handle_finished()

    def closeEvent(self, event):
        self.save_data()
        if self.message_thread:
            self.message_thread.stop()
        if self.scheduler_thread:
            self.scheduler_thread.stop()
        if self.bot:
            self.bot.quit()
        super().closeEvent(event)


def main():
    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setStyleSheet(MAIN_STYLE)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
