#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
בוט שליחת הודעות וואצאפ
© Yonatan Cohen
"""

import sys
import time
import datetime
import random
import os
import re
import shutil
import json

import pandas as pd

try:
    import pyperclip
except ImportError:
    pyperclip = None

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit, QMessageBox,
    QComboBox, QRadioButton, QSpinBox, QGroupBox, QGridLayout, QListWidget,
    QProgressBar, QFrame, QSplashScreen
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPixmap, QColor, QPalette

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# -------------------------------------------------------------------
# עיצוב מודרני בסגנון וואצאפ
# -------------------------------------------------------------------
MAIN_STYLE = """
QMainWindow {
    background-color: #0b141a;
}
QWidget {
    background-color: #0b141a;
    font-family: "Segoe UI", "Arial", sans-serif;
    color: #e9edef;
    font-size: 13px;
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
    padding: 12px 24px;
    font-size: 14px;
    font-weight: bold;
    color: white;
    min-height: 20px;
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
QPushButton#stopButton {
    background-color: #ea4335;
}
QPushButton#stopButton:hover {
    background-color: #f55246;
}
QPushButton#pauseButton {
    background-color: #fbbc04;
    color: #1a1a1a;
}
QPushButton#pauseButton:hover {
    background-color: #ffc929;
}
QLineEdit, QTextEdit {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 10px;
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
}
QLabel#titleLabel {
    font-size: 24px;
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
QComboBox {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 8px;
    color: #e9edef;
    min-width: 120px;
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
QSpinBox {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 8px;
    padding: 8px;
    color: #e9edef;
    min-width: 80px;
}
QSpinBox:focus {
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
}
QListWidget::item:selected {
    background-color: #00a884;
    color: white;
}
QListWidget::item:hover {
    background-color: #2a3942;
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
QRadioButton {
    color: #e9edef;
    spacing: 8px;
}
QRadioButton::indicator {
    width: 18px;
    height: 18px;
}
QRadioButton::indicator:checked {
    background-color: #00a884;
    border: 2px solid #00a884;
    border-radius: 9px;
}
QRadioButton::indicator:unchecked {
    background-color: #1f2c33;
    border: 2px solid #2a3942;
    border-radius: 9px;
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

    def send_message(self, number, message, mode="text"):
        try:
            # סינון תווים בעייתיים
            message = ''.join(ch for ch in message if ord(ch) <= 0xFFFF)

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
                self.log("הודעה הודבקה עם תצוגה מקדימה")
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
    update_count = pyqtSignal(int, int)  # sent, total

    def __init__(self, contacts, message, mode, min_delay, max_delay, pause_after, pause_duration, bot):
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
        total = len(self.contacts)
        for i, number in enumerate(self.contacts, start=1):
            if self._stop_flag:
                self.update_log.emit("השליחה נעצרה על ידי המשתמש")
                break

            while self._pause_flag and not self._stop_flag:
                time.sleep(0.5)

            if self._stop_flag:
                break

            timestamp = datetime.datetime.now().strftime('%H:%M:%S')
            success = self.bot.send_message(number, self.message, self.mode)

            if not success:
                self.fail_count += 1
                self.update_log.emit(f"[{timestamp}] נכשל: {number}")
            else:
                self.update_log.emit(f"[{timestamp}] נשלח: {number}")

            self.sent_count += 1
            self.update_count.emit(self.sent_count, total)

            if self.pause_after > 0 and self.sent_count % self.pause_after == 0:
                self.update_log.emit(f"הפסקה של {self.pause_duration} שניות...")
                time.sleep(self.pause_duration)

            if i < total:
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
        self.update_log.emit(f"סיום! נשלחו {self.sent_count} הודעות, {self.fail_count} כשלונות ({int(total_time)} שניות)")
        self.finished_sending.emit()

    def stop(self):
        self._stop_flag = True

    def pause(self):
        self._pause_flag = True

    def resume(self):
        self._pause_flag = False


# -------------------------------------------------------------------
# חלון ראשי
# -------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("בוט שליחת הודעות וואצאפ")
        self.setGeometry(100, 100, 1100, 750)
        self.setLayoutDirection(Qt.RightToLeft)

        self.contacts = []
        self.base_folder = os.path.dirname(os.path.abspath(__file__))
        self.browser_choice = "Chrome"
        self.bot = WhatsAppBot(self.log_message, parent=self)
        self.message_thread = None
        self.paused = False
        self.link_preview_requested = False
        self.current_job_folder = None

        self.initUI()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # כותרת
        header_layout = QVBoxLayout()
        title_label = QLabel("בוט שליחת הודעות וואצאפ")
        title_label.setObjectName("titleLabel")
        title_label.setAlignment(Qt.AlignCenter)

        subtitle_label = QLabel("שלח הודעות בקלות ובמהירות")
        subtitle_label.setObjectName("subtitleLabel")
        subtitle_label.setAlignment(Qt.AlignCenter)

        header_layout.addWidget(title_label)
        header_layout.addWidget(subtitle_label)
        main_layout.addLayout(header_layout)

        # תוכן עיקרי - שתי עמודות
        content_layout = QHBoxLayout()
        content_layout.setSpacing(20)

        # עמודה ימנית - אנשי קשר והודעה
        right_column = QVBoxLayout()
        right_column.setSpacing(15)

        # קבוצת אנשי קשר
        contacts_group = QGroupBox("אנשי קשר")
        contacts_layout = QVBoxLayout()
        contacts_layout.setSpacing(10)

        buttons_layout = QHBoxLayout()
        self.upload_button = QPushButton("העלה קובץ אנשי קשר")
        self.upload_button.clicked.connect(self.upload_contacts)
        buttons_layout.addWidget(self.upload_button)

        self.load_title_button = QPushButton("טען מקבוצה פתוחה")
        self.load_title_button.clicked.connect(self.load_group_participants_via_title)
        buttons_layout.addWidget(self.load_title_button)

        contacts_layout.addLayout(buttons_layout)

        self.manual_text = QTextEdit()
        self.manual_text.setPlaceholderText("הדבק מספרים כאן (כל מספר בשורה נפרדת)\nלדוגמה:\n0541234567\n0521234567")
        self.manual_text.setMaximumHeight(120)
        contacts_layout.addWidget(self.manual_text)

        self.process_button = QPushButton("עבד אנשי קשר")
        self.process_button.clicked.connect(self.process_contacts)
        contacts_layout.addWidget(self.process_button)

        self.contacts_status = QLabel("0 אנשי קשר מוכנים")
        self.contacts_status.setAlignment(Qt.AlignCenter)
        contacts_layout.addWidget(self.contacts_status)

        contacts_group.setLayout(contacts_layout)
        right_column.addWidget(contacts_group)

        # קבוצת הודעה
        message_group = QGroupBox("תוכן ההודעה")
        message_layout = QVBoxLayout()
        message_layout.setSpacing(10)

        # כפתורי עיצוב
        format_layout = QHBoxLayout()
        bold_btn = QPushButton("מודגש")
        bold_btn.setMaximumWidth(80)
        bold_btn.clicked.connect(lambda: self.format_selection("*", "*"))
        italic_btn = QPushButton("נטוי")
        italic_btn.setMaximumWidth(80)
        italic_btn.clicked.connect(lambda: self.format_selection("_", "_"))
        strike_btn = QPushButton("קו חוצה")
        strike_btn.setMaximumWidth(80)
        strike_btn.clicked.connect(lambda: self.format_selection("~", "~"))

        self.link_btn = QPushButton("זיהוי קישור")
        self.link_btn.setMaximumWidth(100)
        self.link_btn.clicked.connect(self.detect_links)

        format_layout.addWidget(bold_btn)
        format_layout.addWidget(italic_btn)
        format_layout.addWidget(strike_btn)
        format_layout.addWidget(self.link_btn)
        format_layout.addStretch()

        message_layout.addLayout(format_layout)

        self.message_input = QTextEdit()
        self.message_input.setPlaceholderText("כתוב את ההודעה כאן...")
        self.message_input.setMinimumHeight(150)
        message_layout.addWidget(self.message_input)

        message_group.setLayout(message_layout)
        right_column.addWidget(message_group)

        content_layout.addLayout(right_column, 1)

        # עמודה שמאלית - הגדרות ולוג
        left_column = QVBoxLayout()
        left_column.setSpacing(15)

        # הגדרות
        settings_group = QGroupBox("הגדרות")
        settings_layout = QGridLayout()
        settings_layout.setSpacing(10)

        # דפדפן
        settings_layout.addWidget(QLabel("דפדפן:"), 0, 0)
        self.browser_combo = QComboBox()
        self.browser_combo.addItem("Chrome", "Chrome")
        self.browser_combo.addItem("Edge", "Edge")
        self.browser_combo.currentIndexChanged.connect(self.browser_changed)
        settings_layout.addWidget(self.browser_combo, 0, 1)

        # השהיות
        settings_layout.addWidget(QLabel("השהיה מינימלית (שניות):"), 1, 0)
        self.min_delay_spin = QSpinBox()
        self.min_delay_spin.setRange(1, 60)
        self.min_delay_spin.setValue(5)
        settings_layout.addWidget(self.min_delay_spin, 1, 1)

        settings_layout.addWidget(QLabel("השהיה מקסימלית (שניות):"), 2, 0)
        self.max_delay_spin = QSpinBox()
        self.max_delay_spin.setRange(1, 120)
        self.max_delay_spin.setValue(10)
        settings_layout.addWidget(self.max_delay_spin, 2, 1)

        settings_layout.addWidget(QLabel("הפסקה אחרי (מספר הודעות):"), 3, 0)
        self.pause_after_spin = QSpinBox()
        self.pause_after_spin.setRange(0, 1000)
        self.pause_after_spin.setValue(50)
        settings_layout.addWidget(self.pause_after_spin, 3, 1)

        settings_layout.addWidget(QLabel("משך הפסקה (שניות):"), 4, 0)
        self.pause_duration_spin = QSpinBox()
        self.pause_duration_spin.setRange(1, 600)
        self.pause_duration_spin.setValue(60)
        settings_layout.addWidget(self.pause_duration_spin, 4, 1)

        settings_group.setLayout(settings_layout)
        left_column.addWidget(settings_group)

        # לוג
        log_group = QGroupBox("יומן פעילות")
        log_layout = QVBoxLayout()

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMinimumHeight(150)
        log_layout.addWidget(self.log_area)

        # פס התקדמות
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%v / %m הודעות")
        log_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("מוכן לשליחה")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignCenter)
        log_layout.addWidget(self.status_label)

        log_group.setLayout(log_layout)
        left_column.addWidget(log_group)

        content_layout.addLayout(left_column, 1)
        main_layout.addLayout(content_layout)

        # כפתורי פעולה
        action_layout = QHBoxLayout()
        action_layout.setSpacing(15)

        self.test_input = QLineEdit()
        self.test_input.setPlaceholderText("מספר לבדיקה (לדוגמה: 0541234567)")
        self.test_input.setMaximumWidth(200)
        action_layout.addWidget(self.test_input)

        self.test_button = QPushButton("שלח הודעת בדיקה")
        self.test_button.clicked.connect(self.send_test_message)
        action_layout.addWidget(self.test_button)

        action_layout.addStretch()

        self.pause_button = QPushButton("השהה")
        self.pause_button.setObjectName("pauseButton")
        self.pause_button.clicked.connect(self.toggle_pause)
        self.pause_button.setEnabled(False)
        action_layout.addWidget(self.pause_button)

        self.stop_button = QPushButton("עצור")
        self.stop_button.setObjectName("stopButton")
        self.stop_button.clicked.connect(self.stop_sending)
        self.stop_button.setEnabled(False)
        action_layout.addWidget(self.stop_button)

        self.send_all_button = QPushButton("התחל שליחה לכולם")
        self.send_all_button.setMinimumWidth(200)
        self.send_all_button.clicked.connect(self.send_all_messages)
        action_layout.addWidget(self.send_all_button)

        main_layout.addLayout(action_layout)

        # כותרת תחתית
        footer_label = QLabel("© Yonatan Cohen")
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setStyleSheet("color: #4a5568; font-size: 11px;")
        main_layout.addWidget(footer_label)

        central_widget.setLayout(main_layout)

    def browser_changed(self):
        self.browser_choice = self.browser_combo.currentData()
        self.log_message(f"דפדפן נבחר: {self.browser_choice}")

    def format_selection(self, prefix, suffix):
        cursor = self.message_input.textCursor()
        selected_text = cursor.selectedText()
        if selected_text:
            cursor.insertText(prefix + selected_text + suffix)
        else:
            cursor.insertText(prefix + suffix)
        self.message_input.setFocus()

    def detect_links(self):
        text = self.message_input.toPlainText()
        pattern = r'(https?://[^\s]+)'
        links = re.findall(pattern, text)
        if links:
            reply = QMessageBox.question(
                self, "קישורים זוהו",
                f"נמצאו {len(links)} קישורים.\n\nהאם לכלול תצוגה מקדימה?",
                QMessageBox.Yes | QMessageBox.No
            )
            self.link_preview_requested = (reply == QMessageBox.Yes)
            status = "כן" if self.link_preview_requested else "לא"
            self.log_message(f"תצוגה מקדימה לקישור: {status}")
        else:
            QMessageBox.information(self, "קישורים", "לא זוהו קישורים בהודעה")
            self.link_preview_requested = False

    def load_group_participants_via_title(self):
        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()

            span_element = WebDriverWait(self.bot.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[contains(@class, 'copyable-text') and @title]"))
            )
            title_content = span_element.get_attribute("title")
            self.log_message(f"טוען מספרים מהקבוצה...")

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

            self.log_message(f"נטענו {len(clean_numbers)} מספרים מהקבוצה")
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בטעינת מספרים:\n{e}")

    def upload_contacts(self):
        fileName, _ = QFileDialog.getOpenFileName(
            self, "בחר קובץ אנשי קשר", "",
            "Excel Files (*.xlsx);;CSV Files (*.csv);;JSON Files (*.json)"
        )
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

                self.log_message(f"נטענו {len(new_contacts)} אנשי קשר מהקובץ")
            except Exception as e:
                QMessageBox.critical(self, "שגיאה", f"שגיאה בטעינת הקובץ:\n{e}")

    def process_contacts(self):
        raw_numbers = self.manual_text.toPlainText().splitlines()
        processed = []

        for num in raw_numbers:
            num = num.strip()
            if not num:
                continue

            # נרמול מספרים ישראליים
            if num.startswith("+972"):
                pass
            elif num.startswith("972"):
                num = "+" + num
            elif num.startswith("0"):
                num = "+972" + num[1:]
            else:
                num = "+972" + num

            # בדיקת תקינות
            rest = num[4:] if num.startswith("+972") else num
            if rest.startswith("5") or rest.startswith("05"):
                if len(num) in [13, 14]:
                    processed.append(num)
                else:
                    self.log_message(f"מספר לא תקין (אורך): {num}")
            else:
                self.log_message(f"מספר לא תקין (לא נייד): {num}")

        # הסרת כפילויות
        unique = list(dict.fromkeys(processed))
        self.contacts = unique

        self.manual_text.clear()
        for num in self.contacts:
            self.manual_text.append(num)

        self.contacts_status.setText(f"{len(self.contacts)} אנשי קשר מוכנים")
        self.progress_bar.setMaximum(len(self.contacts) if self.contacts else 1)
        self.log_message(f"עובדו {len(self.contacts)} מספרים תקינים")

    def send_test_message(self):
        test_number = self.test_input.text().strip()
        if not test_number:
            QMessageBox.warning(self, "שגיאה", "הזן מספר לבדיקה")
            return

        message = self.message_input.toPlainText().strip()
        if not message:
            QMessageBox.warning(self, "שגיאה", "כתוב הודעה לשליחה")
            return

        # נרמול מספר
        if not test_number.startswith("+"):
            if test_number.startswith("0"):
                test_number = "+972" + test_number[1:]
            else:
                test_number = "+972" + test_number

        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בפתיחת WhatsApp:\n{e}")
            return

        clean_message = ''.join(ch for ch in message if ord(ch) <= 0xFFFF)
        success = self.bot.send_message(test_number, clean_message, "text")

        if success:
            QMessageBox.information(self, "הצלחה", f"הודעת בדיקה נשלחה ל-{test_number}")
        else:
            QMessageBox.critical(self, "שגיאה", "כשלון בשליחת הודעת הבדיקה")

    def send_all_messages(self):
        if not self.contacts:
            QMessageBox.warning(self, "שגיאה", "אין אנשי קשר לשליחה.\nעבד אנשי קשר קודם.")
            return

        message = self.message_input.toPlainText().strip()
        if not message:
            QMessageBox.warning(self, "שגיאה", "כתוב הודעה לשליחה")
            return

        try:
            if self.bot.driver is None:
                self.bot.open_whatsapp()
        except Exception as e:
            QMessageBox.critical(self, "שגיאה", f"שגיאה בפתיחת WhatsApp:\n{e}")
            return

        # יצירת תיקיית משימה
        daily_folder = os.path.join(self.base_folder, "logs", datetime.datetime.now().strftime("%Y-%m-%d"))
        os.makedirs(daily_folder, exist_ok=True)
        self.current_job_folder = os.path.join(daily_folder, datetime.datetime.now().strftime("%H-%M-%S"))
        os.makedirs(self.current_job_folder, exist_ok=True)

        # הכנת הודעה
        clean_message = ''.join(ch for ch in message if ord(ch) <= 0xFFFF)

        # הפעלת Thread
        self.message_thread = MessageSenderThread(
            self.contacts, clean_message, "text",
            self.min_delay_spin.value(),
            self.max_delay_spin.value(),
            self.pause_after_spin.value(),
            self.pause_duration_spin.value(),
            self.bot
        )

        self.message_thread.update_log.connect(self.log_message)
        self.message_thread.finished_sending.connect(self.handle_finished)
        self.message_thread.update_count.connect(self.update_progress)

        self.message_thread.start()

        # עדכון ממשק
        self.send_all_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.stop_button.setEnabled(True)
        self.status_label.setText("שולח הודעות...")

        self.log_message(f"מתחיל שליחה ל-{len(self.contacts)} אנשי קשר")

    def update_progress(self, sent, total):
        self.progress_bar.setValue(sent)
        self.progress_bar.setMaximum(total)
        self.status_label.setText(f"נשלחו {sent} מתוך {total} הודעות")

    def handle_finished(self):
        self.send_all_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.status_label.setText("השליחה הושלמה!")

        # שמירת דוח
        if self.current_job_folder and self.message_thread:
            summary = f"""דוח שליחה
============
תאריך: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}
הודעות שנשלחו: {self.message_thread.sent_count}
כשלונות: {self.message_thread.fail_count}
זמן כולל: {int(time.time() - self.message_thread.start_time)} שניות
"""
            report_file = os.path.join(self.current_job_folder, "report.txt")
            with open(report_file, "w", encoding="utf-8") as f:
                f.write(summary)
            self.log_message(f"דוח נשמר ב: {report_file}")

        # איפוס
        self.contacts = []
        self.manual_text.clear()
        self.contacts_status.setText("0 אנשי קשר מוכנים")
        self.paused = False
        self.pause_button.setText("השהה")

        QMessageBox.information(self, "סיום", "השליחה הושלמה בהצלחה!")

    def toggle_pause(self):
        if not self.message_thread:
            return

        if not self.paused:
            self.message_thread.pause()
            self.pause_button.setText("המשך")
            self.paused = True
            self.status_label.setText("מושהה")
            self.log_message("השליחה הושהתה")
        else:
            self.message_thread.resume()
            self.pause_button.setText("השהה")
            self.paused = False
            self.status_label.setText("שולח הודעות...")
            self.log_message("השליחה ממשיכה")

    def stop_sending(self):
        if self.message_thread:
            self.message_thread.stop()
            self.log_message("השליחה נעצרה")

            self.send_all_button.setEnabled(True)
            self.pause_button.setEnabled(False)
            self.stop_button.setEnabled(False)
            self.status_label.setText("השליחה נעצרה")

            self.contacts = []
            self.manual_text.clear()
            self.contacts_status.setText("0 אנשי קשר מוכנים")
            self.paused = False
            self.pause_button.setText("השהה")

    def log_message(self, msg):
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        self.log_area.append(f"[{timestamp}] {msg}")

    def closeEvent(self, event):
        if self.message_thread:
            self.message_thread.stop()
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
