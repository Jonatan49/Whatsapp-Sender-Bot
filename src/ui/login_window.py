"""Login window for WhatsApp Sender Bot."""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QMessageBox, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor
from src.services.auth_service import AuthService
from src.core.database import get_db
from src.utils.logger import get_logger


class LoginWindow(QDialog):
    """Login dialog window."""

    login_successful = pyqtSignal(object)  # Emits the logged-in user

    def __init__(self, parent=None):
        """Initialize login window."""
        super().__init__(parent)
        self.logger = get_logger('LoginWindow')
        self.auth_service = AuthService()
        self.user = None

        self.init_ui()

    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle('Login - WhatsApp Bot')
        self.setFixedSize(450, 400)
        self.setModal(True)

        # Main layout
        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(50, 40, 50, 40)

        # Header with icon
        header_layout = QVBoxLayout()
        header_layout.setSpacing(10)

        # Title
        title = QLabel('WhatsApp Bot')
        title_font = QFont('Segoe UI', 20, QFont.Bold)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(title)

        # Subtitle
        subtitle = QLabel('Professional Message Sender')
        subtitle_font = QFont('Segoe UI', 10)
        subtitle.setFont(subtitle_font)
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet('color: #666; margin-bottom: 10px;')
        header_layout.addWidget(subtitle)

        layout.addLayout(header_layout)
        layout.addSpacing(20)

        # Login form
        form_layout = QVBoxLayout()
        form_layout.setSpacing(12)

        # Username section
        username_label = QLabel('Username')
        username_label.setFont(QFont('Segoe UI', 10))
        username_label.setStyleSheet('color: #333; font-weight: 500;')
        form_layout.addWidget(username_label)

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText('Enter your username')
        self.username_input.setText('admin')
        self.username_input.setFont(QFont('Segoe UI', 11))
        self.username_input.setMinimumHeight(40)
        self.username_input.returnPressed.connect(self.handle_login)
        form_layout.addWidget(self.username_input)

        form_layout.addSpacing(5)

        # Password section
        password_label = QLabel('Password')
        password_label.setFont(QFont('Segoe UI', 10))
        password_label.setStyleSheet('color: #333; font-weight: 500;')
        form_layout.addWidget(password_label)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText('Enter your password')
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setFont(QFont('Segoe UI', 11))
        self.password_input.setMinimumHeight(40)
        self.password_input.returnPressed.connect(self.handle_login)
        form_layout.addWidget(self.password_input)

        layout.addLayout(form_layout)
        layout.addSpacing(10)

        # Login button
        self.login_button = QPushButton('LOGIN')
        self.login_button.setMinimumHeight(45)
        self.login_button.setFont(QFont('Segoe UI', 11, QFont.Bold))
        self.login_button.clicked.connect(self.handle_login)
        self.login_button.setDefault(True)
        self.login_button.setCursor(Qt.PointingHandCursor)
        layout.addWidget(self.login_button)

        layout.addSpacing(10)

        # Info label
        info_label = QLabel('Default credentials: admin / admin123')
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setFont(QFont('Segoe UI', 9))
        info_label.setStyleSheet('color: #999; padding: 5px;')
        layout.addWidget(info_label)

        # Version
        version_label = QLabel('v2.0.0')
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setFont(QFont('Segoe UI', 8))
        version_label.setStyleSheet('color: #ccc; margin-top: 5px;')
        layout.addWidget(version_label)

        self.setLayout(layout)

        # Apply styling
        self.apply_styles()

        # Set focus to password field if username is filled
        if self.username_input.text():
            self.password_input.setFocus()

    def apply_styles(self):
        """Apply custom styles to the window."""
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
            }
            QLabel {
                color: #333;
            }
            QLineEdit {
                padding: 10px 12px;
                border: 2px solid #e0e0e0;
                border-radius: 6px;
                background-color: #fafafa;
                font-size: 13px;
                color: #333;
            }
            QLineEdit:focus {
                border: 2px solid #25D366;
                background-color: white;
            }
            QLineEdit:hover {
                border: 2px solid #bbb;
            }
            QPushButton {
                background-color: #25D366;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px;
                font-size: 13px;
                font-weight: bold;
                letter-spacing: 1px;
            }
            QPushButton:hover {
                background-color: #1fb855;
            }
            QPushButton:pressed {
                background-color: #1a9d47;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888;
            }
        """)

    def handle_login(self):
        """Handle login button click."""
        username = self.username_input.text().strip()
        password = self.password_input.text()

        if not username or not password:
            QMessageBox.warning(
                self,
                'Login Error',
                'Please enter both username and password.'
            )
            return

        # Disable button during login
        self.login_button.setEnabled(False)
        self.login_button.setText('Logging in...')

        try:
            # Authenticate user
            self.user = self.auth_service.authenticate(username, password)

            if self.user:
                self.logger.info(f"User '{username}' logged in successfully")

                # Check if account is locked
                if not self.user.is_active:
                    QMessageBox.warning(
                        self,
                        'Account Locked',
                        'Your account has been locked. Please contact an administrator.'
                    )
                    self.login_button.setEnabled(True)
                    self.login_button.setText('Login')
                    return

                # Emit signal and accept dialog
                self.login_successful.emit(self.user)
                self.accept()
            else:
                self.logger.warning(f"Failed login attempt for user '{username}'")
                QMessageBox.critical(
                    self,
                    'Login Failed',
                    'Invalid username or password.\n\n'
                    'Your account will be locked after 5 failed attempts.'
                )
                self.password_input.clear()
                self.password_input.setFocus()

        except Exception as e:
            self.logger.exception(f"Login error: {e}")
            QMessageBox.critical(
                self,
                'Error',
                f'An error occurred during login:\n{str(e)}'
            )
        finally:
            self.login_button.setEnabled(True)
            self.login_button.setText('Login')
