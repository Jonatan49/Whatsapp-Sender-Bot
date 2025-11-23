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
        self.setWindowTitle('WhatsApp Sender Bot Pro - Login')
        self.setFixedSize(400, 300)
        self.setModal(True)

        # Main layout
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(40, 40, 40, 40)

        # Title
        title = QLabel('WhatsApp Sender Bot Pro')
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Subtitle
        subtitle = QLabel('v2.0.0')
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet('color: #666;')
        layout.addWidget(subtitle)

        # Spacer
        layout.addSpacing(20)

        # Username
        username_label = QLabel('Username:')
        layout.addWidget(username_label)

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText('Enter username')
        self.username_input.setText('admin')  # Default for convenience
        self.username_input.returnPressed.connect(self.handle_login)
        layout.addWidget(self.username_input)

        # Password
        password_label = QLabel('Password:')
        layout.addWidget(password_label)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText('Enter password')
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.returnPressed.connect(self.handle_login)
        layout.addWidget(self.password_input)

        # Spacer
        layout.addSpacing(10)

        # Login button
        self.login_button = QPushButton('Login')
        self.login_button.setMinimumHeight(40)
        self.login_button.clicked.connect(self.handle_login)
        self.login_button.setDefault(True)
        layout.addWidget(self.login_button)

        # Info label
        info_label = QLabel('Default: admin / admin123')
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet('color: #999; font-size: 10px;')
        layout.addWidget(info_label)

        self.setLayout(layout)

        # Apply styling
        self.apply_styles()

    def apply_styles(self):
        """Apply custom styles to the window."""
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #333;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background-color: white;
                font-size: 12px;
            }
            QLineEdit:focus {
                border: 1px solid #25D366;
            }
            QPushButton {
                background-color: #25D366;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #22c55e;
            }
            QPushButton:pressed {
                background-color: #20b558;
            }
            QPushButton:disabled {
                background-color: #ccc;
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
