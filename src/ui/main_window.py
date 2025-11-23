"""Main application window for WhatsApp Sender Bot."""

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QTabWidget, QTableWidget,
    QTableWidgetItem, QMessageBox, QHeaderView, QStatusBar,
    QAction, QMenu, QMenuBar
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QIcon
from src.core.database import get_db
from src.utils.logger import get_logger


class MainWindow(QMainWindow):
    """Main application window."""

    def __init__(self, user, parent=None):
        """Initialize main window.

        Args:
            user: Logged-in user object
        """
        super().__init__(parent)
        self.logger = get_logger('MainWindow')
        self.user = user
        self.db = get_db()

        self.init_ui()
        self.load_data()

    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle(f'WhatsApp Sender Bot Pro - {self.user.full_name}')
        self.setGeometry(100, 100, 1200, 800)

        # Create menu bar
        self.create_menu_bar()

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main layout
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)

        # Header
        header = self.create_header()
        layout.addWidget(header)

        # Statistics
        stats_widget = self.create_statistics()
        layout.addWidget(stats_widget)

        # Tab widget for different sections
        self.tabs = QTabWidget()
        self.tabs.addTab(self.create_dashboard_tab(), "Dashboard")
        self.tabs.addTab(self.create_contacts_tab(), "Contacts")
        self.tabs.addTab(self.create_campaigns_tab(), "Campaigns")
        self.tabs.addTab(self.create_messages_tab(), "Messages")
        self.tabs.addTab(self.create_templates_tab(), "Templates")
        layout.addWidget(self.tabs)

        central_widget.setLayout(layout)

        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage('Ready')

        # Apply styling
        self.apply_styles()

    def create_menu_bar(self):
        """Create menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu('File')

        refresh_action = QAction('Refresh', self)
        refresh_action.triggered.connect(self.load_data)
        file_menu.addAction(refresh_action)

        file_menu.addSeparator()

        exit_action = QAction('Exit', self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Help menu
        help_menu = menubar.addMenu('Help')

        about_action = QAction('About', self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def create_header(self):
        """Create header widget."""
        header = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 20)

        # Title
        title = QLabel('WhatsApp Sender Bot Pro')
        title_font = QFont()
        title_font.setPointSize(24)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        layout.addStretch()

        # User info
        user_label = QLabel(f'Logged in as: {self.user.username}')
        user_label.setStyleSheet('color: #666; font-size: 14px;')
        layout.addWidget(user_label)

        # Logout button
        logout_btn = QPushButton('Logout')
        logout_btn.clicked.connect(self.logout)
        layout.addWidget(logout_btn)

        header.setLayout(layout)
        return header

    def create_statistics(self):
        """Create statistics widget."""
        stats_widget = QWidget()
        layout = QHBoxLayout()
        layout.setSpacing(20)

        # Get statistics
        stats = self.db.get_statistics()

        # Create stat cards
        self.stat_cards = {
            'contacts': self.create_stat_card('Contacts', stats.get('contacts', 0), '#3b82f6'),
            'campaigns': self.create_stat_card('Campaigns', stats.get('campaigns', 0), '#8b5cf6'),
            'messages': self.create_stat_card('Messages', stats.get('messages', 0), '#10b981'),
            'templates': self.create_stat_card('Templates', stats.get('templates', 0), '#f59e0b'),
        }

        for card in self.stat_cards.values():
            layout.addWidget(card)

        layout.addStretch()
        stats_widget.setLayout(layout)
        return stats_widget

    def create_stat_card(self, title, value, color):
        """Create a statistics card."""
        card = QWidget()
        card.setMinimumSize(200, 100)
        card.setStyleSheet(f"""
            QWidget {{
                background-color: {color};
                border-radius: 8px;
                padding: 15px;
            }}
            QLabel {{
                color: white;
            }}
        """)

        layout = QVBoxLayout()

        title_label = QLabel(title)
        title_label.setStyleSheet('font-size: 14px; font-weight: normal;')
        layout.addWidget(title_label)

        value_label = QLabel(str(value))
        value_font = QFont()
        value_font.setPointSize(32)
        value_font.setBold(True)
        value_label.setFont(value_font)
        layout.addWidget(value_label)

        card.setLayout(layout)
        return card

    def create_dashboard_tab(self):
        """Create dashboard tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        welcome = QLabel(f'Welcome, {self.user.full_name}!')
        welcome_font = QFont()
        welcome_font.setPointSize(16)
        welcome.setFont(welcome_font)
        layout.addWidget(welcome)

        info = QLabel(
            'This is WhatsApp Sender Bot Pro v2.0.0\n\n'
            'Features:\n'
            '• Manage contacts and contact groups\n'
            '• Create and run message campaigns\n'
            '• Use customizable message templates\n'
            '• Track message delivery status\n'
            '• Secure authentication and user management\n\n'
            'Get started by adding contacts or creating a campaign!'
        )
        info.setWordWrap(True)
        info.setStyleSheet('color: #666; line-height: 1.6;')
        layout.addWidget(info)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def create_contacts_tab(self):
        """Create contacts tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Toolbar
        toolbar = QHBoxLayout()
        add_btn = QPushButton('Add Contact')
        add_btn.clicked.connect(self.add_contact)
        toolbar.addWidget(add_btn)

        import_btn = QPushButton('Import from Excel')
        import_btn.clicked.connect(self.import_contacts)
        toolbar.addWidget(import_btn)

        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Table
        self.contacts_table = QTableWidget()
        self.contacts_table.setColumnCount(4)
        self.contacts_table.setHorizontalHeaderLabels(['Name', 'Phone', 'Group', 'Status'])
        self.contacts_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.contacts_table)

        widget.setLayout(layout)
        return widget

    def create_campaigns_tab(self):
        """Create campaigns tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Toolbar
        toolbar = QHBoxLayout()
        new_btn = QPushButton('New Campaign')
        new_btn.clicked.connect(self.new_campaign)
        toolbar.addWidget(new_btn)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Table
        self.campaigns_table = QTableWidget()
        self.campaigns_table.setColumnCount(5)
        self.campaigns_table.setHorizontalHeaderLabels(['Name', 'Status', 'Progress', 'Created', 'Actions'])
        self.campaigns_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.campaigns_table)

        widget.setLayout(layout)
        return widget

    def create_messages_tab(self):
        """Create messages tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Table
        self.messages_table = QTableWidget()
        self.messages_table.setColumnCount(5)
        self.messages_table.setHorizontalHeaderLabels(['Contact', 'Message', 'Status', 'Sent At', 'Campaign'])
        self.messages_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.messages_table)

        widget.setLayout(layout)
        return widget

    def create_templates_tab(self):
        """Create templates tab."""
        widget = QWidget()
        layout = QVBoxLayout()

        # Toolbar
        toolbar = QHBoxLayout()
        new_btn = QPushButton('New Template')
        new_btn.clicked.connect(self.new_template)
        toolbar.addWidget(new_btn)
        toolbar.addStretch()
        layout.addLayout(toolbar)

        # Table
        self.templates_table = QTableWidget()
        self.templates_table.setColumnCount(4)
        self.templates_table.setHorizontalHeaderLabels(['Name', 'Category', 'Variables', 'Actions'])
        self.templates_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.templates_table)

        widget.setLayout(layout)
        return widget

    def apply_styles(self):
        """Apply custom styles."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f9fafb;
            }
            QPushButton {
                background-color: #25D366;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #22c55e;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #e5e7eb;
                border-radius: 4px;
            }
            QHeaderView::section {
                background-color: #f3f4f6;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                background-color: white;
                border-radius: 4px;
            }
            QTabBar::tab {
                background-color: #f3f4f6;
                padding: 10px 20px;
                border: none;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: white;
                border-bottom: 2px solid #25D366;
            }
        """)

    def load_data(self):
        """Load data from database."""
        self.logger.info("Loading data...")
        try:
            # Update statistics
            stats = self.db.get_statistics()
            # Update stat cards if they exist
            # (Would need to update the actual values here)

            self.statusBar.showMessage('Data loaded successfully', 3000)
        except Exception as e:
            self.logger.exception(f"Error loading data: {e}")
            QMessageBox.critical(self, 'Error', f'Failed to load data:\n{str(e)}')

    def add_contact(self):
        """Add new contact."""
        QMessageBox.information(self, 'Coming Soon', 'Add contact feature coming soon!')

    def import_contacts(self):
        """Import contacts from Excel."""
        QMessageBox.information(self, 'Coming Soon', 'Import contacts feature coming soon!')

    def new_campaign(self):
        """Create new campaign."""
        QMessageBox.information(self, 'Coming Soon', 'New campaign feature coming soon!')

    def new_template(self):
        """Create new template."""
        QMessageBox.information(self, 'Coming Soon', 'New template feature coming soon!')

    def logout(self):
        """Logout current user."""
        reply = QMessageBox.question(
            self,
            'Logout',
            'Are you sure you want to logout?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.logger.info(f"User '{self.user.username}' logged out")
            self.close()

    def show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self,
            'About WhatsApp Sender Bot Pro',
            'WhatsApp Sender Bot Pro v2.0.0\n\n'
            'A professional WhatsApp automation tool.\n\n'
            '© 2024 Yonatan Cohen\n'
            'All rights reserved.'
        )
