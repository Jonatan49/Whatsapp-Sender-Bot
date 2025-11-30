#!/usr/bin/env python3
"""WhatsApp Sender Bot Pro - Main entry point."""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon
from src.core.database import init_db
from src.core.config import get_config
from src.utils.logger import get_logger

# Import the original app for now (we can upgrade UI incrementally)
# For now, let's create a simple launcher that uses the upgraded backend


def main():
    """Main application entry point."""
    # Initialize logger
    logger = get_logger('main')
    logger.info("=" * 80)
    logger.info("WhatsApp Sender Bot Pro v2.0.0")
    logger.info("=" * 80)

    # Load configuration
    logger.info("Loading configuration...")
    config = get_config()

    # Initialize database
    logger.info("Initializing database...")
    db = init_db()

    stats = db.get_statistics()
    logger.info(f"Database ready: {stats}")

    # Create Qt Application
    app = QApplication(sys.argv)

    # Set application metadata
    app.setApplicationName(config.app_name)
    app.setApplicationVersion(config.app_version)
    app.setOrganizationName("Yonatan Cohen")

    logger.info("Starting application...")

    try:
        # Import UI components
        from src.ui import LoginWindow, MainWindow

        # Show login window
        login_window = LoginWindow()

        # Connect login success to main window
        def on_login_success(user):
            logger.info(f"Login successful for user: {user.username}")
            main_window = MainWindow(user)
            main_window.show()

        login_window.login_successful.connect(on_login_success)

        # Show login dialog
        if login_window.exec_() == LoginWindow.Accepted:
            logger.info("Application started successfully")
            logger.info("Use Ctrl+C to exit")

            # Run application
            sys.exit(app.exec_())
        else:
            logger.info("Login cancelled by user")
            sys.exit(0)

    except KeyboardInterrupt:
        logger.info("Application interrupted by user")
        sys.exit(0)
    except Exception as e:
        logger.exception(f"Application error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
