"""Advanced logging system with rotation and colored console output."""

import logging
import os
from logging.handlers import RotatingFileHandler
from typing import Optional
import colorlog
from datetime import datetime


class Logger:
    """Advanced logger with file rotation and colored console output."""

    def __init__(
        self,
        name: str,
        log_dir: str = 'logs',
        level: str = 'INFO',
        max_bytes: int = 10485760,
        backup_count: int = 5,
        console: bool = True,
        file: bool = True
    ):
        """Initialize logger.

        Args:
            name: Logger name
            log_dir: Directory for log files
            level: Logging level (DEBUG, INFO, WARNING, ERROR, CRITICAL)
            max_bytes: Maximum size of log file before rotation
            backup_count: Number of backup files to keep
            console: Enable console logging
            file: Enable file logging
        """
        self.name = name
        self.log_dir = log_dir
        self.logger = logging.getLogger(name)

        # Set level
        numeric_level = getattr(logging, level.upper(), logging.INFO)
        self.logger.setLevel(numeric_level)

        # Remove existing handlers
        self.logger.handlers.clear()

        # Create log directory
        os.makedirs(log_dir, exist_ok=True)

        # Console handler with colors
        if console:
            console_handler = colorlog.StreamHandler()
            console_handler.setLevel(numeric_level)

            console_formatter = colorlog.ColoredFormatter(
                '%(log_color)s%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S',
                log_colors={
                    'DEBUG': 'cyan',
                    'INFO': 'green',
                    'WARNING': 'yellow',
                    'ERROR': 'red',
                    'CRITICAL': 'red,bg_white',
                }
            )
            console_handler.setFormatter(console_formatter)
            self.logger.addHandler(console_handler)

        # File handler with rotation
        if file:
            log_file = os.path.join(log_dir, f'{name}.log')
            file_handler = RotatingFileHandler(
                log_file,
                maxBytes=max_bytes,
                backupCount=backup_count,
                encoding='utf-8'
            )
            file_handler.setLevel(numeric_level)

            file_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)

    def debug(self, message: str, *args, **kwargs) -> None:
        """Log debug message."""
        self.logger.debug(message, *args, **kwargs)

    def info(self, message: str, *args, **kwargs) -> None:
        """Log info message."""
        self.logger.info(message, *args, **kwargs)

    def warning(self, message: str, *args, **kwargs) -> None:
        """Log warning message."""
        self.logger.warning(message, *args, **kwargs)

    def error(self, message: str, *args, **kwargs) -> None:
        """Log error message."""
        self.logger.error(message, *args, **kwargs)

    def critical(self, message: str, *args, **kwargs) -> None:
        """Log critical message."""
        self.logger.critical(message, *args, **kwargs)

    def exception(self, message: str, *args, **kwargs) -> None:
        """Log exception with traceback."""
        self.logger.exception(message, *args, **kwargs)


class CampaignLogger(Logger):
    """Specialized logger for campaign tracking."""

    def __init__(self, campaign_id: int, campaign_name: str, log_dir: str = 'logs'):
        """Initialize campaign logger.

        Args:
            campaign_id: Campaign ID
            campaign_name: Campaign name
            log_dir: Directory for log files
        """
        self.campaign_id = campaign_id
        self.campaign_name = campaign_name

        # Create campaign-specific log directory
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.campaign_log_dir = os.path.join(log_dir, 'campaigns', f'{campaign_id}_{timestamp}')
        os.makedirs(self.campaign_log_dir, exist_ok=True)

        super().__init__(
            name=f'campaign_{campaign_id}',
            log_dir=self.campaign_log_dir,
            level='INFO'
        )

        self.log_message_sent = []
        self.log_message_failed = []

    def message_sent(self, phone: str, message: str) -> None:
        """Log successful message send.

        Args:
            phone: Phone number
            message: Message content (truncated in log)
        """
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] SENT to {phone}: {message[:50]}..."
        self.log_message_sent.append(log_entry)
        self.info(f"Message sent to {phone}")

    def message_failed(self, phone: str, error: str) -> None:
        """Log failed message send.

        Args:
            phone: Phone number
            error: Error message
        """
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] FAILED to {phone}: {error}"
        self.log_message_failed.append(log_entry)
        self.error(f"Message failed to {phone}: {error}")

    def save_report(self) -> str:
        """Save campaign report to file.

        Returns:
            Path to report file
        """
        report_path = os.path.join(self.campaign_log_dir, 'campaign_report.txt')

        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(f"Campaign Report: {self.campaign_name}\n")
            f.write(f"Campaign ID: {self.campaign_id}\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")

            f.write(f"Successful Messages: {len(self.log_message_sent)}\n")
            f.write(f"Failed Messages: {len(self.log_message_failed)}\n\n")

            if self.log_message_sent:
                f.write("SUCCESSFUL MESSAGES\n")
                f.write("-" * 80 + "\n")
                for entry in self.log_message_sent:
                    f.write(entry + "\n")
                f.write("\n")

            if self.log_message_failed:
                f.write("FAILED MESSAGES\n")
                f.write("-" * 80 + "\n")
                for entry in self.log_message_failed:
                    f.write(entry + "\n")

        return report_path


# Global logger instances
_loggers: dict[str, Logger] = {}


def get_logger(name: str, **kwargs) -> Logger:
    """Get or create a logger instance.

    Args:
        name: Logger name
        **kwargs: Additional logger configuration

    Returns:
        Logger instance
    """
    if name not in _loggers:
        _loggers[name] = Logger(name, **kwargs)
    return _loggers[name]
