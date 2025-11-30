"""Configuration management with environment variables and YAML."""

import os
import yaml
from typing import Any, Optional
from dataclasses import dataclass, field
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


@dataclass
class BrowserConfig:
    """Browser configuration."""
    default: str = 'chrome'
    headless: bool = False
    detach: bool = True
    user_agent: str = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'


@dataclass
class WhatsAppConfig:
    """WhatsApp configuration."""
    base_url: str = 'https://web.whatsapp.com'
    qr_timeout: int = 120
    message_timeout: int = 20
    default_country_code: str = '+972'


@dataclass
class DelayConfig:
    """Delay and timing configuration."""
    min_delay: float = 5.0
    max_delay: float = 10.0
    pause_after: int = 100
    pause_duration: int = 120
    typing_speed_min: float = 0.05
    typing_speed_max: float = 0.15


@dataclass
class AntiDetectionConfig:
    """Anti-detection configuration."""
    enabled: bool = True
    random_mouse_movements: bool = True
    random_scrolling: bool = True
    viewport_randomization: bool = True
    retry_attempts: int = 3
    retry_delay: int = 5


@dataclass
class SecurityConfig:
    """Security configuration."""
    password_hash_rounds: int = 12
    session_timeout: int = 3600
    max_login_attempts: int = 5
    lockout_duration: int = 300
    encrypt_database: bool = True
    encrypt_logs: bool = True
    secret_key: str = field(default_factory=lambda: os.getenv('SECRET_KEY', 'change-me-in-production'))


@dataclass
class LoggingConfig:
    """Logging configuration."""
    level: str = 'INFO'
    format: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    max_bytes: int = 10485760
    backup_count: int = 5
    encrypt: bool = True
    console: bool = True
    file: bool = True


@dataclass
class DatabaseConfig:
    """Database configuration."""
    url: str = field(default_factory=lambda: os.getenv('DATABASE_URL', 'sqlite:///data/whatsapp_bot.db'))
    echo: bool = False
    pool_size: int = 5


class Config:
    """Main configuration class."""

    def __init__(self, config_path: str = 'config/config.yaml'):
        """Initialize configuration from YAML file and environment variables.

        Args:
            config_path: Path to YAML configuration file
        """
        self.config_path = config_path
        self._config_data = {}

        # Load YAML configuration if exists
        if os.path.exists(config_path):
            with open(config_path, 'r', encoding='utf-8') as f:
                self._config_data = yaml.safe_load(f) or {}

        # Initialize configuration sections
        self.browser = self._load_browser_config()
        self.whatsapp = self._load_whatsapp_config()
        self.delays = self._load_delay_config()
        self.anti_detection = self._load_anti_detection_config()
        self.security = self._load_security_config()
        self.logging = self._load_logging_config()
        self.database = self._load_database_config()

        # Application settings
        self.app_name = self.get('application.name', 'WhatsApp Sender Bot Pro')
        self.app_version = self.get('application.version', '2.0.0')
        self.default_language = self.get('application.language', 'he')
        self.default_theme = self.get('application.theme', 'whatsapp')

        # Folders
        self.folders = {
            'data': self.get('folders.data', 'data'),
            'logs': self.get('folders.logs', 'logs'),
            'backups': self.get('folders.backups', 'data/backups'),
            'exports': self.get('folders.exports', 'data/exports'),
            'temp': self.get('folders.temp', 'data/temp')
        }

        # Create necessary folders
        self._create_folders()

    def get(self, key: str, default: Any = None) -> Any:
        """Get configuration value using dot notation.

        Args:
            key: Configuration key in dot notation (e.g., 'browser.default')
            default: Default value if key not found

        Returns:
            Configuration value or default
        """
        keys = key.split('.')
        value = self._config_data

        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default

        return value

    def _load_browser_config(self) -> BrowserConfig:
        """Load browser configuration."""
        return BrowserConfig(
            default=self.get('browser.default', 'chrome'),
            headless=self.get('browser.headless', False),
            detach=self.get('browser.detach', True),
            user_agent=self.get('browser.user_agent', BrowserConfig.user_agent)
        )

    def _load_whatsapp_config(self) -> WhatsAppConfig:
        """Load WhatsApp configuration."""
        return WhatsAppConfig(
            base_url=self.get('whatsapp.base_url', 'https://web.whatsapp.com'),
            qr_timeout=int(os.getenv('WHATSAPP_QR_TIMEOUT', self.get('whatsapp.qr_timeout', 120))),
            message_timeout=int(os.getenv('WHATSAPP_MESSAGE_TIMEOUT', self.get('whatsapp.message_timeout', 20))),
            default_country_code=os.getenv('WHATSAPP_DEFAULT_COUNTRY_CODE', self.get('whatsapp.default_country_code', '+972'))
        )

    def _load_delay_config(self) -> DelayConfig:
        """Load delay configuration."""
        return DelayConfig(
            min_delay=float(os.getenv('MIN_DELAY', self.get('delays.min_delay', 5.0))),
            max_delay=float(os.getenv('MAX_DELAY', self.get('delays.max_delay', 10.0))),
            pause_after=int(os.getenv('PAUSE_AFTER', self.get('delays.pause_after', 100))),
            pause_duration=int(os.getenv('PAUSE_DURATION', self.get('delays.pause_duration', 120))),
            typing_speed_min=self.get('delays.typing_speed.min', 0.05),
            typing_speed_max=self.get('delays.typing_speed.max', 0.15)
        )

    def _load_anti_detection_config(self) -> AntiDetectionConfig:
        """Load anti-detection configuration."""
        return AntiDetectionConfig(
            enabled=self.get('anti_detection.enabled', True),
            random_mouse_movements=self.get('anti_detection.random_mouse_movements', True),
            random_scrolling=self.get('anti_detection.random_scrolling', True),
            viewport_randomization=self.get('anti_detection.viewport_randomization', True),
            retry_attempts=self.get('anti_detection.retry_attempts', 3),
            retry_delay=self.get('anti_detection.retry_delay', 5)
        )

    def _load_security_config(self) -> SecurityConfig:
        """Load security configuration."""
        return SecurityConfig(
            password_hash_rounds=self.get('security.password_hash_rounds', 12),
            session_timeout=self.get('security.session_timeout', 3600),
            max_login_attempts=self.get('security.max_login_attempts', 5),
            lockout_duration=self.get('security.lockout_duration', 300),
            encrypt_database=self.get('security.encrypt_database', True),
            encrypt_logs=self.get('security.encrypt_logs', True),
            secret_key=os.getenv('SECRET_KEY', 'change-me-in-production')
        )

    def _load_logging_config(self) -> LoggingConfig:
        """Load logging configuration."""
        return LoggingConfig(
            level=os.getenv('LOG_LEVEL', self.get('logging.level', 'INFO')),
            format=os.getenv('LOG_FORMAT', self.get('logging.format', LoggingConfig.format)),
            max_bytes=int(os.getenv('LOG_MAX_BYTES', self.get('logging.max_bytes', 10485760))),
            backup_count=int(os.getenv('LOG_BACKUP_COUNT', self.get('logging.backup_count', 5))),
            encrypt=self.get('logging.encrypt', True),
            console=self.get('logging.console', True),
            file=self.get('logging.file', True)
        )

    def _load_database_config(self) -> DatabaseConfig:
        """Load database configuration."""
        return DatabaseConfig(
            url=os.getenv('DATABASE_URL', self.get('database.path', 'sqlite:///data/whatsapp_bot.db')),
            echo=self.get('database.echo', False),
            pool_size=self.get('database.pool_size', 5)
        )

    def _create_folders(self) -> None:
        """Create necessary application folders."""
        for folder in self.folders.values():
            os.makedirs(folder, exist_ok=True)

    def save(self, config_path: str = None) -> None:
        """Save current configuration to YAML file.

        Args:
            config_path: Path to save configuration (defaults to current path)
        """
        if config_path is None:
            config_path = self.config_path

        os.makedirs(os.path.dirname(config_path), exist_ok=True)

        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(self._config_data, f, default_flow_style=False, allow_unicode=True)


# Global configuration instance
_config_instance: Optional[Config] = None


def get_config() -> Config:
    """Get or create global configuration instance."""
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    return _config_instance
