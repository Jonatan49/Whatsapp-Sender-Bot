"""Enhanced WhatsApp bot with retry logic and anti-detection features."""

import time
import random
from typing import Optional, Callable
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type
import pyperclip

from src.utils.logger import get_logger
from src.core.config import get_config


class WhatsAppBot:
    """Enhanced WhatsApp Web automation bot with retry logic and anti-detection."""

    def __init__(self, log_callback: Optional[Callable] = None):
        """Initialize WhatsApp bot.

        Args:
            log_callback: Optional callback function for logging to UI
        """
        self.config = get_config()
        self.logger = get_logger('whatsapp_bot')
        self.log_callback = log_callback
        self.driver: Optional[webdriver.Chrome | webdriver.Edge] = None
        self.is_logged_in = False

    def log(self, message: str, level: str = 'INFO') -> None:
        """Log message to logger and optional UI callback.

        Args:
            message: Log message
            level: Log level (INFO, WARNING, ERROR, etc.)
        """
        # Log to file
        getattr(self.logger, level.lower(), self.logger.info)(message)

        # Log to UI callback if provided
        if self.log_callback:
            self.log_callback(message, level)

    def init_driver(self, browser: str = None, version: str = None) -> None:
        """Initialize browser driver with anti-detection features.

        Args:
            browser: Browser choice ('chrome' or 'edge')
            version: Specific browser version (optional)
        """
        try:
            if browser is None:
                browser = self.config.browser.default

            self.log(f"Initializing {browser} browser...")

            if browser.lower() == 'chrome':
                self._init_chrome_driver(version)
            elif browser.lower() == 'edge':
                self._init_edge_driver(version)
            else:
                raise ValueError(f"Unsupported browser: {browser}")

            # Apply anti-detection measures
            self._apply_anti_detection()

            self.log(f"Browser initialized successfully: {browser}")

        except Exception as e:
            self.log(f"Failed to initialize browser: {e}", level='ERROR')
            raise

    def _init_chrome_driver(self, version: Optional[str] = None) -> None:
        """Initialize Chrome driver."""
        from selenium.webdriver.chrome.service import Service as ChromeService
        from webdriver_manager.chrome import ChromeDriverManager

        options = webdriver.ChromeOptions()

        # Anti-detection options
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # Performance options
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--no-sandbox')

        # User agent
        if self.config.browser.user_agent:
            options.add_argument(f'user-agent={self.config.browser.user_agent}')

        # Headless mode
        if self.config.browser.headless:
            options.add_argument('--headless=new')

        # Detach browser (keeps it open after script ends)
        if self.config.browser.detach:
            options.add_experimental_option("detach", True)

        # Install and create driver
        if version:
            driver_path = ChromeDriverManager(version=version).install()
        else:
            driver_path = ChromeDriverManager().install()

        service = ChromeService(driver_path)
        self.driver = webdriver.Chrome(service=service, options=options)

    def _init_edge_driver(self, version: Optional[str] = None) -> None:
        """Initialize Edge driver."""
        from selenium.webdriver.edge.service import Service as EdgeService
        from webdriver_manager.microsoft import EdgeChromiumDriverManager

        options = webdriver.EdgeOptions()

        # Anti-detection options
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        # User agent
        if self.config.browser.user_agent:
            options.add_argument(f'user-agent={self.config.browser.user_agent}')

        # Headless mode
        if self.config.browser.headless:
            options.add_argument('--headless=new')

        # Detach browser
        if self.config.browser.detach:
            options.add_experimental_option("detach", True)

        # Install and create driver
        if version:
            driver_path = EdgeChromiumDriverManager(version=version).install()
        else:
            driver_path = EdgeChromiumDriverManager().install()

        service = EdgeService(executable_path=driver_path)
        self.driver = webdriver.Edge(service=service, options=options)

    def _apply_anti_detection(self) -> None:
        """Apply anti-detection JavaScript to hide automation."""
        if not self.driver:
            return

        # Remove webdriver property
        self.driver.execute_script("""
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
        """)

        # Override navigator properties
        self.driver.execute_script("""
            Object.defineProperty(navigator, 'plugins', {
                get: () => [1, 2, 3, 4, 5]
            });
        """)

        # Randomize viewport if enabled
        if self.config.anti_detection.viewport_randomization:
            width = random.randint(1200, 1920)
            height = random.randint(800, 1080)
            self.driver.set_window_size(width, height)

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((TimeoutException, WebDriverException))
    )
    def open_whatsapp(self, qr_callback: Optional[Callable] = None) -> None:
        """Open WhatsApp Web and wait for login.

        Args:
            qr_callback: Optional callback when QR code scan is needed
        """
        if self.driver is None:
            self.init_driver()

        try:
            self.log("Opening WhatsApp Web...")
            self.driver.get(self.config.whatsapp.base_url)

            # Wait for either QR code or chat list (already logged in)
            self.log("Waiting for login...")

            # Check if already logged in
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@id='side']"))
                )
                self.is_logged_in = True
                self.log("Already logged in to WhatsApp Web")
                return
            except TimeoutException:
                pass

            # Need to scan QR code
            self.log("Please scan QR code...")

            if qr_callback:
                qr_callback()

            # Wait for login (check for side panel)
            WebDriverWait(self.driver, self.config.whatsapp.qr_timeout).until(
                EC.presence_of_element_located((By.XPATH, "//div[@id='side']"))
            )

            self.is_logged_in = True
            self.log("Successfully logged in to WhatsApp Web")

            # Random delay to appear human-like
            time.sleep(random.uniform(1, 3))

        except TimeoutException:
            self.log("QR code scan timeout", level='ERROR')
            raise
        except Exception as e:
            self.log(f"Error opening WhatsApp Web: {e}", level='ERROR')
            raise

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((TimeoutException, NoSuchElementException))
    )
    def send_message(
        self,
        phone: str,
        message: str,
        link_preview: bool = False,
        typing_simulation: bool = True
    ) -> bool:
        """Send message to phone number with retry logic.

        Args:
            phone: Phone number in international format
            message: Message content
            link_preview: Enable link preview
            typing_simulation: Simulate human typing

        Returns:
            True if successful, False otherwise
        """
        try:
            # Ensure driver is initialized
            if self.driver is None:
                raise RuntimeError("Browser not initialized")

            # Filter non-BMP characters
            message = ''.join(ch for ch in message if ord(ch) <= 0xFFFF)

            # Navigate to chat
            self.log(f"Opening chat for {phone}...")
            self.driver.get(f"{self.config.whatsapp.base_url}/send?phone={phone}")

            # Wait for message input box
            self.log("Waiting for message box...")
            text_box = WebDriverWait(self.driver, self.config.whatsapp.message_timeout).until(
                EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='10']"))
            )

            # Random delay before typing
            time.sleep(random.uniform(0.5, 1.5))

            # Send message
            if link_preview and pyperclip:
                # Use clipboard for link preview
                pyperclip.copy(message)
                text_box.send_keys(Keys.CONTROL, 'v')
                self.log("Message pasted (link preview enabled)")
                time.sleep(3)  # Wait for preview to load
            else:
                if typing_simulation:
                    # Simulate human typing
                    self._type_like_human(text_box, message)
                else:
                    # Send directly
                    multiline_message = message.replace('\n', Keys.SHIFT + Keys.ENTER)
                    text_box.send_keys(multiline_message)

            # Random delay before sending
            time.sleep(random.uniform(0.3, 0.8))

            # Send the message
            text_box.send_keys(Keys.ENTER)

            self.log(f"Message sent to {phone}")

            # Random delay after sending
            time.sleep(random.uniform(0.5, 1.0))

            return True

        except TimeoutException:
            self.log(f"Timeout sending message to {phone}", level='ERROR')
            return False
        except Exception as e:
            self.log(f"Error sending message to {phone}: {e}", level='ERROR')
            return False

    def _type_like_human(self, element, text: str) -> None:
        """Simulate human typing with random delays.

        Args:
            element: Web element to type into
            text: Text to type
        """
        typing_speed_min = self.config.delays.typing_speed_min
        typing_speed_max = self.config.delays.typing_speed_max

        for char in text:
            if char == '\n':
                element.send_keys(Keys.SHIFT, Keys.ENTER)
            else:
                element.send_keys(char)

            # Random typing delay
            delay = random.uniform(typing_speed_min, typing_speed_max)
            time.sleep(delay)

    def random_scroll(self) -> None:
        """Perform random scrolling to appear human-like."""
        if not self.config.anti_detection.random_scrolling:
            return

        try:
            scroll_amount = random.randint(100, 500)
            self.driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
            time.sleep(random.uniform(0.5, 1.5))
        except Exception as e:
            self.log(f"Error during random scroll: {e}", level='WARNING')

    def is_browser_alive(self) -> bool:
        """Check if browser is still running.

        Returns:
            True if browser is alive, False otherwise
        """
        try:
            if self.driver is None:
                return False
            _ = self.driver.window_handles
            return True
        except Exception:
            return False

    def quit(self) -> None:
        """Close the browser."""
        if self.driver:
            try:
                self.driver.quit()
                self.log("Browser closed")
            except Exception as e:
                self.log(f"Error closing browser: {e}", level='WARNING')
            finally:
                self.driver = None
                self.is_logged_in = False

    def __del__(self) -> None:
        """Cleanup on deletion."""
        self.quit()
