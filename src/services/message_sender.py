"""Message sender service with campaign management integration."""

import time
import random
from typing import Optional, Callable
from PyQt5.QtCore import QThread, pyqtSignal
from datetime import datetime

from src.core.bot import WhatsAppBot
from src.services.campaign_service import CampaignService
from src.models.message import MessageStatus
from src.models.campaign import CampaignStatus
from src.core.config import get_config
from src.utils.logger import get_logger, CampaignLogger


class MessageSenderThread(QThread):
    """Thread for sending messages asynchronously."""

    # Signals
    update_log = pyqtSignal(str, str)  # message, level
    update_progress = pyqtSignal(int, int)  # sent, total
    update_stats = pyqtSignal(dict)  # statistics
    finished = pyqtSignal()

    def __init__(
        self,
        campaign_id: int,
        bot: WhatsAppBot,
        log_callback: Optional[Callable] = None
    ):
        """Initialize message sender thread.

        Args:
            campaign_id: Campaign ID
            bot: WhatsAppBot instance
            log_callback: Optional callback for logging
        """
        super().__init__()

        self.campaign_id = campaign_id
        self.bot = bot
        self.config = get_config()
        self.campaign_service = CampaignService()
        self.logger = get_logger('message_sender')

        # Thread control
        self._stop_flag = False
        self._pause_flag = False

        # Statistics
        self.start_time = None
        self.sent_count = 0
        self.failed_count = 0
        self.total_count = 0

        # Campaign logger
        campaign = self.campaign_service.get_campaign(campaign_id)
        if campaign:
            self.campaign_logger = CampaignLogger(campaign_id, campaign.name)

    def run(self) -> None:
        """Run the message sending process."""
        try:
            self.log("Starting campaign...", "INFO")
            self.start_time = time.time()

            # Get campaign
            campaign = self.campaign_service.get_campaign(self.campaign_id)
            if not campaign:
                self.log("Campaign not found", "ERROR")
                return

            # Start campaign
            self.campaign_service.start_campaign(self.campaign_id)

            # Get pending messages
            pending_messages = self.campaign_service.get_pending_messages(self.campaign_id)
            self.total_count = len(pending_messages)

            self.log(f"Found {self.total_count} messages to send", "INFO")

            # Ensure WhatsApp Web is open
            if not self.bot.is_logged_in:
                self.log("Opening WhatsApp Web...", "INFO")
                self.bot.open_whatsapp()

            # Send messages
            for i, message in enumerate(pending_messages, start=1):
                # Check stop flag
                if self._stop_flag:
                    self.log("Campaign stopped by user", "WARNING")
                    break

                # Check pause flag
                while self._pause_flag and not self._stop_flag:
                    time.sleep(0.5)

                if self._stop_flag:
                    break

                # Send message
                self.log(f"Sending message {i}/{self.total_count} to {message.phone_number}", "INFO")

                # Update message status to SENDING
                self.campaign_service.update_message_status(message.id, MessageStatus.SENDING)

                # Send the message
                success = self.bot.send_message(
                    phone=message.phone_number,
                    message=message.content,
                    link_preview=campaign.enable_link_preview,
                    typing_simulation=True
                )

                # Update message status
                if success:
                    self.campaign_service.update_message_status(message.id, MessageStatus.SENT)
                    self.campaign_logger.message_sent(message.phone_number, message.content)
                    self.sent_count += 1
                    self.log(f"✓ Message sent to {message.phone_number}", "INFO")
                else:
                    self.campaign_service.update_message_status(
                        message.id,
                        MessageStatus.FAILED,
                        error_message="Failed to send"
                    )
                    self.campaign_logger.message_failed(message.phone_number, "Failed to send")
                    self.failed_count += 1
                    self.log(f"✗ Failed to send message to {message.phone_number}", "ERROR")

                # Update progress
                self.update_progress.emit(self.sent_count + self.failed_count, self.total_count)

                # Apply delay
                delay = self._calculate_delay()
                self.log(f"Waiting {delay:.1f} seconds before next message...", "DEBUG")
                self._sleep_with_pause_check(delay)

                # Check for pause after N messages
                if campaign.pause_after > 0 and (self.sent_count + self.failed_count) % campaign.pause_after == 0:
                    pause_duration = campaign.pause_duration
                    self.log(f"Pausing for {pause_duration} seconds after {campaign.pause_after} messages", "INFO")
                    self._sleep_with_pause_check(pause_duration)

            # Calculate total time
            total_time = time.time() - self.start_time

            # Complete campaign
            self.campaign_service.complete_campaign(self.campaign_id)

            # Get final statistics
            stats = self.campaign_service.get_campaign_statistics(self.campaign_id)
            stats['total_duration'] = total_time

            # Save campaign report
            report_path = self.campaign_logger.save_report()
            self.log(f"Campaign report saved to: {report_path}", "INFO")

            # Emit final statistics
            self.update_stats.emit(stats)

            # Log completion
            success_rate = (self.sent_count / self.total_count * 100) if self.total_count > 0 else 0
            self.log(
                f"Campaign completed: {self.sent_count}/{self.total_count} sent ({success_rate:.1f}%), "
                f"{self.failed_count} failed, Duration: {total_time:.1f}s",
                "INFO"
            )

        except Exception as e:
            self.log(f"Error in campaign execution: {e}", "ERROR")
            self.logger.exception("Campaign execution error")

        finally:
            self.finished.emit()

    def _calculate_delay(self) -> float:
        """Calculate random delay with human-like variance.

        Returns:
            Delay in seconds
        """
        min_delay = self.config.delays.min_delay
        max_delay = self.config.delays.max_delay

        # Base delay
        delay = random.uniform(min_delay, max_delay)

        # Add variance (10% of delay)
        if self.config.anti_detection.enabled:
            variance = delay * 0.1
            delay += random.uniform(-variance, variance)

        return max(1.0, delay)  # Minimum 1 second

    def _sleep_with_pause_check(self, duration: float) -> None:
        """Sleep with periodic pause flag checking.

        Args:
            duration: Sleep duration in seconds
        """
        steps = int(duration * 2)  # Check every 0.5 seconds
        step_duration = duration / steps if steps > 0 else duration

        for _ in range(steps):
            if self._stop_flag:
                break

            while self._pause_flag and not self._stop_flag:
                time.sleep(0.5)

            if self._stop_flag:
                break

            time.sleep(step_duration)

    def log(self, message: str, level: str = "INFO") -> None:
        """Log message.

        Args:
            message: Log message
            level: Log level
        """
        # Log to file
        getattr(self.logger, level.lower(), self.logger.info)(message)

        # Emit to UI
        self.update_log.emit(message, level)

    def stop(self) -> None:
        """Stop the sending process."""
        self._stop_flag = True
        self.log("Stop requested...", "WARNING")

    def pause(self) -> None:
        """Pause the sending process."""
        self._pause_flag = True
        self.log("Paused", "INFO")

    def resume(self) -> None:
        """Resume the sending process."""
        self._pause_flag = False
        self.log("Resumed", "INFO")

    def is_paused(self) -> bool:
        """Check if paused.

        Returns:
            True if paused, False otherwise
        """
        return self._pause_flag


class MessageSender:
    """High-level message sender service."""

    def __init__(self):
        """Initialize message sender service."""
        self.bot = WhatsAppBot()
        self.sender_thread: Optional[MessageSenderThread] = None

    def start_campaign(
        self,
        campaign_id: int,
        log_callback: Optional[Callable] = None,
        progress_callback: Optional[Callable] = None,
        stats_callback: Optional[Callable] = None,
        finished_callback: Optional[Callable] = None
    ) -> MessageSenderThread:
        """Start sending messages for a campaign.

        Args:
            campaign_id: Campaign ID
            log_callback: Callback for log messages
            progress_callback: Callback for progress updates
            stats_callback: Callback for statistics updates
            finished_callback: Callback when finished

        Returns:
            MessageSenderThread instance
        """
        # Create sender thread
        self.sender_thread = MessageSenderThread(campaign_id, self.bot, log_callback)

        # Connect callbacks
        if log_callback:
            self.sender_thread.update_log.connect(log_callback)

        if progress_callback:
            self.sender_thread.update_progress.connect(progress_callback)

        if stats_callback:
            self.sender_thread.update_stats.connect(stats_callback)

        if finished_callback:
            self.sender_thread.finished.connect(finished_callback)

        # Start thread
        self.sender_thread.start()

        return self.sender_thread

    def stop(self) -> None:
        """Stop the current campaign."""
        if self.sender_thread:
            self.sender_thread.stop()

    def pause(self) -> None:
        """Pause the current campaign."""
        if self.sender_thread:
            self.sender_thread.pause()

    def resume(self) -> None:
        """Resume the current campaign."""
        if self.sender_thread:
            self.sender_thread.resume()

    def cleanup(self) -> None:
        """Cleanup resources."""
        if self.sender_thread and self.sender_thread.isRunning():
            self.sender_thread.stop()
            self.sender_thread.wait()

        if self.bot:
            self.bot.quit()
