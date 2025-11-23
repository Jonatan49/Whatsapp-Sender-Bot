"""Campaign model for tracking message campaigns."""

from sqlalchemy import Column, String, Integer, ForeignKey, Float, JSON, Enum as SQLEnum, Boolean
from sqlalchemy.orm import relationship
from .base import BaseModel
import enum


class CampaignStatus(enum.Enum):
    """Campaign status enumeration."""
    DRAFT = "draft"
    SCHEDULED = "scheduled"
    RUNNING = "running"
    PAUSED = "paused"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class Campaign(BaseModel):
    """Campaign model for managing message sending campaigns."""

    __tablename__ = 'campaigns'

    name = Column(String(100), nullable=False, index=True)
    description = Column(String(500), nullable=True)
    status = Column(SQLEnum(CampaignStatus), default=CampaignStatus.DRAFT, nullable=False)

    # User who created the campaign
    user_id = Column(Integer, ForeignKey('users.id'), nullable=False)
    user = relationship('User', backref='campaigns')

    # Message content
    message_content = Column(String(5000), nullable=False)
    template_id = Column(Integer, ForeignKey('message_templates.id'), nullable=True)
    template = relationship('MessageTemplate')

    # Settings
    min_delay = Column(Float, default=5.0, nullable=False)
    max_delay = Column(Float, default=10.0, nullable=False)
    pause_after = Column(Integer, default=100, nullable=False)
    pause_duration = Column(Integer, default=120, nullable=False)
    enable_link_preview = Column(Boolean, default=False, nullable=False)

    # Statistics
    total_contacts = Column(Integer, default=0, nullable=False)
    messages_sent = Column(Integer, default=0, nullable=False)
    messages_failed = Column(Integer, default=0, nullable=False)
    messages_pending = Column(Integer, default=0, nullable=False)

    # Timing
    scheduled_at = Column(String(50), nullable=True)
    started_at = Column(String(50), nullable=True)
    completed_at = Column(String(50), nullable=True)
    total_duration = Column(Float, default=0.0, nullable=False)  # in seconds

    # Metadata
    folder_path = Column(String(500), nullable=True)
    config = Column(JSON, nullable=True)  # Additional configuration as JSON

    # Relationships
    messages = relationship('Message', back_populates='campaign', cascade='all, delete-orphan')

    @property
    def success_rate(self) -> float:
        """Calculate success rate percentage."""
        if self.messages_sent == 0:
            return 0.0
        successful = self.messages_sent - self.messages_failed
        return (successful / self.messages_sent) * 100

    @property
    def completion_percentage(self) -> float:
        """Calculate completion percentage."""
        if self.total_contacts == 0:
            return 0.0
        return (self.messages_sent / self.total_contacts) * 100

    def update_statistics(self) -> None:
        """Update campaign statistics from messages."""
        from sqlalchemy import func
        from .message import Message, MessageStatus

        # This would typically be called with a session
        # For now, we'll calculate from the messages relationship
        self.messages_sent = sum(1 for m in self.messages if m.status in [MessageStatus.SENT, MessageStatus.FAILED])
        self.messages_failed = sum(1 for m in self.messages if m.status == MessageStatus.FAILED)
        self.messages_pending = sum(1 for m in self.messages if m.status == MessageStatus.PENDING)

    def __repr__(self) -> str:
        """String representation."""
        return f"<Campaign(id={self.id}, name='{self.name}', status={self.status.value}, sent={self.messages_sent}/{self.total_contacts})>"


from sqlalchemy import Boolean
