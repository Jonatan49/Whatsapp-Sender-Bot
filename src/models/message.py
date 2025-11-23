"""Message model for tracking individual messages."""

from sqlalchemy import Column, String, Integer, ForeignKey, Float, Enum as SQLEnum, Text
from sqlalchemy.orm import relationship
from .base import BaseModel
import enum


class MessageStatus(enum.Enum):
    """Message status enumeration."""
    PENDING = "pending"
    SENDING = "sending"
    SENT = "sent"
    FAILED = "failed"
    SKIPPED = "skipped"


class Message(BaseModel):
    """Message model for tracking individual message sends."""

    __tablename__ = 'messages'

    # Relationships
    campaign_id = Column(Integer, ForeignKey('campaigns.id'), nullable=False, index=True)
    campaign = relationship('Campaign', back_populates='messages')

    contact_id = Column(Integer, ForeignKey('contacts.id'), nullable=False, index=True)
    contact = relationship('Contact', back_populates='messages')

    # Message details
    content = Column(Text, nullable=False)
    status = Column(SQLEnum(MessageStatus), default=MessageStatus.PENDING, nullable=False, index=True)

    # Timing
    sent_at = Column(String(50), nullable=True)
    delivery_time = Column(Float, nullable=True)  # Time taken to send in seconds

    # Error handling
    error_message = Column(String(500), nullable=True)
    retry_count = Column(Integer, default=0, nullable=False)

    # Metadata
    phone_number = Column(String(20), nullable=False)  # Denormalized for performance
    link_preview = Column(Boolean, default=False, nullable=False)

    def __repr__(self) -> str:
        """String representation."""
        return f"<Message(id={self.id}, campaign_id={self.campaign_id}, contact={self.phone_number}, status={self.status.value})>"


from sqlalchemy import Boolean
