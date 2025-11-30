"""User model for authentication."""

from datetime import datetime
from sqlalchemy import Column, String, Boolean, Integer, DateTime
from .base import BaseModel
import bcrypt


class User(BaseModel):
    """User model with secure password handling."""

    __tablename__ = 'users'

    username = Column(String(50), unique=True, nullable=False, index=True)
    password_hash = Column(String(255), nullable=False)
    email = Column(String(100), unique=True, nullable=True)
    full_name = Column(String(100), nullable=True)
    is_active = Column(Boolean, default=True, nullable=False)
    is_admin = Column(Boolean, default=False, nullable=False)
    last_login = Column(DateTime, nullable=True)
    failed_login_attempts = Column(Integer, default=0, nullable=False)
    locked_until = Column(DateTime, nullable=True)
    language = Column(String(5), default='he', nullable=False)
    theme = Column(String(20), default='whatsapp', nullable=False)

    def set_password(self, password: str) -> None:
        """Hash and set the user's password."""
        salt = bcrypt.gensalt(rounds=12)
        self.password_hash = bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')

    def check_password(self, password: str) -> bool:
        """Verify the provided password against the hash."""
        return bcrypt.checkpw(
            password.encode('utf-8'),
            self.password_hash.encode('utf-8')
        )

    def is_locked(self) -> bool:
        """Check if the account is locked."""
        if self.locked_until is None:
            return False
        return datetime.utcnow() < self.locked_until

    def record_login_attempt(self, success: bool) -> None:
        """Record a login attempt."""
        if success:
            self.failed_login_attempts = 0
            self.locked_until = None
            self.last_login = datetime.utcnow()
        else:
            self.failed_login_attempts += 1
            if self.failed_login_attempts >= 5:
                from datetime import timedelta
                self.locked_until = datetime.utcnow() + timedelta(seconds=300)

    def __repr__(self) -> str:
        """String representation."""
        return f"<User(id={self.id}, username='{self.username}', is_active={self.is_active})>"
