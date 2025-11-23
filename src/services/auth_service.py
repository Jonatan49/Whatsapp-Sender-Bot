"""Authentication service for user management."""

from typing import Optional
from datetime import datetime, timedelta
from sqlalchemy.orm import Session

from src.models.user import User
from src.core.database import get_db
from src.utils.logger import get_logger


class AuthService:
    """Authentication service with secure password handling."""

    def __init__(self):
        """Initialize authentication service."""
        self.db = get_db()
        self.logger = get_logger('auth_service')

    def authenticate(self, username: str, password: str) -> Optional[User]:
        """Authenticate user with username and password.

        Args:
            username: Username
            password: Plain text password

        Returns:
            User object if authentication successful, None otherwise
        """
        with self.db.session_scope() as session:
            user = session.query(User).filter_by(username=username).first()

            if not user:
                self.logger.warning(f"Authentication failed: User '{username}' not found")
                return None

            if not user.is_active:
                self.logger.warning(f"Authentication failed: User '{username}' is inactive")
                return None

            if user.is_locked():
                self.logger.warning(f"Authentication failed: User '{username}' is locked")
                return None

            # Check password
            if user.check_password(password):
                user.record_login_attempt(success=True)
                session.commit()
                self.logger.info(f"User '{username}' authenticated successfully")
                return user
            else:
                user.record_login_attempt(success=False)
                session.commit()
                self.logger.warning(f"Authentication failed: Invalid password for user '{username}'")
                return None

    def create_user(
        self,
        username: str,
        password: str,
        email: Optional[str] = None,
        full_name: Optional[str] = None,
        is_admin: bool = False
    ) -> Optional[User]:
        """Create a new user.

        Args:
            username: Username
            password: Plain text password
            email: Email address
            full_name: Full name
            is_admin: Admin flag

        Returns:
            Created User object or None if username exists
        """
        with self.db.session_scope() as session:
            # Check if username exists
            existing = session.query(User).filter_by(username=username).first()
            if existing:
                self.logger.warning(f"User creation failed: Username '{username}' already exists")
                return None

            # Create user
            user = User(
                username=username,
                email=email,
                full_name=full_name,
                is_admin=is_admin,
                is_active=True
            )
            user.set_password(password)

            session.add(user)
            session.commit()

            self.logger.info(f"User '{username}' created successfully")
            return user

    def change_password(self, user_id: int, old_password: str, new_password: str) -> bool:
        """Change user password.

        Args:
            user_id: User ID
            old_password: Current password
            new_password: New password

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            user = session.query(User).filter_by(id=user_id).first()

            if not user:
                return False

            if not user.check_password(old_password):
                self.logger.warning(f"Password change failed for user {user_id}: Invalid old password")
                return False

            user.set_password(new_password)
            session.commit()

            self.logger.info(f"Password changed successfully for user {user_id}")
            return True

    def get_user_by_id(self, user_id: int) -> Optional[User]:
        """Get user by ID.

        Args:
            user_id: User ID

        Returns:
            User object or None
        """
        with self.db.session_scope() as session:
            return session.query(User).filter_by(id=user_id).first()

    def get_user_by_username(self, username: str) -> Optional[User]:
        """Get user by username.

        Args:
            username: Username

        Returns:
            User object or None
        """
        with self.db.session_scope() as session:
            return session.query(User).filter_by(username=username).first()
