"""Database connection and session management."""

from typing import Generator
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, Session, scoped_session
from sqlalchemy.pool import StaticPool
from contextlib import contextmanager
import os

from src.models.base import Base
from src.models import User, Contact, Campaign, Message, MessageTemplate
from src.models.contact import ContactGroup


class Database:
    """Database manager with connection pooling and session management."""

    def __init__(self, database_url: str = None, echo: bool = False):
        """Initialize database connection.

        Args:
            database_url: SQLAlchemy database URL
            echo: Enable SQL query logging
        """
        if database_url is None:
            database_url = os.getenv('DATABASE_URL', 'sqlite:///data/whatsapp_bot.db')

        # Ensure data directory exists
        if database_url.startswith('sqlite:///'):
            db_path = database_url.replace('sqlite:///', '')
            os.makedirs(os.path.dirname(db_path) if os.path.dirname(db_path) else 'data', exist_ok=True)

        # Create engine
        if 'sqlite' in database_url:
            # Use NullPool for SQLite to avoid threading issues
            # Each thread will get its own connection
            from sqlalchemy.pool import NullPool
            self.engine = create_engine(
                database_url,
                echo=echo,
                connect_args={'check_same_thread': False},
                poolclass=NullPool  # Better for threading than StaticPool
            )
        else:
            self.engine = create_engine(database_url, echo=echo, pool_size=5, max_overflow=10)

        # Create session factory
        self.SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=self.engine)
        # Use scoped_session for thread safety - each thread gets its own session
        self.ScopedSession = scoped_session(self.SessionLocal)

    def create_tables(self) -> None:
        """Create all database tables."""
        Base.metadata.create_all(bind=self.engine)

    def drop_tables(self) -> None:
        """Drop all database tables. Use with caution!"""
        Base.metadata.drop_all(bind=self.engine)

    def get_session(self) -> Session:
        """Get a new database session."""
        return self.SessionLocal()

    @contextmanager
    def session_scope(self) -> Generator[Session, None, None]:
        """Provide a transactional scope for database operations.

        Usage:
            with db.session_scope() as session:
                user = session.query(User).first()
        """
        session = self.get_session()
        try:
            yield session
            session.commit()
        except Exception:
            session.rollback()
            raise
        finally:
            session.close()

    def init_default_data(self) -> None:
        """Initialize database with default data."""
        with self.session_scope() as session:
            # Check if admin user exists
            admin = session.query(User).filter_by(username='admin').first()

            if not admin:
                # Create default admin user
                admin = User(
                    username='admin',
                    email='admin@whatsappbot.local',
                    full_name='Administrator',
                    is_admin=True,
                    is_active=True,
                    language='he',
                    theme='whatsapp'
                )
                admin.set_password('admin123')  # Default password - should be changed
                session.add(admin)
                session.flush()  # Flush to get admin.id before using it

                print("✓ Created default admin user (username: admin, password: admin123)")
                print("⚠ Please change the default password immediately!")

            # Create default message templates
            template_count = session.query(MessageTemplate).count()
            if template_count == 0:
                templates = [
                    MessageTemplate(
                        name='Welcome Message',
                        content='שלום {name}, ברוכים הבאים! אנחנו שמחים שהצטרפת אלינו.',
                        description='Default welcome message template',
                        category='general',
                        user_id=admin.id,
                        variables=['name']
                    ),
                    MessageTemplate(
                        name='Reminder',
                        content='היי {name}, רק רציתי להזכיר לך על {event} ב-{date}.',
                        description='Event reminder template',
                        category='reminder',
                        user_id=admin.id,
                        variables=['name', 'event', 'date']
                    ),
                    MessageTemplate(
                        name='Thank You',
                        content='תודה רבה {name}! נשמח לראותך שוב.',
                        description='Thank you message',
                        category='general',
                        user_id=admin.id,
                        variables=['name']
                    )
                ]

                for template in templates:
                    template.save_variables()
                    session.add(template)

                print(f"✓ Created {len(templates)} default message templates")

    def backup_database(self, backup_path: str = None) -> str:
        """Create a backup of the database.

        Args:
            backup_path: Path to save backup file

        Returns:
            Path to backup file
        """
        if backup_path is None:
            import datetime
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_path = f"data/backups/whatsapp_bot_backup_{timestamp}.db"

        os.makedirs(os.path.dirname(backup_path), exist_ok=True)

        # For SQLite, just copy the file
        if 'sqlite' in str(self.engine.url):
            import shutil
            db_path = str(self.engine.url).replace('sqlite:///', '')
            shutil.copy2(db_path, backup_path)
            return backup_path

        # For other databases, would need different approach
        raise NotImplementedError("Backup not implemented for non-SQLite databases")

    def get_statistics(self) -> dict:
        """Get database statistics.

        Returns:
            Dictionary with database statistics
        """
        with self.session_scope() as session:
            return {
                'users': session.query(User).count(),
                'contacts': session.query(Contact).count(),
                'campaigns': session.query(Campaign).count(),
                'messages': session.query(Message).count(),
                'templates': session.query(MessageTemplate).count(),
                'contact_groups': session.query(ContactGroup).count()
            }


# Global database instance
_db_instance = None


def get_db() -> Database:
    """Get or create global database instance."""
    global _db_instance
    if _db_instance is None:
        _db_instance = Database()
    return _db_instance


def init_db() -> Database:
    """Initialize database with tables and default data."""
    db = get_db()
    db.create_tables()
    db.init_default_data()
    return db
