"""Contact model for managing phone numbers."""

from sqlalchemy import Column, String, Boolean, Integer, ForeignKey, Table
from sqlalchemy.orm import relationship
from .base import BaseModel


# Association table for contact groups
contact_group = Table(
    'contact_group',
    BaseModel.metadata,
    Column('contact_id', Integer, ForeignKey('contacts.id'), primary_key=True),
    Column('group_id', Integer, ForeignKey('contact_groups.id'), primary_key=True)
)


class Contact(BaseModel):
    """Contact model for storing phone numbers and metadata."""

    __tablename__ = 'contacts'

    phone_number = Column(String(20), unique=True, nullable=False, index=True)
    name = Column(String(100), nullable=True)
    country_code = Column(String(5), default='+972', nullable=False)
    is_valid = Column(Boolean, default=True, nullable=False)
    is_blocked = Column(Boolean, default=False, nullable=False)
    notes = Column(String(500), nullable=True)
    source = Column(String(50), nullable=True)  # manual, file, group, etc.

    # Statistics
    messages_sent = Column(Integer, default=0, nullable=False)
    messages_failed = Column(Integer, default=0, nullable=False)
    last_contacted = Column(String(50), nullable=True)

    # Relationships
    groups = relationship('ContactGroup', secondary=contact_group, back_populates='contacts')
    messages = relationship('Message', back_populates='contact', cascade='all, delete-orphan')

    @classmethod
    def normalize_phone(cls, phone: str, country_code: str = '+972') -> str:
        """Normalize phone number to international format."""
        phone = phone.strip().replace(' ', '').replace('-', '').replace('(', '').replace(')', '')

        if phone.startswith('+'):
            return phone
        elif phone.startswith('972'):
            return '+' + phone
        elif phone.startswith('0'):
            return country_code + phone[1:]
        else:
            return country_code + phone

    @classmethod
    def validate_israeli_phone(cls, phone: str) -> bool:
        """Validate Israeli phone number format."""
        if not phone.startswith('+972'):
            return False

        rest = phone[4:]

        # Must start with 5 (mobile) or 05
        if not (rest.startswith('5') or rest.startswith('05')):
            return False

        # Check length (13-14 characters total)
        if len(phone) not in [13, 14]:
            return False

        return True

    def __repr__(self) -> str:
        """String representation."""
        name = f", name='{self.name}'" if self.name else ""
        return f"<Contact(id={self.id}, phone='{self.phone_number}'{name})>"


class ContactGroup(BaseModel):
    """Contact group for organizing contacts."""

    __tablename__ = 'contact_groups'

    name = Column(String(100), unique=True, nullable=False, index=True)
    description = Column(String(500), nullable=True)
    color = Column(String(7), default='#25d366', nullable=False)

    # Relationships
    contacts = relationship('Contact', secondary=contact_group, back_populates='groups')

    def __repr__(self) -> str:
        """String representation."""
        # Don't access lazy-loaded contacts to avoid DetachedInstanceError
        try:
            contact_count = len(self.contacts) if hasattr(self, '_sa_instance_state') and self._sa_instance_state.session else '?'
        except:
            contact_count = '?'
        return f"<ContactGroup(id={self.id}, name='{self.name}', contacts={contact_count})>"
