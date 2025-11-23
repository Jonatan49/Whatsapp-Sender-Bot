"""Contact management service."""

from typing import List, Optional
from sqlalchemy.orm import Session
import pandas as pd
import json

from src.models.contact import Contact, ContactGroup
from src.core.database import get_db
from src.utils.logger import get_logger


class ContactService:
    """Service for managing contacts."""

    def __init__(self):
        """Initialize contact service."""
        self.db = get_db()
        self.logger = get_logger('contact_service')

    def add_contact(
        self,
        phone: str,
        name: Optional[str] = None,
        country_code: str = '+972',
        source: str = 'manual',
        **kwargs
    ) -> Optional[Contact]:
        """Add a new contact.

        Args:
            phone: Phone number
            name: Contact name
            country_code: Country code
            source: Source of contact (manual, file, group, etc.)
            **kwargs: Additional contact fields

        Returns:
            Created Contact object or None if exists
        """
        with self.db.session_scope() as session:
            # Normalize phone number
            normalized_phone = Contact.normalize_phone(phone, country_code)

            # Check if contact exists
            existing = session.query(Contact).filter_by(phone_number=normalized_phone).first()
            if existing:
                self.logger.debug(f"Contact {normalized_phone} already exists")
                return existing

            # Validate phone number
            is_valid = Contact.validate_israeli_phone(normalized_phone)

            # Create contact
            contact = Contact(
                phone_number=normalized_phone,
                name=name,
                country_code=country_code,
                is_valid=is_valid,
                source=source,
                **kwargs
            )

            session.add(contact)
            session.commit()

            self.logger.info(f"Contact added: {normalized_phone}")
            return contact

    def import_contacts_from_file(self, file_path: str, source: str = 'file') -> tuple[List[Contact], List[str]]:
        """Import contacts from file (Excel, CSV, JSON).

        Args:
            file_path: Path to file
            source: Source label for contacts

        Returns:
            Tuple of (successfully imported contacts, error messages)
        """
        contacts = []
        errors = []

        try:
            # Determine file type and read
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, engine='openpyxl')
                phone_numbers = df.iloc[:, 0].dropna().astype(str).tolist()

                # Try to get names from second column if exists
                names = df.iloc[:, 1].dropna().astype(str).tolist() if len(df.columns) > 1 else []

            elif file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
                phone_numbers = df.iloc[:, 0].dropna().astype(str).tolist()

                # Try to get names from second column if exists
                names = df.iloc[:, 1].dropna().astype(str).tolist() if len(df.columns) > 1 else []

            elif file_path.endswith('.json'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                if isinstance(data, list):
                    if isinstance(data[0], dict):
                        phone_numbers = [item.get('phone', item.get('number', '')) for item in data]
                        names = [item.get('name', '') for item in data]
                    else:
                        phone_numbers = data
                        names = []
                else:
                    errors.append("Invalid JSON format")
                    return contacts, errors

            else:
                errors.append(f"Unsupported file format: {file_path}")
                return contacts, errors

            # Import contacts
            for i, phone in enumerate(phone_numbers):
                try:
                    name = names[i] if i < len(names) else None
                    contact = self.add_contact(phone, name=name, source=source)
                    if contact:
                        contacts.append(contact)
                except Exception as e:
                    errors.append(f"Error importing {phone}: {str(e)}")

            self.logger.info(f"Imported {len(contacts)} contacts from {file_path}")

        except Exception as e:
            errors.append(f"Error reading file: {str(e)}")

        return contacts, errors

    def get_contact(self, contact_id: int) -> Optional[Contact]:
        """Get contact by ID.

        Args:
            contact_id: Contact ID

        Returns:
            Contact object or None
        """
        with self.db.session_scope() as session:
            return session.query(Contact).filter_by(id=contact_id).first()

    def get_contact_by_phone(self, phone: str) -> Optional[Contact]:
        """Get contact by phone number.

        Args:
            phone: Phone number

        Returns:
            Contact object or None
        """
        with self.db.session_scope() as session:
            return session.query(Contact).filter_by(phone_number=phone).first()

    def get_all_contacts(self, valid_only: bool = True) -> List[Contact]:
        """Get all contacts.

        Args:
            valid_only: Only return valid contacts

        Returns:
            List of Contact objects
        """
        with self.db.session_scope() as session:
            query = session.query(Contact)

            if valid_only:
                query = query.filter_by(is_valid=True, is_blocked=False)

            return query.all()

    def update_contact(self, contact_id: int, **kwargs) -> Optional[Contact]:
        """Update contact fields.

        Args:
            contact_id: Contact ID
            **kwargs: Fields to update

        Returns:
            Updated Contact object or None
        """
        with self.db.session_scope() as session:
            contact = session.query(Contact).filter_by(id=contact_id).first()

            if not contact:
                return None

            for key, value in kwargs.items():
                if hasattr(contact, key):
                    setattr(contact, key, value)

            session.commit()
            self.logger.info(f"Contact {contact_id} updated")
            return contact

    def delete_contact(self, contact_id: int) -> bool:
        """Delete contact.

        Args:
            contact_id: Contact ID

        Returns:
            True if deleted, False otherwise
        """
        with self.db.session_scope() as session:
            contact = session.query(Contact).filter_by(id=contact_id).first()

            if not contact:
                return False

            session.delete(contact)
            session.commit()

            self.logger.info(f"Contact {contact_id} deleted")
            return True

    def create_group(self, name: str, description: Optional[str] = None, color: str = '#25d366') -> ContactGroup:
        """Create a contact group.

        Args:
            name: Group name
            description: Group description
            color: Group color (hex)

        Returns:
            Created ContactGroup object
        """
        with self.db.session_scope() as session:
            group = ContactGroup(name=name, description=description, color=color)
            session.add(group)
            session.commit()

            self.logger.info(f"Contact group '{name}' created")
            return group

    def add_contact_to_group(self, contact_id: int, group_id: int) -> bool:
        """Add contact to group.

        Args:
            contact_id: Contact ID
            group_id: Group ID

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            contact = session.query(Contact).filter_by(id=contact_id).first()
            group = session.query(ContactGroup).filter_by(id=group_id).first()

            if not contact or not group:
                return False

            if contact not in group.contacts:
                group.contacts.append(contact)
                session.commit()
                self.logger.info(f"Contact {contact_id} added to group {group_id}")

            return True
