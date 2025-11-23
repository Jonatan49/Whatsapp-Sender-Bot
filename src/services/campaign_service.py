"""Campaign management service."""

from typing import List, Optional
from datetime import datetime
from sqlalchemy.orm import Session

from src.models.campaign import Campaign, CampaignStatus
from src.models.message import Message, MessageStatus
from src.models.contact import Contact
from src.core.database import get_db
from src.utils.logger import get_logger


class CampaignService:
    """Service for managing campaigns."""

    def __init__(self):
        """Initialize campaign service."""
        self.db = get_db()
        self.logger = get_logger('campaign_service')

    def create_campaign(
        self,
        name: str,
        message_content: str,
        user_id: int,
        contact_ids: List[int],
        description: Optional[str] = None,
        template_id: Optional[int] = None,
        **settings
    ) -> Campaign:
        """Create a new campaign.

        Args:
            name: Campaign name
            message_content: Message content
            user_id: User ID who creates the campaign
            contact_ids: List of contact IDs
            description: Campaign description
            template_id: Message template ID (if using template)
            **settings: Additional campaign settings

        Returns:
            Created Campaign object
        """
        with self.db.session_scope() as session:
            # Create campaign
            campaign = Campaign(
                name=name,
                description=description,
                message_content=message_content,
                user_id=user_id,
                template_id=template_id,
                total_contacts=len(contact_ids),
                messages_pending=len(contact_ids),
                status=CampaignStatus.DRAFT,
                **settings
            )

            session.add(campaign)
            session.flush()  # Get campaign ID

            # Create messages for each contact
            for contact_id in contact_ids:
                contact = session.query(Contact).filter_by(id=contact_id).first()
                if contact:
                    message = Message(
                        campaign_id=campaign.id,
                        contact_id=contact_id,
                        phone_number=contact.phone_number,
                        content=message_content,
                        status=MessageStatus.PENDING
                    )
                    session.add(message)

            session.commit()

            self.logger.info(f"Campaign '{name}' created with {len(contact_ids)} contacts")
            return campaign

    def get_campaign(self, campaign_id: int) -> Optional[Campaign]:
        """Get campaign by ID.

        Args:
            campaign_id: Campaign ID

        Returns:
            Campaign object or None
        """
        with self.db.session_scope() as session:
            return session.query(Campaign).filter_by(id=campaign_id).first()

    def get_all_campaigns(self, user_id: Optional[int] = None) -> List[Campaign]:
        """Get all campaigns.

        Args:
            user_id: Filter by user ID (optional)

        Returns:
            List of Campaign objects
        """
        with self.db.session_scope() as session:
            query = session.query(Campaign)

            if user_id:
                query = query.filter_by(user_id=user_id)

            return query.order_by(Campaign.created_at.desc()).all()

    def start_campaign(self, campaign_id: int) -> bool:
        """Start a campaign.

        Args:
            campaign_id: Campaign ID

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            campaign = session.query(Campaign).filter_by(id=campaign_id).first()

            if not campaign:
                return False

            if campaign.status != CampaignStatus.DRAFT:
                self.logger.warning(f"Cannot start campaign {campaign_id}: Invalid status {campaign.status}")
                return False

            campaign.status = CampaignStatus.RUNNING
            campaign.started_at = datetime.now().isoformat()

            session.commit()

            self.logger.info(f"Campaign {campaign_id} started")
            return True

    def pause_campaign(self, campaign_id: int) -> bool:
        """Pause a running campaign.

        Args:
            campaign_id: Campaign ID

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            campaign = session.query(Campaign).filter_by(id=campaign_id).first()

            if not campaign:
                return False

            if campaign.status != CampaignStatus.RUNNING:
                return False

            campaign.status = CampaignStatus.PAUSED
            session.commit()

            self.logger.info(f"Campaign {campaign_id} paused")
            return True

    def resume_campaign(self, campaign_id: int) -> bool:
        """Resume a paused campaign.

        Args:
            campaign_id: Campaign ID

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            campaign = session.query(Campaign).filter_by(id=campaign_id).first()

            if not campaign:
                return False

            if campaign.status != CampaignStatus.PAUSED:
                return False

            campaign.status = CampaignStatus.RUNNING
            session.commit()

            self.logger.info(f"Campaign {campaign_id} resumed")
            return True

    def complete_campaign(self, campaign_id: int) -> bool:
        """Mark campaign as completed.

        Args:
            campaign_id: Campaign ID

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            campaign = session.query(Campaign).filter_by(id=campaign_id).first()

            if not campaign:
                return False

            campaign.status = CampaignStatus.COMPLETED
            campaign.completed_at = datetime.now().isoformat()

            # Update statistics
            campaign.update_statistics()

            session.commit()

            self.logger.info(f"Campaign {campaign_id} completed")
            return True

    def get_pending_messages(self, campaign_id: int) -> List[Message]:
        """Get pending messages for a campaign.

        Args:
            campaign_id: Campaign ID

        Returns:
            List of pending Message objects
        """
        with self.db.session_scope() as session:
            return session.query(Message).filter_by(
                campaign_id=campaign_id,
                status=MessageStatus.PENDING
            ).all()

    def update_message_status(
        self,
        message_id: int,
        status: MessageStatus,
        error_message: Optional[str] = None
    ) -> bool:
        """Update message status.

        Args:
            message_id: Message ID
            status: New status
            error_message: Error message if failed

        Returns:
            True if successful, False otherwise
        """
        with self.db.session_scope() as session:
            message = session.query(Message).filter_by(id=message_id).first()

            if not message:
                return False

            message.status = status

            if status == MessageStatus.SENT:
                message.sent_at = datetime.now().isoformat()
            elif status == MessageStatus.FAILED:
                message.error_message = error_message

            session.commit()
            return True

    def get_campaign_statistics(self, campaign_id: int) -> dict:
        """Get campaign statistics.

        Args:
            campaign_id: Campaign ID

        Returns:
            Dictionary with campaign statistics
        """
        with self.db.session_scope() as session:
            campaign = session.query(Campaign).filter_by(id=campaign_id).first()

            if not campaign:
                return {}

            campaign.update_statistics()

            return {
                'id': campaign.id,
                'name': campaign.name,
                'status': campaign.status.value,
                'total_contacts': campaign.total_contacts,
                'messages_sent': campaign.messages_sent,
                'messages_failed': campaign.messages_failed,
                'messages_pending': campaign.messages_pending,
                'success_rate': campaign.success_rate,
                'completion_percentage': campaign.completion_percentage,
                'started_at': campaign.started_at,
                'completed_at': campaign.completed_at,
                'total_duration': campaign.total_duration
            }
