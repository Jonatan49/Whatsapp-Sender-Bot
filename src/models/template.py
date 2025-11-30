"""Message template model for reusable messages."""

from sqlalchemy import Column, String, Integer, ForeignKey, JSON
from sqlalchemy.orm import relationship
from .base import BaseModel
import re


class MessageTemplate(BaseModel):
    """Message template model with variable substitution."""

    __tablename__ = 'message_templates'

    name = Column(String(100), unique=True, nullable=False, index=True)
    content = Column(String(5000), nullable=False)
    description = Column(String(500), nullable=True)
    category = Column(String(50), nullable=True)

    # User who created the template
    user_id = Column(Integer, ForeignKey('users.id'), nullable=False)
    user = relationship('User', backref='templates')

    # Variables (stored as JSON array)
    variables = Column(JSON, nullable=True)  # e.g., ["name", "date", "amount"]

    # Usage statistics
    usage_count = Column(Integer, default=0, nullable=False)

    @classmethod
    def extract_variables(cls, content: str) -> list:
        """Extract variables from template content.

        Variables are in the format: {variable_name}
        Example: "Hello {name}, your appointment is on {date}"
        Returns: ["name", "date"]
        """
        pattern = r'\{([a-zA-Z_][a-zA-Z0-9_]*)\}'
        return list(set(re.findall(pattern, content)))

    def render(self, context: dict) -> str:
        """Render template with provided context.

        Args:
            context: Dictionary with variable values

        Returns:
            Rendered message string

        Example:
            template.content = "Hello {name}, your code is {code}"
            template.render({"name": "John", "code": "1234"})
            Returns: "Hello John, your code is 1234"
        """
        result = self.content

        # Extract all variables from template
        variables = self.extract_variables(self.content)

        # Replace each variable
        for var in variables:
            placeholder = '{' + var + '}'
            value = context.get(var, placeholder)  # Keep placeholder if not provided
            result = result.replace(placeholder, str(value))

        return result

    def validate_context(self, context: dict) -> tuple[bool, list]:
        """Validate that all required variables are provided.

        Args:
            context: Dictionary with variable values

        Returns:
            Tuple of (is_valid, missing_variables)
        """
        variables = self.extract_variables(self.content)
        missing = [var for var in variables if var not in context]
        return (len(missing) == 0, missing)

    def save_variables(self) -> None:
        """Extract and save variables from content."""
        self.variables = self.extract_variables(self.content)

    def __repr__(self) -> str:
        """String representation."""
        return f"<MessageTemplate(id={self.id}, name='{self.name}', variables={self.variables})>"
