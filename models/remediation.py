"""
SlideGuard Remediation Action Data Model.
Captures specific modifications proposed or executed during remediation.
"""

from dataclasses import dataclass, field
from typing import Optional, Dict, Any
from models.enums import RemediationStatus


@dataclass
class RemediationAction:
    """Detailed record of a single remediation step."""
    action_id: str
    rule_id: str
    slide_number: int
    description: str
    status: RemediationStatus = RemediationStatus.NOT_STARTED
    target_shape_id: Optional[int] = None
    original_state: Dict[str, Any] = field(default_factory=dict)
    remediated_state: Dict[str, Any] = field(default_factory=dict)
    confidence: float = 1.0
    error_message: Optional[str] = None
    is_user_approved: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "action_id": self.action_id,
            "rule_id": self.rule_id,
            "slide_number": self.slide_number,
            "target_shape_id": self.target_shape_id,
            "description": self.description,
            "status": self.status.value,
            "original_state": self.original_state,
            "remediated_state": self.remediated_state,
            "confidence": self.confidence,
            "error_message": self.error_message,
            "is_user_approved": self.is_user_approved,
        }
