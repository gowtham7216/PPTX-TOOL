"""
SlideGuard Issue Data Model.
Captures individual accessibility violations or observations.
"""

from dataclasses import dataclass
from typing import Optional, Dict, Any
from models.enums import IssueSeverity, RuleBehavior, AccessibilityStatus


@dataclass
class Issue:
    """An identified accessibility issue requiring remediation or audit review."""
    issue_id: str
    rule_id: str
    rule_name: str
    slide_number: int
    severity: IssueSeverity
    description: str
    recommendation: str
    behavior: RuleBehavior
    shape_id: Optional[int] = None
    shape_name: Optional[str] = None
    wcag_mapping: Optional[str] = None
    extra_data: Optional[Dict[str, Any]] = None
    status: AccessibilityStatus = AccessibilityStatus.DETECTED
    confidence: str = "HIGH"
    explanation: Optional[str] = None
    verification_status: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "issue_id": self.issue_id,
            "rule_id": self.rule_id,
            "rule_name": self.rule_name,
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "severity": self.severity.value,
            "wcag_mapping": self.wcag_mapping,
            "description": self.description,
            "recommendation": self.recommendation,
            "behavior": self.behavior.value,
            "status": self.status.value if isinstance(self.status, AccessibilityStatus) else str(self.status),
            "confidence": self.confidence,
            "explanation": self.explanation,
            "verification_status": self.verification_status,
            "extra_data": self.extra_data or {},
        }
