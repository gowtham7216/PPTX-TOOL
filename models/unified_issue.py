"""
SlideGuard Unified Accessibility Issue Model (Phase 8).
Represents a cross-rule correlated, context-aware unified issue with confidence scoring,
remediation eligibility, and pre/post-remediation reconciliation status.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Optional, List, Dict, Any

from models.enums import IssueSeverity, AccessibilityStatus
from models.issue import Issue


class ConfidenceLevel(str, Enum):
    """Evaluated confidence level for issue classification and remediation safety."""
    HIGH = "HIGH"
    MEDIUM = "MEDIUM"
    LOW = "LOW"


class RemediationEligibility(str, Enum):
    """Determines whether and how an issue may be remediated."""
    ELIGIBLE_AUTO = "ELIGIBLE_AUTO"                  # Satisfies deterministic high-confidence criteria
    REQUIRES_USER_APPROVAL = "REQUIRES_USER_APPROVAL"# Medium-confidence or heuristic action needing confirmation
    MANUAL_ONLY = "MANUAL_ONLY"                      # Cannot safely automate; requires human authoring
    UNSUPPORTED = "UNSUPPORTED"                      # Beyond automated scope (e.g. complex merged tables)


class ReconciliationStatus(str, Enum):
    """Classification of issue lifecycle across remediation."""
    RESOLVED = "RESOLVED"          # Issue detected before, verified gone after remediation
    REMEDIATED = "REMEDIATED"      # Deterministic remediation action applied
    UNCHANGED = "UNCHANGED"        # Still present after remediation (skipped/unremediated)
    NEW = "NEW"                    # Issue not present prior to remediation (side-effect/regression)
    CHANGED = "CHANGED"            # Issue altered in severity or nature post-remediation
    MANUAL_REVIEW = "MANUAL_REVIEW"# Flagged for manual inspection
    UNSUPPORTED = "UNSUPPORTED"    # Out of scope for automated remediation


class ContextCategory(str, Enum):
    """Contextual classification category."""
    TITLE = "TITLE"
    IMAGE = "IMAGE"
    TEXT = "TEXT"
    TABLE = "TABLE"
    READING_ORDER = "READING_ORDER"
    GROUP = "GROUP"
    LIST = "LIST"
    DOCUMENT = "DOCUMENT"
    OTHER = "OTHER"


@dataclass
class UnifiedIssue:
    """
    Cross-rule correlated accessibility issue.
    Unifies overlapping findings into a single coherent decision object
    while fully preserving all original rule IDs and evidence.
    """
    unified_id: str
    slide_number: int
    shape_id: Optional[int] = None
    shape_name: Optional[str] = None
    primary_rule_id: str = ""
    primary_rule_name: str = ""
    related_rule_ids: List[str] = field(default_factory=list)
    severity: IssueSeverity = IssueSeverity.WARNING
    confidence: ConfidenceLevel = ConfidenceLevel.HIGH
    explanation: str = ""
    recommended_action: str = ""
    remediation_eligibility: RemediationEligibility = RemediationEligibility.REQUIRES_USER_APPROVAL
    semantic_role: Optional[str] = None
    context_category: str = ContextCategory.OTHER.value
    raw_issues: List[Issue] = field(default_factory=list)
    status: AccessibilityStatus = AccessibilityStatus.DETECTED
    reconciliation_status: Optional[ReconciliationStatus] = None
    verification_status: Optional[str] = None
    extra_data: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "unified_id": self.unified_id,
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "primary_rule_id": self.primary_rule_id,
            "primary_rule_name": self.primary_rule_name,
            "related_rule_ids": list(self.related_rule_ids),
            "severity": self.severity.value,
            "confidence": self.confidence.value,
            "explanation": self.explanation,
            "recommended_action": self.recommended_action,
            "remediation_eligibility": self.remediation_eligibility.value,
            "semantic_role": self.semantic_role,
            "context_category": self.context_category,
            "status": self.status.value if isinstance(self.status, AccessibilityStatus) else str(self.status),
            "reconciliation_status": self.reconciliation_status.value if self.reconciliation_status else None,
            "verification_status": self.verification_status,
            "raw_issues_count": len(self.raw_issues),
            "raw_issues": [issue.to_dict() for issue in self.raw_issues],
            "extra_data": self.extra_data,
        }
