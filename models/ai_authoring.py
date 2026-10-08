"""
SlideGuard AI-Assisted Accessibility Authoring Data Models (Phase 9).
Defines structured models, enums, audit records, and summaries for AI content authoring.
Ensures strict decoupling: AI outputs semantic suggestions; deterministic services apply changes.
"""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Dict, Any, List, Optional


class AuthoringType(str, Enum):
    """Categorized authoring target types for AI assistance."""
    ALT_TEXT = "ALT_TEXT"
    IMAGE_DESCRIPTION = "IMAGE_DESCRIPTION"
    CHART_DESCRIPTION = "CHART_DESCRIPTION"
    DIAGRAM_DESCRIPTION = "DIAGRAM_DESCRIPTION"
    SLIDE_TITLE = "SLIDE_TITLE"
    TABLE_DESCRIPTION = "TABLE_DESCRIPTION"
    TABLE_HEADER_INTERPRETATION = "TABLE_HEADER_INTERPRETATION"
    ACCESSIBILITY_EXPLANATION = "ACCESSIBILITY_EXPLANATION"


class AuthoringStatus(str, Enum):
    """Lifecycle review and application status of an AI authoring recommendation."""
    PENDING_REVIEW = "PENDING_REVIEW"
    APPROVED = "APPROVED"
    EDITED = "EDITED"
    REJECTED = "REJECTED"
    APPLIED = "APPLIED"
    FAILED = "FAILED"
    EXPIRED = "EXPIRED"
    MANUAL_REVIEW = "MANUAL_REVIEW"


class FactGrounding(str, Enum):
    """Fact grounding classification to protect against hallucination."""
    SUPPORTED_FACT = "SUPPORTED_FACT"
    INFERENCE = "INFERENCE"
    UNCERTAIN = "UNCERTAIN"


@dataclass
class AuthoringRecommendation:
    """
    Structured AI authoring recommendation.
    Contains AI suggestion, evidence, confidence, and human review decisions.
    Never directly touches PPTX; strictly acts as an approval payload.
    """
    recommendation_id: str
    rule_id: str
    slide_number: int
    shape_id: Optional[int] = None
    shape_name: str = ""
    authoring_type: AuthoringType = AuthoringType.ALT_TEXT
    current_value: str = ""
    suggested_value: str = ""
    user_edited_value: Optional[str] = None
    rationale: str = ""
    confidence: float = 0.8
    fact_grounding: FactGrounding = FactGrounding.SUPPORTED_FACT
    evidence: Dict[str, Any] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    requires_human_review: bool = True
    status: AuthoringStatus = AuthoringStatus.PENDING_REVIEW
    model: str = "gemini-3.8-flash-high"
    created_at: str = field(default_factory=lambda: datetime.utcnow().isoformat())
    applied_at: Optional[str] = None
    original_value: str = ""

    def __post_init__(self):
        if not self.original_value and self.current_value:
            self.original_value = self.current_value

    @property
    def effective_value(self) -> str:
        """Returns the user-edited value if edited/approved, else the AI suggested value."""
        if self.status == AuthoringStatus.EDITED and self.user_edited_value is not None:
            return self.user_edited_value
        if self.status == AuthoringStatus.APPROVED and self.user_edited_value is not None:
            return self.user_edited_value
        return self.suggested_value

    def approve(self):
        """Marks the recommendation as user-approved."""
        self.status = AuthoringStatus.APPROVED

    def edit(self, edited_text: str):
        """Updates with user-edited text and marks as EDITED."""
        self.user_edited_value = edited_text
        self.status = AuthoringStatus.EDITED

    def reject(self):
        """Marks the recommendation as rejected by the user."""
        self.status = AuthoringStatus.REJECTED

    def mark_decorative(self):
        """Marks the visual shape as decorative with empty alt text and decorative tag."""
        self.user_edited_value = "[MARK_DECORATIVE]"
        self.status = AuthoringStatus.APPROVED
        self.evidence["action"] = "MARK_DECORATIVE"

    def keep_existing(self):
        """Keeps author's existing alt text unmodified."""
        self.user_edited_value = self.original_value or ""
        self.status = AuthoringStatus.APPROVED
        self.evidence["action"] = "PRESERVE"


    def mark_applied(self):
        """Marks the recommendation as successfully applied to OOXML."""
        self.status = AuthoringStatus.APPLIED
        self.applied_at = datetime.utcnow().isoformat()

    def mark_failed(self, error: str):
        """Marks application failure."""
        self.status = AuthoringStatus.FAILED
        self.warnings.append(f"Application failure: {error}")

    def to_dict(self) -> Dict[str, Any]:
        return {
            "recommendation_id": self.recommendation_id,
            "rule_id": self.rule_id,
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "authoring_type": self.authoring_type.value,
            "current_value": self.current_value,
            "suggested_value": self.suggested_value,
            "user_edited_value": self.user_edited_value,
            "effective_value": self.effective_value,
            "rationale": self.rationale,
            "confidence": round(self.confidence, 2),
            "fact_grounding": self.fact_grounding.value,
            "evidence": self.evidence,
            "warnings": self.warnings,
            "requires_human_review": self.requires_human_review,
            "status": self.status.value,
            "model": self.model,
            "created_at": self.created_at,
            "applied_at": self.applied_at,
            "original_value": self.original_value,
        }


@dataclass
class AuthoringAuditRecord:
    """Audit log entry capturing the exact lineage of an AI authoring action."""
    audit_id: str
    recommendation_id: str
    model: str
    authoring_type: str
    slide_number: int
    shape_id: Optional[int]
    original_value: str
    ai_suggestion: str
    user_decision: str
    final_applied_value: str
    confidence: float
    status: str
    timestamp: str = field(default_factory=lambda: datetime.utcnow().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return {
            "audit_id": self.audit_id,
            "recommendation_id": self.recommendation_id,
            "model": self.model,
            "authoring_type": self.authoring_type,
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "original_value": self.original_value,
            "ai_suggestion": self.ai_suggestion,
            "user_decision": self.user_decision,
            "final_applied_value": self.final_applied_value,
            "confidence": round(self.confidence, 2),
            "status": self.status,
            "timestamp": self.timestamp,
        }


@dataclass
class AuthoringSummary:
    """Aggregated metrics for AI-assisted accessibility authoring."""
    generated: int = 0
    approved: int = 0
    edited: int = 0
    rejected: int = 0
    applied: int = 0
    failed: int = 0
    manual_review: int = 0
    fallback_cases: int = 0
    malformed_responses: int = 0
    by_type: Dict[str, int] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "generated": self.generated,
            "approved": self.approved,
            "edited": self.edited,
            "rejected": self.rejected,
            "applied": self.applied,
            "failed": self.failed,
            "manual_review": self.manual_review,
            "fallback_cases": self.fallback_cases,
            "malformed_responses": self.malformed_responses,
            "by_type": self.by_type,
        }
