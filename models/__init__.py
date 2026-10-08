"""
SlideGuard Data Models Package.
"""

from models.enums import (
    RuleBehavior,
    AccessibilityStatus,
    IssueSeverity,
    RuleStatus,
    RemediationStatus,
    ShapeType,
    PlaceholderType,
    ListType,
    JobStage,
)
from models.shape import ShapeModel, TextRunModel, TextParagraphModel
from models.slide import SlideModel
from models.presentation import PresentationModel, DocumentMetadata
from models.issue import Issue
from models.rule_result import RuleResult
from models.remediation import RemediationAction
from models.verification import VerificationResult, VerificationCheck
from models.comparison import ComparisonResult, ComparisonDifference
from models.job import JobContext, QualityMetrics
from models.unified_issue import (
    UnifiedIssue,
    ConfidenceLevel,
    RemediationEligibility,
    ReconciliationStatus,
    ContextCategory,
)
from models.ai_authoring import (
    AuthoringType,
    AuthoringStatus,
    FactGrounding,
    AuthoringRecommendation,
    AuthoringAuditRecord,
    AuthoringSummary,
)
from models.qa import (
    QAStatus,
    QACheckItem,
    QAResult,
)

__all__ = [
    "RuleBehavior",
    "AccessibilityStatus",
    "IssueSeverity",
    "RuleStatus",
    "RemediationStatus",
    "ShapeType",
    "PlaceholderType",
    "ListType",
    "JobStage",
    "ShapeModel",
    "TextRunModel",
    "TextParagraphModel",
    "SlideModel",
    "PresentationModel",
    "DocumentMetadata",
    "Issue",
    "RuleResult",
    "RemediationAction",
    "VerificationResult",
    "VerificationCheck",
    "ComparisonResult",
    "ComparisonDifference",
    "JobContext",
    "QualityMetrics",
    "UnifiedIssue",
    "ConfidenceLevel",
    "RemediationEligibility",
    "ReconciliationStatus",
    "ContextCategory",
    "AuthoringType",
    "AuthoringStatus",
    "FactGrounding",
    "AuthoringRecommendation",
    "AuthoringAuditRecord",
    "AuthoringSummary",
    "QAStatus",
    "QACheckItem",
    "QAResult",
]

