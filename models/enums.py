"""
SlideGuard Core Enumerations.
Defines standard status, severity, behavior, and structural types.
"""

from enum import Enum


class RuleBehavior(str, Enum):
    """Execution behavior classification for accessibility rules."""
    DETECT_ONLY = "DETECT_ONLY"
    SAFE_AUTO_REMEDIATE = "SAFE_AUTO_REMEDIATE"
    AI_ASSISTED_REVIEW = "AI_ASSISTED_REVIEW"
    MANUAL_REVIEW = "MANUAL_REVIEW"


class AccessibilityStatus(str, Enum):
    """Status classification for accessibility issues and remediation."""
    DETECTED = "DETECTED"
    SAFE_TO_REMEDIATE = "SAFE_TO_REMEDIATE"
    REMEDIATED = "REMEDIATED"
    PASSED = "PASSED"
    MANUAL_REVIEW = "MANUAL_REVIEW"
    UNSUPPORTED = "UNSUPPORTED"


class IssueSeverity(str, Enum):
    """Severity levels for accessibility issues."""
    CRITICAL = "CRITICAL"
    ERROR = "ERROR"
    WARNING = "WARNING"
    INFO = "INFO"


class RuleStatus(str, Enum):
    """Execution status for a rule on a specific slide or document."""
    PASSED = "PASSED"
    FAILED = "FAILED"
    WARNING = "WARNING"
    NOT_APPLICABLE = "NOT_APPLICABLE"
    SKIPPED = "SKIPPED"
    ERROR = "ERROR"


class RemediationStatus(str, Enum):
    """Status of an attempted remediation action."""
    NOT_STARTED = "NOT_STARTED"
    APPLIED = "APPLIED"
    NO_CHANGES = "NO_CHANGES"
    SKIPPED = "SKIPPED"
    FAILED = "FAILED"
    NOT_IMPLEMENTED = "NOT_IMPLEMENTED"


class ShapeType(str, Enum):
    """Categorized shape type for PowerPoint objects."""
    TEXT_BOX = "TEXT_BOX"
    PLACEHOLDER = "PLACEHOLDER"
    PICTURE = "PICTURE"
    GROUP = "GROUP"
    TABLE = "TABLE"
    AUTO_SHAPE = "AUTO_SHAPE"
    MEDIA = "MEDIA"
    OTHER = "OTHER"


class PlaceholderType(str, Enum):
    """Identified placeholder type."""
    TITLE = "TITLE"
    BODY = "BODY"
    SUBTITLE = "SUBTITLE"
    FOOTER = "FOOTER"
    HEADER = "HEADER"
    DATE_TIME = "DATE_TIME"
    SLIDE_NUMBER = "SLIDE_NUMBER"
    PICTURE = "PICTURE"
    TABLE = "TABLE"
    CHART = "CHART"
    OBJECT = "OBJECT"
    OTHER = "OTHER"
    NONE = "NONE"


class ListType(str, Enum):
    """Semantic list style identified in text frames."""
    NONE = "NONE"
    BULLET = "BULLET"
    NUMBERED_NUMERIC = "NUMBERED_NUMERIC"
    NUMBERED_ALPHA_LOWER = "NUMBERED_ALPHA_LOWER"
    NUMBERED_ALPHA_UPPER = "NUMBERED_ALPHA_UPPER"
    NUMBERED_ROMAN = "NUMBERED_ROMAN"
    OTHER = "OTHER"


class JobStage(str, Enum):
    """Lifecycle stage of a SlideGuard processing job."""
    INITIALIZED = "INITIALIZED"
    ANALYZING = "ANALYZING"
    ANALYZED = "ANALYZED"
    REMEDIATING = "REMEDIATING"
    REMEDIATED = "REMEDIATED"
    VERIFYING = "VERIFYING"
    VERIFIED = "VERIFIED"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"
