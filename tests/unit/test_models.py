"""
Unit tests for SlideGuard internal data models and enums.
"""

from pathlib import Path
from models.enums import (
    RuleBehavior,
    IssueSeverity,
    RuleStatus,
    RemediationStatus,
    ShapeType,
    PlaceholderType,
)
from models.shape import ShapeModel, TextRunModel, TextParagraphModel
from models.slide import SlideModel
from models.presentation import PresentationModel, DocumentMetadata
from models.issue import Issue
from models.rule_result import RuleResult
from models.remediation import RemediationAction
from models.verification import VerificationResult
from models.comparison import ComparisonResult, ComparisonDifference
from models.job import JobContext, QualityMetrics


def test_enums():
    assert RuleBehavior.DETECT_ONLY.value == "DETECT_ONLY"
    assert RuleBehavior.SAFE_AUTO_REMEDIATE.value == "SAFE_AUTO_REMEDIATE"
    assert RuleBehavior.AI_ASSISTED_REVIEW.value == "AI_ASSISTED_REVIEW"

    assert IssueSeverity.CRITICAL.value == "CRITICAL"
    assert RuleStatus.PASSED.value == "PASSED"
    assert RemediationStatus.NOT_IMPLEMENTED.value == "NOT_IMPLEMENTED"


def test_shape_model():
    shape = ShapeModel(
        shape_id=101,
        name="Title 1",
        shape_type=ShapeType.PLACEHOLDER,
        left_pt=72.0,
        top_pt=36.0,
        width_pt=500.0,
        height_pt=60.0,
        placeholder_type=PlaceholderType.TITLE,
        text_content="Accessible Slide Title"
    )
    assert shape.right_pt == 572.0
    assert shape.bottom_pt == 96.0
    assert shape.has_text is True


def test_slide_model():
    shape = ShapeModel(
        shape_id=1,
        name="Body Box",
        shape_type=ShapeType.TEXT_BOX,
        left_pt=10.0,
        top_pt=10.0,
        width_pt=100.0,
        height_pt=100.0
    )
    slide = SlideModel(
        slide_id=256,
        slide_number=1,
        title="Welcome",
        shapes=[shape]
    )
    assert slide.shape_count == 1
    found = slide.get_shape_by_id(1)
    assert found is not None
    assert found.name == "Body Box"
    assert slide.get_shape_by_id(999) is None


def test_issue_serialization():
    issue = Issue(
        issue_id="ISSUE_001",
        rule_id="RULE_01",
        rule_name="Test Rule",
        slide_number=2,
        severity=IssueSeverity.WARNING,
        description="Missing alt text",
        recommendation="Add descriptive text",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW
    )
    d = issue.to_dict()
    assert d["issue_id"] == "ISSUE_001"
    assert d["severity"] == "WARNING"
    assert d["behavior"] == "AI_ASSISTED_REVIEW"


def test_quality_metrics_calculation():
    metrics = QualityMetrics(
        total_rules_evaluated=12,
        passed_rules=10,
        failed_rules=2,
        total_issues=2,
        warning_issues=2,
        accessibility_score=96.0
    )
    d = metrics.to_dict()
    assert d["accessibility_score"] == 96.0
    assert d["passed_rules"] == 10
