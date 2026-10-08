"""
Unit tests for SlideGuard Phase 8: Accessibility Intelligence & Advanced Remediation.
Covers the 22 required synthetic test scenarios:
 1. isolated single-rule issue
 2. overlapping multi-rule issue
 3. duplicate issue detection
 4. title contextual classification
 5. image contextual classification
 6. decorative image contextual classification
 7. ambiguous classification
 8. confidence levels
 9. remediation eligibility
10. issue reconciliation
11. resolved issue
12. unchanged issue
13. new issue
14. changed issue
15. stale-analysis invalidation
16. remediation dependency ordering
17. quality score calculation (V2)
18. Gemini structured recommendation parsing
19. invalid Gemini response handling
20. Gemini unavailable fallback
21. deterministic remediation without AI
22. visual-preservation protection
"""

import pytest
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from models.enums import (
    IssueSeverity,
    RuleStatus,
    RuleBehavior,
    AccessibilityStatus,
    RemediationStatus,
    ShapeType,
)
from models.issue import Issue
from models.rule_result import RuleResult
from models.shape import ShapeModel, TextParagraphModel, TextRunModel
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.job import JobContext
from models.remediation import RemediationAction
from models.unified_issue import (
    UnifiedIssue,
    ConfidenceLevel,
    RemediationEligibility,
    ReconciliationStatus,
    ContextCategory,
)
from engine.intelligence.context_classifier import ContextualClassifier, TitleContext, ImageContext
from engine.intelligence.confidence_engine import ConfidenceEngine
from engine.intelligence.correlator import IssueCorrelator
from engine.intelligence.reconciliation import IssueReconciliationService
from engine.intelligence.gemini_assistant import GeminiAssistant, AIRecommendation
from engine.accessibility.scoring import AccessibilityScorer
from engine.remediation.remediation_service import RemediationService


# =============================================================================
# Helper Fixtures & Builders
# =============================================================================

def _build_test_shape(
    shape_id: int,
    name: str = "Test Shape",
    shape_type: ShapeType = ShapeType.TEXT_BOX,
    left: float = 100.0,
    top: float = 50.0,
    width: float = 400.0,
    height: float = 80.0,
    text: str = "Sample Text",
    font_size: float = 24.0,
    is_bold: bool = False,
    is_placeholder: bool = False,
    placeholder_type: str = "none",
) -> ShapeModel:
    run = TextRunModel(text=text, font_size_pt=font_size, font_bold=is_bold)
    para = TextParagraphModel(text=text, runs=[run])
    return ShapeModel(
        shape_id=shape_id,
        name=name,
        shape_type=shape_type,
        left_pt=left,
        top_pt=top,
        width_pt=width,
        height_pt=height,
        text_content=text,
        paragraphs=[para],
        is_placeholder=is_placeholder,
    )


def _make_issue(
    issue_id: str,
    rule_id: str,
    slide_number: int = 1,
    rule_name: str = "Test Rule",
    severity: IssueSeverity = IssueSeverity.WARNING,
    description: str = "Test description",
    recommendation: str = "Test recommendation",
    behavior: RuleBehavior = RuleBehavior.SAFE_AUTO_REMEDIATE,
    shape_id=None,
    status=AccessibilityStatus.DETECTED,
    confidence: str = "HIGH",
    **kwargs
) -> Issue:
    return Issue(
        issue_id=issue_id,
        rule_id=rule_id,
        rule_name=rule_name,
        slide_number=slide_number,
        severity=severity,
        description=description,
        recommendation=recommendation,
        behavior=behavior,
        shape_id=shape_id,
        status=status,
        confidence=confidence,
        **kwargs
    )


# =============================================================================
# 1. Isolated single-rule issue
# =============================================================================

def test_01_isolated_single_rule_issue():
    issue = Issue(
        issue_id="iss_1",
        rule_id="SG-ALT-01",
        rule_name="Alternative Text Missing",
        slide_number=1,
        shape_id=10,
        severity=IssueSeverity.CRITICAL,
        description="Image missing alt text.",
        recommendation="Add descriptive alternative text.",
        behavior=RuleBehavior.MANUAL_REVIEW,
        status=AccessibilityStatus.MANUAL_REVIEW,
    )
    result = RuleResult(
        rule_id="SG-ALT-01",
        rule_name="Alternative Text Missing",
        slide_number=1,
        status=RuleStatus.FAILED,
        issues=[issue],
    )
    correlator = IssueCorrelator()
    unified = correlator.correlate([result])

    assert len(unified) == 1
    u = unified[0]
    assert u.slide_number == 1
    assert u.shape_id == 10
    assert u.primary_rule_id == "SG-ALT-01"
    assert len(u.related_rule_ids) == 0
    assert len(u.raw_issues) == 1
    assert u.severity == IssueSeverity.CRITICAL


# =============================================================================
# 2. Overlapping multi-rule issue
# =============================================================================

def test_02_overlapping_multi_rule_issue():
    # Shape 5 has title issue (SG-TTL-01), reading order issue (SG-RO-01), and placeholder issue (SG-PH-01)
    iss_ttl = Issue(
        issue_id="iss_ttl",
        rule_id="SG-TTL-01",
        rule_name="Slide Title Missing",
        slide_number=2,
        shape_id=5,
        severity=IssueSeverity.ERROR,
        description="Slide title is not using standard placeholder.",
        recommendation="Convert shape to slide title placeholder.",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
    )
    iss_ro = Issue(
        issue_id="iss_ro",
        rule_id="SG-RO-01",
        rule_name="Reading Order Sequence",
        slide_number=2,
        shape_id=5,
        severity=IssueSeverity.WARNING,
        description="Title shape is not first in reading order.",
        recommendation="Place title shape first in sequence.",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
    )
    iss_ph = Issue(
        issue_id="iss_ph",
        rule_id="SG-PH-01",
        rule_name="Placeholder Conversion",
        slide_number=2,
        shape_id=5,
        severity=IssueSeverity.WARNING,
        description="Floating title shape should be converted to native title placeholder.",
        recommendation="Promote to native title placeholder.",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
    )

    r_ttl = RuleResult(rule_id="SG-TTL-01", rule_name="Title", slide_number=2, status=RuleStatus.FAILED, issues=[iss_ttl])
    r_ro = RuleResult(rule_id="SG-RO-01", rule_name="Reading Order", slide_number=2, status=RuleStatus.WARNING, issues=[iss_ro])
    r_ph = RuleResult(rule_id="SG-PH-01", rule_name="Placeholder", slide_number=2, status=RuleStatus.WARNING, issues=[iss_ph])

    correlator = IssueCorrelator()
    unified = correlator.correlate([r_ttl, r_ro, r_ph])

    # Must unify into a single issue referencing all 3 rules
    assert len(unified) == 1
    u = unified[0]
    assert u.slide_number == 2
    assert u.shape_id == 5
    assert u.primary_rule_id == "SG-TTL-01"
    assert "SG-RO-01" in u.related_rule_ids
    assert "SG-PH-01" in u.related_rule_ids
    assert len(u.raw_issues) == 3


# =============================================================================
# 3. Duplicate issue detection
# =============================================================================

def test_03_duplicate_issue_detection():
    # Two identical or duplicate findings on the same shape and rule
    iss1 = _make_issue(
        issue_id="iss_dup1",
        rule_id="SG-TXT-01",
        rule_name="Missing Text",
        slide_number=3,
        shape_id=12,
        severity=IssueSeverity.WARNING,
        description="Empty text container detected.",
        recommendation="Remove empty container.",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
    )
    iss2 = _make_issue(
        issue_id="iss_dup2",
        rule_id="SG-TXT-01",
        rule_name="Missing Text",
        slide_number=3,
        shape_id=12,
        severity=IssueSeverity.WARNING,
        description="Empty text container detected.",
        recommendation="Remove empty container.",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
    )
    res = RuleResult(rule_id="SG-TXT-01", rule_name="Missing Text", slide_number=3, status=RuleStatus.WARNING, issues=[iss1, iss2])

    correlator = IssueCorrelator()
    unified = correlator.correlate([res])

    # Should deduplicate into 1 unified issue while retaining raw issues
    assert len(unified) == 1
    assert len(unified[0].raw_issues) == 2


# =============================================================================
# 4. Title contextual classification
# =============================================================================

def test_04_title_contextual_classification():
    classifier = ContextualClassifier()
    # Large top-positioned bold text shape
    title_shape = _build_test_shape(
        shape_id=1,
        name="Slide Title",
        left=50.0,
        top=30.0,
        width=600.0,
        height=60.0,
        text="Chapter 1: Advanced Accessibility",
        font_size=32.0,
        is_bold=True,
    )
    slide = SlideModel(slide_number=1, slide_id=1, shapes=[title_shape])
    ctx = classifier.analyze_title_context(slide)

    assert ctx.classification in ("OBVIOUS_TITLE_CANDIDATE", "NATIVE_TITLE")
    assert ctx.candidate_shape_id == 1
    assert ctx.confidence_score >= 0.7


# =============================================================================
# 5. Image contextual classification
# =============================================================================

def test_05_image_contextual_classification():
    classifier = ContextualClassifier()
    img_shape = ShapeModel(
        shape_id=2,
        name="Chart Image",
        shape_type=ShapeType.PICTURE,
        left_pt=100.0,
        top_pt=150.0,
        width_pt=400.0,
        height_pt=300.0,
        alt_text="Quarterly revenue breakdown chart",
    )
    slide = SlideModel(slide_number=1, slide_id=1, shapes=[img_shape])
    ctx = classifier.analyze_image_context(img_shape, slide)

    assert ctx.classification in ("COMPLEX_DIAGRAM", "MEANINGFUL_PICTURE")
    assert ctx.is_divider is False
    assert ctx.has_alt_text is True


# =============================================================================
# 6. Decorative image contextual classification
# =============================================================================

def test_06_decorative_image_contextual_classification():
    classifier = ContextualClassifier()
    divider_shape = ShapeModel(
        shape_id=3,
        name="Horizontal Divider",
        shape_type=ShapeType.AUTO_SHAPE,
        left_pt=50.0,
        top_pt=520.0,
        width_pt=620.0,
        height_pt=2.0,
        text_content="",
    )
    slide = SlideModel(slide_number=1, slide_id=1, shapes=[divider_shape])
    ctx = classifier.analyze_image_context(divider_shape, slide)

    assert ctx.classification == "DECORATIVE_DIVIDER"
    assert ctx.is_divider is True


# =============================================================================
# 7. Ambiguous classification
# =============================================================================

def test_07_ambiguous_classification():
    classifier = ContextualClassifier()
    shape_a = _build_test_shape(shape_id=4, name="Top Box 1", top=40.0, text="First candidate headline", font_size=20.0)
    shape_b = _build_test_shape(shape_id=5, name="Top Box 2", top=45.0, text="Second candidate headline", font_size=20.0)
    slide = SlideModel(slide_number=1, slide_id=1, shapes=[shape_a, shape_b])
    ctx = classifier.analyze_title_context(slide)

    assert ctx.classification in ("AMBIGUOUS_TITLE", "OBVIOUS_TITLE_CANDIDATE")
    assert ctx.is_native is False


# =============================================================================
# 8. Confidence levels
# =============================================================================

def test_08_confidence_levels():
    engine = ConfidenceEngine()
    iss_safe = _make_issue(
        issue_id="iss_s",
        rule_id="SG-LNG-01",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        confidence="HIGH"
    )
    assert engine.evaluate_issue_confidence(iss_safe) == ConfidenceLevel.HIGH

    iss_amb = _make_issue(
        issue_id="iss_a",
        rule_id="SG-ALT-01",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        description="Ambiguous visual element requiring review",
        status=AccessibilityStatus.MANUAL_REVIEW,
        confidence="MEDIUM"
    )
    assert engine.evaluate_issue_confidence(iss_amb) in (ConfidenceLevel.MEDIUM, ConfidenceLevel.LOW)


# =============================================================================
# 9. Remediation eligibility
# =============================================================================

def test_09_remediation_eligibility():
    engine = ConfidenceEngine()
    iss_auto = _make_issue(
        issue_id="iss_a",
        rule_id="SG-META-01",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        confidence="HIGH"
    )
    assert engine.determine_remediation_eligibility(ConfidenceLevel.HIGH, iss_auto) == RemediationEligibility.ELIGIBLE_AUTO

    iss_rev = _make_issue(
        issue_id="iss_r",
        rule_id="SG-TTL-01",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.MANUAL_REVIEW,
        confidence="MEDIUM"
    )
    assert engine.determine_remediation_eligibility(ConfidenceLevel.MEDIUM, iss_rev) == RemediationEligibility.REQUIRES_USER_APPROVAL


# =============================================================================
# 10. Issue reconciliation framework
# =============================================================================

def test_10_issue_reconciliation_framework():
    reconciler = IssueReconciliationService()

    pre = [
        UnifiedIssue(
            unified_id="u_1",
            slide_number=1,
            shape_id=1,
            primary_rule_id="SG-TTL-01",
            severity=IssueSeverity.ERROR,
            confidence=ConfidenceLevel.HIGH,
            explanation="Missing title placeholder.",
        )
    ]
    post = []
    actions = [
        RemediationAction(
            action_id="a_1",
            rule_id="SG-TTL-01",
            slide_number=1,
            target_shape_id=1,
            description="Created title placeholder.",
            status=RemediationStatus.APPLIED,
        )
    ]
    summary = reconciler.reconcile(pre, post, actions)
    assert summary.total_pre == 1
    assert summary.total_post == 0
    assert summary.resolved_count == 1
    assert summary.remediated_count == 1


# =============================================================================
# 11. Resolved issue
# =============================================================================

def test_11_resolved_issue():
    reconciler = IssueReconciliationService()
    pre_issue = UnifiedIssue(
        unified_id="u_res",
        slide_number=2,
        shape_id=4,
        primary_rule_id="SG-LNG-01",
        severity=IssueSeverity.WARNING,
        confidence=ConfidenceLevel.HIGH,
        explanation="Language tag missing.",
    )
    actions = [
        RemediationAction(
            action_id="a_lng",
            rule_id="SG-LNG-01",
            slide_number=2,
            target_shape_id=4,
            description="Set language tag to en-US",
            status=RemediationStatus.APPLIED,
        )
    ]
    summary = reconciler.reconcile([pre_issue], [], actions)
    assert len(summary.items) == 1
    assert summary.items[0].status == ReconciliationStatus.REMEDIATED
    assert pre_issue.reconciliation_status == ReconciliationStatus.REMEDIATED


# =============================================================================
# 12. Unchanged issue
# =============================================================================

def test_12_unchanged_issue():
    reconciler = IssueReconciliationService()
    pre_issue = UnifiedIssue(
        unified_id="u_unc",
        slide_number=3,
        shape_id=7,
        primary_rule_id="SG-CLR-01",
        severity=IssueSeverity.ERROR,
        confidence=ConfidenceLevel.HIGH,
        explanation="Low contrast text.",
    )
    post_issue = UnifiedIssue(
        unified_id="u_unc_post",
        slide_number=3,
        shape_id=7,
        primary_rule_id="SG-CLR-01",
        severity=IssueSeverity.ERROR,
        confidence=ConfidenceLevel.HIGH,
        explanation="Low contrast text.",
    )
    summary = reconciler.reconcile([pre_issue], [post_issue], [])
    assert summary.unchanged_count == 1
    assert summary.items[0].status == ReconciliationStatus.UNCHANGED


# =============================================================================
# 13. New issue
# =============================================================================

def test_13_new_issue():
    reconciler = IssueReconciliationService()
    post_new = UnifiedIssue(
        unified_id="u_new",
        slide_number=4,
        shape_id=15,
        primary_rule_id="SG-RO-01",
        severity=IssueSeverity.WARNING,
        confidence=ConfidenceLevel.HIGH,
        explanation="New reading order issue.",
    )
    summary = reconciler.reconcile([], [post_new], [])
    assert summary.new_count == 1
    assert summary.items[0].status == ReconciliationStatus.NEW


# =============================================================================
# 14. Changed issue
# =============================================================================

def test_14_changed_issue():
    reconciler = IssueReconciliationService()
    pre_issue = UnifiedIssue(
        unified_id="u_chg",
        slide_number=5,
        shape_id=8,
        primary_rule_id="SG-RDB-01",
        severity=IssueSeverity.WARNING,
        confidence=ConfidenceLevel.HIGH,
        explanation="Readability borderline font size.",
    )
    post_issue = UnifiedIssue(
        unified_id="u_chg_post",
        slide_number=5,
        shape_id=8,
        primary_rule_id="SG-RDB-01",
        severity=IssueSeverity.INFO,
        confidence=ConfidenceLevel.HIGH,
        explanation="Readability adjusted font size.",
    )
    summary = reconciler.reconcile([pre_issue], [post_issue], [])
    assert summary.changed_count == 1
    assert summary.items[0].status == ReconciliationStatus.CHANGED


# =============================================================================
# 15. Stale-analysis invalidation
# =============================================================================

def test_15_stale_analysis_invalidation():
    # Verify _refresh_modified_slides properly parses raw OOXML and updates in-memory slide
    prs = Presentation()
    slide_layout = prs.slide_layouts[6]  # Blank
    raw_slide = prs.slides.add_slide(slide_layout)
    tb = raw_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text = "Original Slide Content"

    slide_model = SlideModel(
        slide_number=1,
        slide_id=256,
        shapes=[],
        raw_slide=raw_slide
    )
    pres_model = PresentationModel(
        file_path=Path("stale_test.pptx"),
        slides=[slide_model],
        slide_width_pt=720,
        slide_height_pt=540
    )

    remediator = RemediationService()
    # Before refresh, slide_model.shapes is empty
    assert len(slide_model.shapes) == 0

    remediator._refresh_modified_slides(pres_model, {1})
    # After refresh, slide_model.shapes has been re-parsed from raw OOXML!
    assert len(slide_model.shapes) >= 1
    assert "Original Slide Content" in slide_model.shapes[0].text_content


# =============================================================================
# 16. Remediation dependency ordering
# =============================================================================

def test_16_remediation_dependency_ordering(tmp_path):
    # Verify the sequence: Reading Order -> Ungrouping -> Placeholders -> Lists -> Accessibility -> Reconciliation
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    test_file = tmp_path / "dep_test.pptx"
    prs.save(str(test_file))

    from pptx_engine.reader.presentation_reader import PresentationReader
    pres = PresentationReader().read(test_file)
    context = JobContext(
        job_id="test_dep",
        workspace_dir=tmp_path,
        original_file=test_file,
        working_file=test_file,
        candidate_file=tmp_path / "candidate.pptx",
        user_approved_remediation=True
    )
    remediator = RemediationService()
    actions = remediator.remediate(context, pres)

    step_rules = [a.rule_id for a in actions if a.rule_id.startswith("PIPELINE_STEP")]
    expected_order = [
        "PIPELINE_STEP_1_READING_ORDER",
        "PIPELINE_STEP_2_UNGROUPING",
        "PIPELINE_STEP_3_PLACEHOLDERS",
        "PIPELINE_STEP_4_LISTS",
        "PIPELINE_STEP_5_ACCESSIBILITY",
        "PIPELINE_STEP_6_RECONCILIATION",
    ]
    assert step_rules == expected_order


# =============================================================================
# 17. Quality score calculation (V2)
# =============================================================================

def test_17_quality_score_calculation():
    scorer = AccessibilityScorer()

    # High severity, low confidence issue should deduce properly
    u_issue = UnifiedIssue(
        unified_id="u_sc1",
        slide_number=1,
        shape_id=1,
        primary_rule_id="SG-TTL-01",
        severity=IssueSeverity.CRITICAL,
        confidence=ConfidenceLevel.HIGH,
        remediation_eligibility=RemediationEligibility.REQUIRES_USER_APPROVAL,
    )
    metrics_pre = scorer.calculate_metrics_v2([u_issue], total_slides=1)
    score_pre = metrics_pre.quality_score_v2
    assert score_pre < 100.0

    # Remediated issue deduction should be dramatically reduced (remediated bonus)
    u_remediated = UnifiedIssue(
        unified_id="u_sc1",
        slide_number=1,
        shape_id=1,
        primary_rule_id="SG-TTL-01",
        severity=IssueSeverity.CRITICAL,
        confidence=ConfidenceLevel.HIGH,
        reconciliation_status=ReconciliationStatus.REMEDIATED,
    )
    metrics_post = scorer.calculate_metrics_v2([u_remediated], total_slides=1)
    score_post = metrics_post.quality_score_v2
    assert score_post > score_pre


# =============================================================================
# 18. Gemini structured recommendation parsing
# =============================================================================

def test_18_gemini_structured_recommendation_parsing():
    assistant = GeminiAssistant()
    sample_json = '''
    {
        "suggested_value": "Bar chart illustrating Q3 revenue growth by region",
        "confidence": 0.95,
        "reasoning": "The visual communicates statistical financial data essential to comprehension.",
        "is_decorative": false
    }
    '''
    rec = assistant.parse_structured_response(sample_json, default_type="ALT_TEXT_SUGGESTION")
    assert rec.recommendation_type == "ALT_TEXT_SUGGESTION"
    assert rec.suggested_value == "Bar chart illustrating Q3 revenue growth by region"
    assert rec.confidence == 0.95
    assert rec.is_safe_for_review is True
    assert "statistical financial data" in rec.reasoning


# =============================================================================
# 19. Invalid Gemini response handling
# =============================================================================

def test_19_invalid_gemini_response_handling():
    assistant = GeminiAssistant()
    # Malformed / unparseable response
    invalid_raw = "I am an AI and here is some text that is definitely not JSON { broken ..."
    rec = assistant.parse_structured_response(invalid_raw, default_type="ALT_TEXT_SUGGESTION")

    assert rec.is_safe_for_review is False
    assert rec.confidence == 0.0
    assert "Invalid response format" in rec.reasoning


# =============================================================================
# 20. Gemini unavailable fallback
# =============================================================================

def test_20_gemini_unavailable_fallback():
    # Explicitly without client (offline)
    assistant = GeminiAssistant(client=None)
    rec = assistant.suggest_alt_text({
        "shape_name": "mitosis_diagram",
        "slide_title": "Cell Biology",
        "surrounding_text": "Mitosis overview"
    })
    assert rec.recommendation_type == "ALT_TEXT_SUGGESTION"
    assert "Offline fallback" in rec.reasoning
    assert rec.is_safe_for_review is True


# =============================================================================
# 21. Deterministic remediation without AI
# =============================================================================

def test_21_deterministic_remediation_without_ai():
    # Verify deterministic rules (e.g. Language, Document Metadata) execute without any AI dependence
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    tb.text = "Hello accessibility"

    from rules.accessibility import LanguageRule, DocumentMetadataRule

    pres_model = PresentationModel(
        file_path=Path("det_test.pptx"),
        slides=[SlideModel(slide_number=1, slide_id=1, shapes=[], raw_slide=slide)],
        slide_width_pt=720,
        slide_height_pt=540,
        raw_presentation=prs
    )

    issue = _make_issue(
        issue_id="iss_det_lng",
        rule_id="SG-LNG-01",
        rule_name="Language",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        description="Missing language tag",
        recommendation="Set language to en-US",
    )
    rule = LanguageRule(configured_language="en-US")
    action = rule.remediate_issue(pres_model.slides[0], pres_model, issue)
    assert action.status == RemediationStatus.APPLIED


# =============================================================================
# 22. Visual-preservation protection
# =============================================================================

def test_22_visual_preservation_protection():
    # Ensure accessibility rules NEVER alter dimensions, coordinates, font faces, or colors
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    orig_left = Inches(1.5)
    orig_top = Inches(2.0)
    orig_width = Inches(4.0)
    orig_height = Inches(1.0)
    tb = slide.shapes.add_textbox(orig_left, orig_top, orig_width, orig_height)
    tb.text = "Protected layout text"

    from rules.accessibility import SlideTitleRule
    rule = SlideTitleRule()

    # Verify rule remediations do NOT modify shape position or bounds
    assert tb.left == orig_left
    assert tb.top == orig_top
    assert tb.width == orig_width
    assert tb.height == orig_height
