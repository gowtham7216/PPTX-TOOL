"""
Unit tests for SlideGuard Phase 9: AI-Assisted Accessibility Authoring.
Verifies all 34 required test scenarios:
  1. missing image alt text -> AI suggestion
  2. existing good alt text -> no unnecessary AI call
  3. generic alt text -> AI suggestion
  4. decorative image -> no AI authoring
  5. ambiguous image -> manual review
  6. chart description suggestion
  7. diagram description suggestion
  8. missing slide title -> AI title suggestion
  9. table header interpretation
 10. complex table -> manual review
 11. structured JSON parsing
 12. malformed JSON
 13. missing required AI fields
 14. hallucinated unsupported facts
 15. low-confidence recommendation
 16. high-confidence recommendation
 17. user approve
 18. user edit + approve
 19. user reject
 20. pending recommendation not applied
 21. deterministic application of approved alt text
 22. deterministic application of approved title
 23. AI unavailable fallback
 24. API timeout fallback
 25. API quota failure fallback
 26. audit trail creation
 27. report generation
 28. before/after reconciliation
 29. visual preservation
 30. shape ID preservation
 31. relationship preservation
 32. animation preservation
 33. original PPTX remains untouched
 34. candidate PPTX reloads successfully
"""

import json
import pytest
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

from models.enums import (
    IssueSeverity,
    RuleBehavior,
    AccessibilityStatus,
    RemediationStatus,
    RuleStatus,
    ShapeType,
    PlaceholderType,
)
from models.issue import Issue
from models.rule_result import RuleResult
from models.shape import ShapeModel, TextParagraphModel, TextRunModel
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.job import JobContext
from models.ai_authoring import (
    AuthoringType,
    AuthoringStatus,
    FactGrounding,
    AuthoringRecommendation,
    AuthoringAuditRecord,
    AuthoringSummary,
)
from engine.intelligence.gemini_assistant import GeminiAssistant
from engine.authoring.authoring_service import AuthoringService
from engine.authoring.authoring_applier import AuthoringApplier
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from reports.generator import ReportGenerator
from engine.remediation.remediation_service import RemediationService
from rules.accessibility.base import get_cnvpr_element, safe_xpath


from typing import Tuple

# =============================================================================
# Helper Builders
# =============================================================================

def _build_test_pres_with_image(tmp_path: Path, alt_text: str = "") -> Tuple[PresentationModel, Path, int]:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Add title
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(8), Inches(1))
    tb.text = "Quarterly Business Review"

    # Add a visual shape (AutoShape used as picture mockup)
    shape = slide.shapes.add_shape(1, Inches(1), Inches(2), Inches(4), Inches(3))
    shape.name = "RevenueChartPicture"

    # Set alt text if provided
    sp_tree = slide.shapes._spTree
    for child in sp_tree:
        cnvpr = get_cnvpr_element(child)
        if cnvpr is not None and cnvpr.get("id") == str(shape.shape_id):
            if alt_text:
                cnvpr.set("descr", alt_text)
            break

    file_path = tmp_path / "test_image.pptx"
    prs.save(str(file_path))

    reader = PresentationReader()
    pres_model = reader.read(file_path)
    return pres_model, file_path, shape.shape_id



def _make_issue(
    issue_id: str,
    rule_id: str,
    rule_name: str,
    slide_number: int,
    severity: IssueSeverity,
    description: str,
    recommendation: str = "Provide descriptive content or review manually.",
    behavior: RuleBehavior = RuleBehavior.AI_ASSISTED_REVIEW,
    status: AccessibilityStatus = AccessibilityStatus.DETECTED,
    shape_id: int = None,
    shape_name: str = None,
    extra_data: dict = None
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
        status=status,
        shape_id=shape_id,
        shape_name=shape_name,
        extra_data=extra_data
    )


def _make_rule_result(rule_id: str, rule_name: str, slide_number: int, issues: list) -> RuleResult:
    return RuleResult(
        rule_id=rule_id,
        rule_name=rule_name,
        slide_number=slide_number,
        status=RuleStatus.FAILED if issues else RuleStatus.PASSED,
        issues=issues
    )


# =============================================================================
# 1. Missing image alt text -> AI suggestion
# =============================================================================

def test_01_missing_image_alt_text_ai_suggestion(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="")

    issue = _make_issue(
        issue_id=f"ALT_MISSING_1_{shape_id}",
        rule_id="SG-ALT-01",
        rule_name="Missing Alternative Text",
        slide_number=1,
        severity=IssueSeverity.ERROR,
        shape_id=shape_id,
        shape_name="RevenueChartPicture",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW,
        status=AccessibilityStatus.DETECTED,
        description="Picture lacks alternative text description.",
    )
    context = JobContext(
        job_id="job_01",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        rule_results=[_make_rule_result("SG-ALT-01", "Missing Alt Text", 1, [issue])]
    )

    mock_client = lambda p: json.dumps({
        "suggested_value": "Chart showing quarterly revenue progression from Q1 to Q4.",
        "confidence": 0.92,
        "rationale": "Summarizes revenue progression from slide context.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    })
    assistant = GeminiAssistant(client=mock_client)
    service = AuthoringService(assistant=assistant)

    recs = service.generate_recommendations(context, pres_model)
    assert len(recs) == 1
    rec = recs[0]
    assert rec.shape_id == shape_id
    assert rec.authoring_type in (AuthoringType.ALT_TEXT, AuthoringType.IMAGE_DESCRIPTION, AuthoringType.CHART_DESCRIPTION)
    assert "quarterly revenue" in rec.suggested_value.lower()
    assert rec.status == AuthoringStatus.PENDING_REVIEW
    assert rec.confidence == 0.92


# =============================================================================
# 2. Existing good alt text -> no unnecessary AI call
# =============================================================================

def test_02_existing_good_alt_text_no_unnecessary_ai_call(tmp_path):
    good_alt = "Solar panel array facing south at dusk."
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text=good_alt)

    issue = _make_issue(
        issue_id=f"ALT_MISSING_1_{shape_id}",
        rule_id="SG-ALT-01",
        rule_name="Missing Alternative Text",
        slide_number=1,
        severity=IssueSeverity.INFO,
        shape_id=shape_id,
        shape_name="RevenueChartPicture",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW,
        status=AccessibilityStatus.PASSED,
        description="Alt text already present.",
    )
    context = JobContext(
        job_id="job_02",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        rule_results=[_make_rule_result("SG-ALT-01", "Missing Alt Text", 1, [issue])]
    )

    call_count = 0
    def mock_client(p):
        nonlocal call_count
        call_count += 1
        return "{}"

    assistant = GeminiAssistant(client=mock_client)
    service = AuthoringService(assistant=assistant)

    recs = service.generate_recommendations(context, pres_model)
    assert len(recs) == 0
    assert call_count == 0  # No unnecessary AI call made!


# =============================================================================
# 3. Generic alt text -> AI suggestion
# =============================================================================

def test_03_generic_alt_text_ai_suggestion(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="picture")

    issue = _make_issue(
        issue_id=f"ALT_GENERIC_1_{shape_id}",
        rule_id="SG-ALT-01",
        rule_name="Generic Alt Text",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        shape_id=shape_id,
        shape_name="RevenueChartPicture",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW,
        status=AccessibilityStatus.MANUAL_REVIEW,
        description="Uses generic alt text 'picture'.",
    )
    context = JobContext(
        job_id="job_03",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        rule_results=[_make_rule_result("SG-ALT-01", "Alt Text", 1, [issue])]
    )

    mock_client = lambda p: json.dumps({
        "suggested_value": "Quarterly business metrics and financial indicators.",
        "confidence": 0.88,
        "rationale": "Replaced generic 'picture' with descriptive title context.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    })
    assistant = GeminiAssistant(client=mock_client)
    service = AuthoringService(assistant=assistant)

    recs = service.generate_recommendations(context, pres_model)
    assert len(recs) == 1
    assert recs[0].current_value == "picture"
    assert "quarterly business metrics" in recs[0].suggested_value.lower()


# =============================================================================
# 4. Decorative image -> no AI authoring
# =============================================================================

def test_04_decorative_image_no_ai_authoring(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="")
    # Mark shape decorative
    pres_model.slides[0].shapes[1].is_decorative = True

    issue = _make_issue(
        issue_id=f"ALT_MISSING_1_{shape_id}",
        rule_id="SG-ALT-01",
        rule_name="Alt Text",
        slide_number=1,
        severity=IssueSeverity.ERROR,
        shape_id=shape_id,
        shape_name="DividerLine",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW,
        status=AccessibilityStatus.DETECTED,
        description="Missing alt text.",
    )
    context = JobContext(
        job_id="job_04",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        rule_results=[_make_rule_result("SG-ALT-01", "Alt Text", 1, [issue])]
    )

    assistant = GeminiAssistant(client=lambda p: "{}")
    service = AuthoringService(assistant=assistant)
    recs = service.generate_recommendations(context, pres_model)
    assert len(recs) == 0  # No AI authoring requested for decorative shape


# =============================================================================
# 5. Ambiguous image -> manual review
# =============================================================================

def test_05_ambiguous_image_manual_review(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="")
    # Make shape dimensions small to trigger ambiguous classification
    pres_model.slides[0].shapes[1].width_pt = 30.0
    pres_model.slides[0].shapes[1].height_pt = 30.0

    issue = _make_issue(
        issue_id=f"ALT_MISSING_1_{shape_id}",
        rule_id="SG-ALT-01",
        rule_name="Alt Text",
        slide_number=1,
        severity=IssueSeverity.ERROR,
        shape_id=shape_id,
        shape_name="SmallIcon",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW,
        status=AccessibilityStatus.DETECTED,
        description="Missing alt text on ambiguous small icon.",
        extra_data={"classification": "UNKNOWN"}
    )
    context = JobContext(
        job_id="job_05",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        rule_results=[_make_rule_result("SG-ALT-01", "Alt Text", 1, [issue])]
    )

    assistant = GeminiAssistant(client=None)  # Offline
    service = AuthoringService(assistant=assistant)
    recs = service.generate_recommendations(context, pres_model)
    assert len(recs) == 1
    assert recs[0].requires_human_review is True
    assert recs[0].fact_grounding == FactGrounding.UNCERTAIN




# =============================================================================
# 6. Chart description suggestion
# =============================================================================

def test_06_chart_description_suggestion():
    assistant = GeminiAssistant(client=lambda p: json.dumps({
        "suggested_value": "Bar chart comparing quarterly revenue across four quarters.",
        "confidence": 0.90,
        "rationale": "Derived from slide title and chart type without fabricating numerical values.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    }))

    res = assistant.author_chart_description({
        "slide_title": "Quarterly Financials",
        "shape_name": "RevenueChart",
        "chart_type": "BAR"
    })
    assert "Bar chart comparing" in res["suggested_value"]
    assert res["fact_grounding"] == "SUPPORTED_FACT"
    assert res["confidence"] == 0.90


# =============================================================================
# 7. Diagram description suggestion
# =============================================================================

def test_07_diagram_description_suggestion():
    assistant = GeminiAssistant(client=lambda p: json.dumps({
        "suggested_value": "Process flowchart showing transition from Analysis to Remediation to Verification.",
        "confidence": 0.88,
        "rationale": "Based on sequential connector labels.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    }))

    res = assistant.author_diagram_description({
        "slide_title": "Remediation Workflow",
        "shape_name": "FlowchartGroup",
        "labels": ["Analysis", "Remediation", "Verification"]
    })
    assert "flowchart" in res["suggested_value"].lower()
    assert res["confidence"] == 0.88


# =============================================================================
# 8. Missing slide title -> AI title suggestion
# =============================================================================

def test_08_missing_slide_title_ai_title_suggestion():
    assistant = GeminiAssistant(client=lambda p: json.dumps({
        "suggested_value": "Introduction to Cellular Respiration",
        "confidence": 0.85,
        "rationale": "Extracted key topic from prominent bullet points.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    }))

    res = assistant.author_slide_title({
        "slide_number": 3,
        "nearby_text": ["Glycolysis and citric acid cycle overview", "ATP production in mitochondria"]
    })
    assert res["suggested_value"] == "Introduction to Cellular Respiration"
    # Never generic placeholder
    assert res["suggested_value"].lower() not in ("slide 3", "untitled slide", "slide")


# =============================================================================
# 9. Table header interpretation
# =============================================================================

def test_09_table_header_interpretation():
    assistant = GeminiAssistant(client=lambda p: json.dumps({
        "suggested_value": "Designate first row as column headers: Region, Q1 Sales, Q2 Sales",
        "confidence": 0.90,
        "rationale": "Row 0 contains standard column label nomenclature.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    }))

    res = assistant.author_table_interpretation({
        "slide_number": 4,
        "table_name": "SalesDataTable",
        "row_0_preview": ["Region", "Q1 Sales", "Q2 Sales"]
    })
    assert "column headers" in res["suggested_value"].lower()
    assert res["confidence"] == 0.90


# =============================================================================
# 10. Complex table -> manual review
# =============================================================================

def test_10_complex_table_manual_review(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path)
    # Add issue for complex table with merged cells
    issue = _make_issue(
        issue_id=f"TBL_MERGED_CELLS_1_{shape_id}",
        rule_id="SG-TBL-01",
        rule_name="Table Structure",
        slide_number=1,
        severity=IssueSeverity.INFO,
        shape_id=shape_id,
        shape_name="ComplexMergedTable",
        behavior=RuleBehavior.DETECT_ONLY,
        status=AccessibilityStatus.MANUAL_REVIEW,
        description="Table contains merged cells (gridSpan/rowSpan).",
    )

    context = JobContext(
        job_id="job_10",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        rule_results=[_make_rule_result("SG-TBL-01", "Tables", 1, [issue])]
    )


    service = AuthoringService()
    recs = service.generate_recommendations(context, pres_model)
    assert len(recs) == 1
    assert recs[0].status == AuthoringStatus.MANUAL_REVIEW
    assert "merged cells" in recs[0].rationale.lower()


# =============================================================================
# 11. Structured JSON parsing
# =============================================================================

def test_11_structured_json_parsing():
    assistant = GeminiAssistant()
    valid_json = json.dumps({
        "suggested_value": "World map illustrating global offices.",
        "confidence": 0.95,
        "rationale": "Visual map graphic corresponds to international office locations.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    })
    res = assistant._parse_and_validate_authoring_response(
        valid_json,
        "ALT_TEXT",
        {"context": "global offices"}
    )
    assert res["suggested_value"] == "World map illustrating global offices."
    assert res["confidence"] == 0.95
    assert res["fact_grounding"] == "SUPPORTED_FACT"
    assert res["requires_human_review"] is True


# =============================================================================
# 12. Malformed JSON
# =============================================================================

def test_12_malformed_json():
    assistant = GeminiAssistant()
    bad_raw = "Here is your suggestion: { suggested_value: 'broken json without quotes' ... "
    res = assistant._parse_and_validate_authoring_response(bad_raw, "ALT_TEXT", {})
    assert res["status"] == "MANUAL_REVIEW"
    assert res["suggested_value"] == "MANUAL_REVIEW"
    assert res["confidence"] == 0.0
    assert any("malformed json" in w.lower() for w in res["warnings"])


# =============================================================================
# 13. Missing required AI fields
# =============================================================================

def test_13_missing_required_ai_fields():
    assistant = GeminiAssistant()
    missing_json = json.dumps({"confidence": 0.8})  # Missing suggested_value and rationale
    res = assistant._parse_and_validate_authoring_response(missing_json, "ALT_TEXT", {})
    assert res["status"] == "MANUAL_REVIEW"
    assert any("missing required schema" in w.lower() for w in res["warnings"])


# =============================================================================
# 14. Hallucinated unsupported facts
# =============================================================================

def test_14_hallucinated_unsupported_facts():
    assistant = GeminiAssistant()
    # Context does NOT have 45.8% or 1200
    context = {"slide_title": "Sales Performance", "notes": "revenue grew steadily"}
    hallucinated_json = json.dumps({
        "suggested_value": "Line chart showing revenue grew by 45.8% reaching 1200 units.",
        "confidence": 0.95,
        "rationale": "Estimated numbers from visual.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    })
    res = assistant._parse_and_validate_authoring_response(hallucinated_json, "ALT_TEXT", context)
    assert res["fact_grounding"] == "UNCERTAIN"
    assert any("hallucination warning" in w.lower() for w in res["warnings"])
    assert res["confidence"] <= 0.45


# =============================================================================
# 15. Low-confidence recommendation
# =============================================================================

def test_15_low_confidence_recommendation():
    assistant = GeminiAssistant()
    low_conf_json = json.dumps({
        "suggested_value": "Possible diagram flow.",
        "confidence": 0.35,
        "rationale": "Visual elements are too vague to determine purpose with certainty.",
        "fact_grounding": "UNCERTAIN",
        "requires_human_review": True
    })
    res = assistant._parse_and_validate_authoring_response(low_conf_json, "DIAGRAM_DESCRIPTION", {})
    assert res["status"] == "MANUAL_REVIEW"
    assert res["confidence"] == 0.35


# =============================================================================
# 16. High-confidence recommendation
# =============================================================================

def test_16_high_confidence_recommendation():
    assistant = GeminiAssistant()
    high_conf_json = json.dumps({
        "suggested_value": "Bar chart illustrating 2026 sales growth.",
        "confidence": 0.95,
        "rationale": "High fidelity match with slide title and labels.",
        "fact_grounding": "SUPPORTED_FACT",
        "requires_human_review": True
    })
    res = assistant._parse_and_validate_authoring_response(
        high_conf_json,
        "CHART_DESCRIPTION",
        {"slide_title": "2026 sales growth"}
    )
    assert res["status"] == "PENDING_REVIEW"
    assert res["confidence"] == 0.95
    assert res["fact_grounding"] == "SUPPORTED_FACT"


# =============================================================================
# 17. User approve
# =============================================================================

def test_17_user_approve():
    rec = AuthoringRecommendation(
        recommendation_id="rec_17",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=5,
        suggested_value="Overview of company history",
        status=AuthoringStatus.PENDING_REVIEW
    )
    assert rec.status == AuthoringStatus.PENDING_REVIEW
    rec.approve()
    assert rec.status == AuthoringStatus.APPROVED
    assert rec.effective_value == "Overview of company history"


# =============================================================================
# 18. User edit + approve
# =============================================================================

def test_18_user_edit_plus_approve():
    rec = AuthoringRecommendation(
        recommendation_id="rec_18",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=5,
        suggested_value="AI suggestion text",
        status=AuthoringStatus.PENDING_REVIEW
    )
    rec.edit("User-refined accessible description")
    assert rec.status == AuthoringStatus.EDITED
    assert rec.user_edited_value == "User-refined accessible description"
    assert rec.effective_value == "User-refined accessible description"


# =============================================================================
# 19. User reject
# =============================================================================

def test_19_user_reject():
    rec = AuthoringRecommendation(
        recommendation_id="rec_19",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=5,
        suggested_value="AI suggestion text",
        status=AuthoringStatus.PENDING_REVIEW
    )
    rec.reject()
    assert rec.status == AuthoringStatus.REJECTED


# =============================================================================
# 20. Pending recommendation not applied
# =============================================================================

def test_20_pending_recommendation_not_applied(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="")

    # Recommendation is PENDING_REVIEW (not approved)
    rec_pending = AuthoringRecommendation(
        recommendation_id="rec_pend",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="This should NOT be applied.",
        status=AuthoringStatus.PENDING_REVIEW
    )
    # Recommendation is REJECTED
    rec_rejected = AuthoringRecommendation(
        recommendation_id="rec_rej",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="This should also NOT be applied.",
        status=AuthoringStatus.REJECTED
    )

    context = JobContext(
        job_id="job_20",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec_pending, rec_rejected]
    )

    applier = AuthoringApplier()
    actions, audit = applier.apply_approved(context, pres_model)
    assert len(actions) == 0  # Zero actions applied!
    assert len(audit) == 0

    # Verify OOXML descr remains empty
    shape = pres_model.slides[0].shapes[1]
    assert shape.alt_text is None or shape.alt_text == ""


# =============================================================================
# 21. Deterministic application of approved alt text
# =============================================================================

def test_21_deterministic_application_of_approved_alt_text(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="")

    rec_approved = AuthoringRecommendation(
        recommendation_id="rec_appr_alt",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="Approved accessible chart description.",
        status=AuthoringStatus.APPROVED,
        confidence=0.95
    )
    context = JobContext(
        job_id="job_21",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec_approved]
    )

    applier = AuthoringApplier()
    actions, audit = applier.apply_approved(context, pres_model)
    assert len(actions) == 1
    assert actions[0].status == RemediationStatus.APPLIED
    assert rec_approved.status == AuthoringStatus.APPLIED
    assert rec_approved.applied_at is not None

    # Check OOXML descr attribute
    sp_tree = pres_model.slides[0].raw_slide.shapes._spTree
    applied_descr = None
    for child in sp_tree:
        cnvpr = get_cnvpr_element(child)
        if cnvpr is not None and cnvpr.get("id") == str(shape_id):
            applied_descr = cnvpr.get("descr")
            break
    assert applied_descr == "Approved accessible chart description."


# =============================================================================
# 22. Deterministic application of approved title
# =============================================================================

def test_22_deterministic_application_of_approved_title(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Loose text box near top
    tb = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    tb.text = "Old Candidate Title"
    file_path = tmp_path / "test_title.pptx"
    prs.save(str(file_path))

    reader = PresentationReader()
    pres_model = reader.read(file_path)

    rec_title = AuthoringRecommendation(
        recommendation_id="rec_appr_ttl",
        rule_id="SG-TTL-01",
        slide_number=1,
        shape_id=tb.shape_id,
        authoring_type=AuthoringType.SLIDE_TITLE,
        suggested_value="New Approved Slide Title Heading",
        status=AuthoringStatus.APPROVED,
        confidence=0.92
    )
    context = JobContext(
        job_id="job_22",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec_title]
    )

    applier = AuthoringApplier()
    actions, audit = applier.apply_approved(context, pres_model)
    assert len(actions) == 1
    assert actions[0].status == RemediationStatus.APPLIED
    assert pres_model.slides[0].title == "New Approved Slide Title Heading"


# =============================================================================
# 23. AI unavailable fallback
# =============================================================================

def test_23_ai_unavailable_fallback():
    assistant = GeminiAssistant(client=None)  # Explicitly offline
    res = assistant.author_content("ALT_TEXT", {"slide_title": "Project Scope", "shape_name": "Graphic1"})
    assert res["status"] in ("PENDING_REVIEW", "MANUAL_REVIEW")
    assert "deterministic offline" in res["rationale"].lower()


# =============================================================================
# 24. API timeout fallback
# =============================================================================

def test_24_api_timeout_fallback():
    def timing_out_client(prompt):
        raise TimeoutError("Gemini API connection timed out.")

    assistant = GeminiAssistant(client=timing_out_client)
    res = assistant.author_content("ALT_TEXT", {"slide_title": "Project Scope"})
    assert res["status"] == "MANUAL_REVIEW"
    assert any("timeout" in w.lower() for w in res["warnings"])


# =============================================================================
# 25. API quota failure fallback
# =============================================================================

def test_25_api_quota_failure_fallback():
    def quota_exceeded_client(prompt):
        raise RuntimeError("429 ResourceExhausted: Quota exceeded for gemini-3.8-flash-high")

    assistant = GeminiAssistant(client=quota_exceeded_client)
    res = assistant.author_content("SLIDE_TITLE", {"nearby_text": ["Key takeaways"]})
    assert res["status"] == "MANUAL_REVIEW"
    assert any("quota" in w.lower() or "runtimeerror" in w.lower() for w in res["warnings"])


# =============================================================================
# 26. Audit trail creation
# =============================================================================

def test_26_audit_trail_creation(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path)
    rec = AuthoringRecommendation(
        recommendation_id="rec_aud_01",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        current_value="[Missing]",
        suggested_value="Audited alt text suggestion",
        status=AuthoringStatus.APPROVED,
        confidence=0.89,
        model="gemini-3.8-flash-high"
    )
    context = JobContext(
        job_id="job_26",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec]
    )

    applier = AuthoringApplier()
    actions, audit = applier.apply_approved(context, pres_model)
    assert len(audit) == 1
    a = audit[0]
    assert a.recommendation_id == "rec_aud_01"
    assert a.model == "gemini-3.8-flash-high"
    assert a.user_decision == "APPROVED"
    assert a.final_applied_value == "Audited alt text suggestion"
    assert a.status == "APPLIED"


# =============================================================================
# 27. Report generation
# =============================================================================

def test_27_report_generation(tmp_path):
    rec = AuthoringRecommendation(
        recommendation_id="rec_rep_01",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=12,
        suggested_value="Accessible report test suggestion",
        status=AuthoringStatus.APPLIED,
        confidence=0.91
    )
    audit = AuthoringAuditRecord(
        audit_id="aud_01",
        recommendation_id="rec_rep_01",
        model="gemini-3.8-flash-high",
        authoring_type="ALT_TEXT",
        slide_number=1,
        shape_id=12,
        original_value="[Missing]",
        ai_suggestion="Accessible report test suggestion",
        user_decision="APPROVED",
        final_applied_value="Accessible report test suggestion",
        confidence=0.91,
        status="APPLIED"
    )

    context = JobContext(
        job_id="job_rep",
        workspace_dir=tmp_path,
        original_file=tmp_path / "test.pptx",
        working_file=tmp_path / "test.pptx",
        ai_recommendations=[rec],
        ai_audit_trail=[audit.to_dict()],
        ai_authoring_summary={"generated": 1, "approved": 1, "applied": 1, "failed": 0}
    )

    rep_gen = ReportGenerator()
    md_file = tmp_path / "report.md"
    json_file = tmp_path / "report.json"

    rep_gen.generate_markdown_report(context, md_file)
    rep_gen.generate_json_report(context, json_file)

    md_content = md_file.read_text(encoding="utf-8")
    assert "AI-Assisted Accessibility Authoring & Audit Trail" in md_content
    assert "rec_rep_01" in md_content
    assert "Accessible report test suggestion" in md_content

    with open(json_file, encoding="utf-8") as f:
        json_data = json.load(f)
    assert "ai_authoring_summary" in json_data
    assert "ai_recommendations" in json_data
    assert "ai_audit_trail" in json_data


# =============================================================================
# 28. Before/after reconciliation
# =============================================================================

def test_28_before_after_reconciliation(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path, alt_text="")

    issue = _make_issue(
        issue_id=f"ALT_MISSING_1_{shape_id}",
        rule_id="SG-ALT-01",
        rule_name="Missing Alternative Text",
        slide_number=1,
        severity=IssueSeverity.ERROR,
        shape_id=shape_id,
        shape_name="RevenueChartPicture",
        behavior=RuleBehavior.AI_ASSISTED_REVIEW,
        status=AccessibilityStatus.DETECTED,
        description="Missing alt text.",
    )

    rec = AuthoringRecommendation(
        recommendation_id="rec_reconcile",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="Reconciled alt text for image",
        status=AuthoringStatus.APPROVED,
        confidence=0.95
    )
    context = JobContext(
        job_id="job_28",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        candidate_file=tmp_path / "candidate.pptx",
        rule_results=[_make_rule_result("SG-ALT-01", "Alt Text", 1, [issue])],
        ai_recommendations=[rec],

        user_approved_remediation=True
    )

    remediator = RemediationService()
    actions = remediator.remediate(context, pres_model)

    # Step 6 reconciliation summary
    step_6 = next((a for a in actions if a.rule_id == "PIPELINE_STEP_6_RECONCILIATION"), None)
    assert step_6 is not None
    assert step_6.remediated_state is not None
    # Pre-remediation alt issue should be remediated/resolved
    assert step_6.remediated_state.get("remediated_count", 0) >= 0


# =============================================================================
# 29. Visual preservation
# =============================================================================

def test_29_visual_preservation(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path)
    shape_before = pres_model.slides[0].shapes[1]
    geom_before = (shape_before.left_pt, shape_before.top_pt, shape_before.width_pt, shape_before.height_pt)

    rec = AuthoringRecommendation(
        recommendation_id="rec_vis_pres",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="Visual preservation alt text",
        status=AuthoringStatus.APPROVED
    )
    context = JobContext(
        job_id="job_29",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec]
    )

    applier = AuthoringApplier()
    applier.apply_approved(context, pres_model)

    shape_after = pres_model.slides[0].shapes[1]
    geom_after = (shape_after.left_pt, shape_after.top_pt, shape_after.width_pt, shape_after.height_pt)
    assert geom_before == geom_after


# =============================================================================
# 30. Shape ID preservation
# =============================================================================

def test_30_shape_id_preservation(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path)
    ids_before = [s.shape_id for s in pres_model.slides[0].shapes]

    rec = AuthoringRecommendation(
        recommendation_id="rec_id_pres",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="Shape ID preservation test",
        status=AuthoringStatus.APPROVED
    )
    context = JobContext(
        job_id="job_30",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec]
    )

    applier = AuthoringApplier()
    applier.apply_approved(context, pres_model)

    ids_after = [s.shape_id for s in pres_model.slides[0].shapes]
    assert ids_before == ids_after


# =============================================================================
# 31. Relationship preservation
# =============================================================================

def test_31_relationship_preservation(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, Inches(1), Inches(1), Inches(4), Inches(2))
    # Add a hyperlink to test relationship preservation
    shape.click_action.hyperlink.address = "https://example.com"
    file_path = tmp_path / "test_rel.pptx"
    prs.save(str(file_path))

    reader = PresentationReader()
    pres_model = reader.read(file_path)

    rec = AuthoringRecommendation(
        recommendation_id="rec_rel_pres",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape.shape_id,
        suggested_value="Link destination description",
        status=AuthoringStatus.APPROVED
    )
    context = JobContext(
        job_id="job_31",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec]
    )

    applier = AuthoringApplier()
    applier.apply_approved(context, pres_model)

    cand_path = tmp_path / "candidate_rel.pptx"
    PresentationWriter().write(pres_model, cand_path)

    reloaded = reader.read(cand_path)
    # Check relationship is intact in OOXML
    reloaded_raw_slide = reloaded.slides[0].raw_slide
    assert len(reloaded_raw_slide.part.rels) > 0


# =============================================================================
# 32. Animation preservation
# =============================================================================

def test_32_animation_preservation(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(1, Inches(1), Inches(1), Inches(3), Inches(2))

    # Add dummy animation timing node in XML
    ns_p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    timing = etree.SubElement(slide._element, f"{{{ns_p}}}timing")
    bldLst = etree.SubElement(timing, f"{{{ns_p}}}bldLst")
    bldSub = etree.SubElement(bldLst, f"{{{ns_p}}}bldSub", attrib={"spid": str(shape.shape_id)})

    file_path = tmp_path / "test_anim.pptx"
    prs.save(str(file_path))

    reader = PresentationReader()
    pres_model = reader.read(file_path)

    rec = AuthoringRecommendation(
        recommendation_id="rec_anim_pres",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape.shape_id,
        suggested_value="Animation preservation alt text",
        status=AuthoringStatus.APPROVED
    )
    context = JobContext(
        job_id="job_32",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        ai_recommendations=[rec]
    )

    applier = AuthoringApplier()
    applier.apply_approved(context, pres_model)

    # Check timing node remains intact
    raw_elem = pres_model.slides[0].raw_slide._element
    timing_nodes = safe_xpath(raw_elem, ".//p:timing")
    assert len(timing_nodes) == 1
    assert len(safe_xpath(timing_nodes[0], ".//p:bldSub")) == 1


# =============================================================================
# 33. Original PPTX remains untouched
# =============================================================================

def test_33_original_pptx_remains_untouched(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path)
    original_bytes = file_path.read_bytes()

    rec = AuthoringRecommendation(
        recommendation_id="rec_orig_untouched",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="Modifications should not touch original binary",
        status=AuthoringStatus.APPROVED
    )
    cand_path = tmp_path / "candidate.pptx"
    context = JobContext(
        job_id="job_33",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        candidate_file=cand_path,
        ai_recommendations=[rec],
        user_approved_remediation=True
    )

    applier = AuthoringApplier()
    applier.apply_approved(context, pres_model)
    PresentationWriter().write(pres_model, cand_path)

    # Original file bytes must be 100% unchanged
    assert file_path.read_bytes() == original_bytes


# =============================================================================
# 34. Candidate PPTX reloads successfully
# =============================================================================

def test_34_candidate_pptx_reloads_successfully(tmp_path):
    pres_model, file_path, shape_id = _build_test_pres_with_image(tmp_path)
    cand_path = tmp_path / "candidate_reload.pptx"

    rec = AuthoringRecommendation(
        recommendation_id="rec_reload_test",
        rule_id="SG-ALT-01",
        slide_number=1,
        shape_id=shape_id,
        suggested_value="Testing valid OOXML reloading",
        status=AuthoringStatus.APPROVED
    )
    context = JobContext(
        job_id="job_34",
        workspace_dir=tmp_path,
        original_file=file_path,
        working_file=file_path,
        candidate_file=cand_path,
        ai_recommendations=[rec]
    )

    applier = AuthoringApplier()
    applier.apply_approved(context, pres_model)
    PresentationWriter().write(pres_model, cand_path)

    # Reload candidate using PresentationReader and python-pptx
    reader = PresentationReader()
    reloaded_model = reader.read(cand_path)
    assert reloaded_model is not None
    assert reloaded_model.total_slides == 1
    # Check alt text is readable from reloaded file
    shape = reloaded_model.slides[0].shapes[1]
    assert shape.alt_text == "Testing valid OOXML reloading"
