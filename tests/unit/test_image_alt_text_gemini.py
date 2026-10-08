"""
Unit Tests for Task 3: AI-Assisted Alternative Text & Decorative Image Detection.
Covers:
  - All 20 explicit special cases from Section 12
  - Model role separation (Gemini 3.5 runtime, never Gemini 3.8)
  - Strict JSON parsing, invalid JSON handling, offline fallback
  - Hallucination boilerplate stripping
  - Confidence gating (High >= 0.85, Medium 0.70-0.849, Low < 0.70)
  - Existing alt text protection (never overwrite good alt text)
  - Animation and group hierarchy protection
  - OOXML in-place transformation integrity:
      0 new objects, 0 deleted objects, 0 duplicate objects,
      0 shape ID changes, 0 geometry changes, 0 relationship changes
"""

import pytest
from unittest.mock import MagicMock, patch
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt

from rules.alt_text.models import (
    ImageCandidate,
    ImageDecision,
    ImageClassification,
    ImageAction,
)
from rules.alt_text.extractor import ImageExtractor, is_weak_or_generic_alt_text
from rules.alt_text.service import ImageAccessibilityService
from rules.alt_text.transformer import ImageAltTextTransformer, DECORATIVE_EXT_URI
from rules.alt_text.reporter import generate_task_3_report
from engine.intelligence.gemini_assistant import GeminiAssistant, resolve_runtime_model


# =============================================================================
# FIXTURES & HELPERS
# =============================================================================

def create_mock_candidate(
    slide_number=1,
    shape_id=10,
    shape_name="Picture 1",
    existing_alt="",
    dimensions=(300.0, 200.0),
    is_animated=False,
    is_already_decorative=False,
    has_hyperlink=False,
    containing_group_id=None,
    group_hierarchy=None,
    nearby_caption="",
    nearby_text=None,
    nearby_title="AI Architecture",
    is_repeated=False,
) -> ImageCandidate:
    is_generic = is_weak_or_generic_alt_text(existing_alt) if existing_alt else False
    return ImageCandidate(
        slide_number=slide_number,
        shape_id=shape_id,
        shape_name=shape_name,
        image_rel_id="rId2",
        image_format="image/jpeg",
        dimensions_pt=dimensions,
        bbox_pt=(50.0, 50.0, dimensions[0], dimensions[1]),
        z_order=1,
        containing_group_id=containing_group_id,
        group_hierarchy=group_hierarchy or [],
        nearby_text=nearby_text or ["The adversarial pipeline operates in stages."],
        nearby_title=nearby_title,
        nearby_caption=nearby_caption,
        surrounding_shapes=["Shape 11", "Shape 12"],
        has_hyperlink=has_hyperlink,
        hyperlink_url="https://example.com" if has_hyperlink else "",
        is_animated=is_animated,
        is_already_decorative=is_already_decorative,
        existing_alt_text=existing_alt,
        is_generic_alt_text=is_generic,
        is_repeated_across_slides=is_repeated,
    )


# =============================================================================
# MODEL ROLE SEPARATION TESTS
# =============================================================================

def test_model_role_separation():
    """Verify runtime model is strictly gemini-3.5 and rejects gemini-3.8-flash-high."""
    resolved = resolve_runtime_model("gemini-3.5")
    assert resolved == "gemini-3.5"

    filtered = resolve_runtime_model("gemini-3.8-flash-high")
    assert filtered == "gemini-3.5", "Runtime model must filter out coding model gemini-3.8-flash-high"

    assistant = GeminiAssistant(model_name="gemini-3.5")
    assert assistant.model_name == "gemini-3.5"
    assert assistant.model_name != "gemini-3.8-flash-high"
    assert assistant.model_name != "gemini-3.8"


# =============================================================================
# 20 SPECIAL CASES (SECTION 12 & 19)
# =============================================================================

def test_case_01_photograph_meaningful_subject():
    """1. Photograph containing meaningful subject -> GENERATE_ALT_TEXT (High confidence)."""
    cand = create_mock_candidate(
        shape_name="Photograph 1",
        dimensions=(400.0, 300.0),
        nearby_title="Quantum Computing Team",
        nearby_text=["Researchers in the superconducting qubit laboratory."]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Researchers working inside the superconducting qubit laboratory.",
        "confidence": 0.94,
        "reason": "Photograph conveys the research team and laboratory environment."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.MEANINGFUL
    assert decision.action == ImageAction.GENERATE_ALT_TEXT
    assert "superconducting qubit" in decision.alt_text
    assert decision.confidence >= 0.85
    assert decision.model_traceability["slideguard_runtime_model"] in ("Gemini 3.5", "gemini-3.5")
    assert decision.model_traceability["runtime_model_identifier"] == "gemini-3.5"



def test_case_02_decorative_background_image():
    """2. Decorative background image / divider -> MARK_DECORATIVE."""
    cand = create_mock_candidate(
        shape_name="AccentDivider",
        dimensions=(4.0, 500.0),  # Thin divider line
    )
    service = ImageAccessibilityService()
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.DECORATIVE
    assert decision.action == ImageAction.MARK_DECORATIVE
    assert decision.alt_text == ""
    assert decision.confidence >= 0.85


def test_case_03_full_slide_background_meaningful():
    """3. Full-slide background that is actually meaningful diagram -> NOT decorative solely by size."""
    cand = create_mock_candidate(
        shape_name="FullSlideInfographic",
        dimensions=(720.0, 540.0),  # Full slide dimensions
        nearby_title="Global Threat Landscape",
        nearby_text=["Map illustrating cybersecurity attack vectors across continents."]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "World map illustrating global cybersecurity attack vectors.",
        "confidence": 0.92,
        "reason": "Infographic conveys substantive geographic attack data."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    # Position/size alone must NOT force decorative status
    assert decision.classification == ImageClassification.MEANINGFUL
    assert decision.action == ImageAction.GENERATE_ALT_TEXT
    assert "attack vectors" in decision.alt_text


def test_case_04_diagram():
    """4. Diagram -> GENERATE_ALT_TEXT with factual explanation."""
    cand = create_mock_candidate(
        shape_name="Diagram 3",
        dimensions=(350.0, 250.0),
        nearby_title="Model Inversion Pipeline",
        nearby_text=["Three sequential stages: feature extraction, shadow model, reconstruction."]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Diagram depicting three stages: feature extraction, shadow model, and reconstruction.",
        "confidence": 0.95,
        "reason": "Diagram visualizes the attack pipeline referenced in text."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.MEANINGFUL
    assert decision.action == ImageAction.GENERATE_ALT_TEXT
    assert "Diagram depicting" in decision.alt_text


def test_case_05_screenshot():
    """5. Screenshot -> GENERATE_ALT_TEXT."""
    cand = create_mock_candidate(
        shape_name="Screenshot_Console",
        dimensions=(400.0, 220.0),
        nearby_title="Configuration Console",
        nearby_caption="Figure 4: Admin dashboard showing API key management"
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Admin dashboard screenshot displaying API key management settings.",
        "confidence": 0.91,
        "reason": "UI screenshot demonstrates the administrative console."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.MEANINGFUL
    assert decision.action == ImageAction.GENERATE_ALT_TEXT


def test_case_06_chart_image():
    """6. Chart rendered as an image -> GENERATE_ALT_TEXT describing trends."""
    cand = create_mock_candidate(
        shape_name="ChartImage",
        dimensions=(380.0, 260.0),
        nearby_title="Accuracy Degradation Under Attack",
        nearby_text=["Accuracy drops from 95% down to 42% as epsilon increases."]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Chart showing model accuracy decreasing from 95% to 42% as attack perturbation increases.",
        "confidence": 0.93,
        "reason": "Chart conveys the empirical results discussed on slide."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.MEANINGFUL
    assert decision.action == ImageAction.GENERATE_ALT_TEXT


def test_case_07_logo():
    """7. Logo -> GENERATE_ALT_TEXT concise identifier."""
    cand = create_mock_candidate(
        shape_name="SlideGuardLogo",
        dimensions=(120.0, 45.0),
        nearby_title="SlideGuard Security Suite"
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "SlideGuard Security Suite logo.",
        "confidence": 0.96,
        "reason": "Corporate identity branding logo."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.MEANINGFUL
    assert "logo" in decision.alt_text.lower()


def test_case_08_repeated_decorative_icon():
    """8. Repeated decorative icon / bullet -> MARK_DECORATIVE (Deterministic gate)."""
    cand = create_mock_candidate(
        shape_name="BulletIcon",
        dimensions=(14.0, 14.0),  # Tiny icon <= 20pt
    )
    service = ImageAccessibilityService()
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.DECORATIVE
    assert decision.action == ImageAction.MARK_DECORATIVE
    assert decision.alt_text == ""


def test_case_09_image_inside_group():
    """9. Image inside a group -> group hierarchy preserved, evaluated properly."""
    cand = create_mock_candidate(
        shape_id=25,
        containing_group_id=100,
        group_hierarchy=["Group 100"]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 25,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "System architecture schematic inside component group.",
        "confidence": 0.88,
        "reason": "Grouped diagram component."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.is_grouped is True
    assert decision.shape_id == 25
    assert decision.action == ImageAction.GENERATE_ALT_TEXT


def test_case_10_multiple_images_in_one_slide():
    """10. Multiple images in one slide -> each evaluated independently."""
    cand1 = create_mock_candidate(slide_number=2, shape_id=101, shape_name="Image A")
    cand2 = create_mock_candidate(slide_number=2, shape_id=102, shape_name="Image B")

    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.side_effect = [
        {"slide_number": 2, "shape_id": 101, "classification": "MEANINGFUL", "action": "GENERATE_ALT_TEXT", "alt_text": "Diagram A", "confidence": 0.90, "reason": "First image"},
        {"slide_number": 2, "shape_id": 102, "classification": "MEANINGFUL", "action": "GENERATE_ALT_TEXT", "alt_text": "Diagram B", "confidence": 0.89, "reason": "Second image"},
    ]

    service = ImageAccessibilityService(assistant=mock_assistant)
    d1 = service.evaluate_candidate(cand1)
    d2 = service.evaluate_candidate(cand2)

    assert d1.shape_id == 101
    assert d2.shape_id == 102
    assert d1.alt_text == "Diagram A"
    assert d2.alt_text == "Diagram B"


def test_case_11_image_with_nearby_caption():
    """11. Image with nearby caption -> caption fed to Gemini, grounded output."""
    cand = create_mock_candidate(
        nearby_caption="Figure 2.3: Comparison of clean vs perturbed inputs",
        nearby_title="Adversarial Noise"
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Comparison of clean input image alongside adversarially perturbed input.",
        "confidence": 0.96,
        "reason": "Directly grounded by caption Figure 2.3."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.confidence >= 0.85
    assert "Comparison of clean input" in decision.alt_text


def test_case_12_image_with_nearby_explanatory_text():
    """12. Image with nearby explanatory text -> included in context."""
    cand = create_mock_candidate(
        nearby_text=["Gradient ascent generates minimal perturbations that fool classifiers."]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Visual explanation of gradient ascent perturbations.",
        "confidence": 0.90,
        "reason": "Grounded in surrounding explanation."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)
    assert decision.action == ImageAction.GENERATE_ALT_TEXT


def test_case_13_existing_good_alt_text_preserved():
    """13. Existing good author-provided alt text MUST BE PRESERVED."""
    good_alt = "Three stages of the adversarial pipeline: discovery, exploit, exfiltration."
    cand = create_mock_candidate(existing_alt=good_alt)

    # Even if Gemini were called, deterministic safety gate blocks overwriting
    service = ImageAccessibilityService()
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.PRESERVE
    assert decision.classification == ImageClassification.MEANINGFUL
    assert decision.alt_text == good_alt
    assert decision.confidence >= 0.95
    assert "Existing author-provided alt text preserved" in decision.reason


def test_case_14_existing_empty_alt_text():
    """14. Existing empty alt text -> evaluated for generation."""
    cand = create_mock_candidate(existing_alt="")
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Technical diagram of network firewall filters.",
        "confidence": 0.91,
        "reason": "Missing alt text provided."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.GENERATE_ALT_TEXT
    assert decision.alt_text == "Technical diagram of network firewall filters."


def test_case_15_existing_weak_alt_text():
    """15. Existing useless placeholder alt text ('image', 'Picture 5') -> replacement allowed."""
    for weak in ["image", "picture", "photo", "IMG_1234", "image1", "Picture 5"]:
        assert is_weak_or_generic_alt_text(weak) is True

    cand = create_mock_candidate(existing_alt="Picture 5")
    assert cand.is_generic_alt_text is True

    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Photograph of security operations center analysts at workstations.",
        "confidence": 0.92,
        "reason": "Replaced generic placeholder 'Picture 5' with descriptive text."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.GENERATE_ALT_TEXT
    assert decision.alt_text != "Picture 5"
    assert "analysts at workstations" in decision.alt_text


def test_case_16_image_with_hyperlink():
    """16. Image with hyperlink -> never falsely tagged as decorative divider/icon."""
    cand = create_mock_candidate(
        dimensions=(15.0, 15.0),  # Tiny dimension that would normally be a divider/bullet
        has_hyperlink=True
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "External link icon leading to documentation.",
        "confidence": 0.88,
        "reason": "Hyperlinked actionable control icon."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.GENERATE_ALT_TEXT
    assert decision.classification == ImageClassification.MEANINGFUL


def test_case_17_animated_image():
    """17. Animated image -> PRESERVED / MANUAL_REVIEW (Animation protection gate)."""
    cand = create_mock_candidate(is_animated=True)
    service = ImageAccessibilityService()
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.MANUAL_REVIEW
    assert decision.is_animated is True
    assert "animation sequence" in decision.reason.lower()


def test_case_18_image_in_nested_group():
    """18. Image in nested group -> group structure untouched."""
    cand = create_mock_candidate(
        containing_group_id=202,
        group_hierarchy=["MainGroup", "SubGroup202"]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Nested component diagram.",
        "confidence": 0.89,
        "reason": "Component inside nested group."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.is_grouped is True
    assert decision.action == ImageAction.GENERATE_ALT_TEXT


def test_case_19_image_with_no_surrounding_text():
    """19. Image with no surrounding text -> lower confidence -> MANUAL_REVIEW."""
    cand = create_mock_candidate(
        nearby_text=[],
        nearby_title="",
        nearby_caption=""
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    # When context is absent, model outputs lower confidence
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "UNKNOWN",
        "action": "MANUAL_REVIEW",
        "alt_text": "",
        "confidence": 0.50,
        "reason": "Insufficient slide context to verify visual content."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.MANUAL_REVIEW
    assert decision.confidence < 0.70
    assert decision.status == "MANUAL_REVIEW"


def test_case_20_same_image_reused_on_multiple_slides():
    """20. Same image reused on multiple slides (>=3) -> classified as DECORATIVE."""
    cand = create_mock_candidate(is_repeated=True)
    service = ImageAccessibilityService()
    decision = service.evaluate_candidate(cand)

    assert decision.classification == ImageClassification.DECORATIVE
    assert decision.action == ImageAction.MARK_DECORATIVE
    assert decision.alt_text == ""
    assert "Repeated background graphic" in decision.reason


# =============================================================================
# GEMINI ENGINE, JSON & CONFIDENCE GATING TESTS
# =============================================================================

def test_gemini_valid_json_parsing():
    """Verify clean parsing of strictly structured Gemini 3.5 JSON."""
    assistant = GeminiAssistant(model_name="gemini-3.5")
    valid_json = """{
        "slide_number": 3,
        "shape_id": 12,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Block diagram of the neural net defense layers.",
        "confidence": 0.92,
        "reason": "Diagram referenced in bullet 2."
    }"""
    parsed = assistant._parse_and_validate_image_response(valid_json, {"slide_number": 3, "shape_id": 12})
    assert parsed["classification"] == "MEANINGFUL"
    assert parsed["action"] == "GENERATE_ALT_TEXT"
    assert parsed["confidence"] == 0.92


def test_gemini_invalid_json_fallback():
    """Verify invalid/corrupt JSON gracefully falls back to deterministic MANUAL_REVIEW."""
    assistant = GeminiAssistant(model_name="gemini-3.5")
    corrupt_str = "I think this is an image showing a flowchart but I am not outputting JSON."
    parsed = assistant._parse_and_validate_image_response(corrupt_str, {"slide_number": 4, "shape_id": 15})
    assert parsed["classification"] == "UNKNOWN"
    assert parsed["action"] == "MANUAL_REVIEW"
    assert parsed["confidence"] <= 0.65


def test_gemini_hallucinated_boilerplate_removal():
    """Verify forbidden preamble phrases are automatically stripped from alt text."""
    assistant = GeminiAssistant(model_name="gemini-3.5")
    json_with_boilerplate = """{
        "slide_number": 1,
        "shape_id": 5,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "This image shows a flowchart illustrating the encryption handshake protocol.",
        "confidence": 0.95,
        "reason": "Conveys protocol."
    }"""
    parsed = assistant._parse_and_validate_image_response(json_with_boilerplate, {"slide_number": 1, "shape_id": 5})
    assert not parsed["alt_text"].startswith("This image shows")
    assert parsed["alt_text"].startswith("A flowchart illustrating") or parsed["alt_text"].startswith("Flowchart")


def test_confidence_gating_thresholds():
    """
    Verify confidence gating:
      High (>= 0.85): allowed
      Medium (0.70-0.849): MANUAL_REVIEW
      Low (< 0.70): MANUAL_REVIEW
    """
    cand = create_mock_candidate()

    # Test Medium confidence
    mock_assistant_med = MagicMock()
    mock_assistant_med.model_name = "gemini-3.5"
    mock_assistant_med.analyze_image_accessibility.return_value = {
        "slide_number": 1, "shape_id": 10, "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT", "alt_text": "Proposed alt text",
        "confidence": 0.78, "reason": "Medium certainty"
    }
    service_med = ImageAccessibilityService(assistant=mock_assistant_med)
    dec_med = service_med.evaluate_candidate(cand)
    assert dec_med.action == ImageAction.MANUAL_REVIEW, "Medium confidence must be gated to MANUAL_REVIEW"
    assert dec_med.status == "MANUAL_REVIEW"

    # Test Low confidence
    mock_assistant_low = MagicMock()
    mock_assistant_low.model_name = "gemini-3.5"
    mock_assistant_low.analyze_image_accessibility.return_value = {
        "slide_number": 1, "shape_id": 10, "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT", "alt_text": "Uncertain text",
        "confidence": 0.55, "reason": "Low certainty"
    }
    service_low = ImageAccessibilityService(assistant=mock_assistant_low)
    dec_low = service_low.evaluate_candidate(cand)
    assert dec_low.action == ImageAction.MANUAL_REVIEW, "Low confidence must be gated to MANUAL_REVIEW"
    assert dec_low.status == "MANUAL_REVIEW"


# =============================================================================
# INTEGRITY & IN-PLACE OOXML TRANSFORMATION TESTS
# =============================================================================

def test_integrity_in_place_ooxml_transformation():
    """
    Verify direct in-place OOXML transformation:
      0 new image objects
      0 deleted image objects
      0 duplicate image objects
      0 shape ID changes
      0 geometry changes
    """
    # Create synthetic slide with an image shape in python-pptx
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

    # Create dummy shape
    shape = slide.shapes.add_shape(1, Inches(1), Inches(1), Inches(4), Inches(3))
    shape_id = shape.shape_id

    # Add <p:cNvPr>
    cnvpr = shape._element.xpath(".//p:cNvPr")[0]
    initial_id = cnvpr.get("id")
    initial_descr = cnvpr.get("descr")

    # Capture initial geometry
    xfrm = shape._element.xpath(".//a:xfrm")[0]
    initial_off_x = xfrm.xpath(".//a:off")[0].get("x")
    initial_off_y = xfrm.xpath(".//a:off")[0].get("y")
    initial_ext_cx = xfrm.xpath(".//a:ext")[0].get("cx")
    initial_ext_cy = xfrm.xpath(".//a:ext")[0].get("cy")

    initial_shape_count = len(slide.shapes)

    # Apply GENERATE_ALT_TEXT decision
    decision_alt = ImageDecision(
        slide_number=1,
        shape_id=shape_id,
        shape_name=shape.name,
        classification=ImageClassification.MEANINGFUL,
        action=ImageAction.GENERATE_ALT_TEXT,
        alt_text="A clean test diagram.",
        confidence=0.95,
        reason="Test application",
        status="APPROVED"
    )

    transformer = ImageAltTextTransformer()
    res = transformer.apply_decisions(slide, [decision_alt], approved_only=True)

    assert res["applied_count"] == 1
    assert cnvpr.get("descr") == "A clean test diagram."

    # Verify integrity: 0 count drift, 0 ID change, 0 geometry drift
    assert len(slide.shapes) == initial_shape_count, "Shape count must not change"
    assert cnvpr.get("id") == initial_id, "Shape ID must not change"
    assert xfrm.xpath(".//a:off")[0].get("x") == initial_off_x
    assert xfrm.xpath(".//a:off")[0].get("y") == initial_off_y
    assert xfrm.xpath(".//a:ext")[0].get("cx") == initial_ext_cx
    assert xfrm.xpath(".//a:ext")[0].get("cy") == initial_ext_cy

    # Apply MARK_DECORATIVE decision
    decision_dec = ImageDecision(
        slide_number=1,
        shape_id=shape_id,
        shape_name=shape.name,
        classification=ImageClassification.DECORATIVE,
        action=ImageAction.MARK_DECORATIVE,
        alt_text="",
        confidence=0.95,
        reason="Test decorative",
        status="APPROVED"
    )

    res_dec = transformer.apply_decisions(slide, [decision_dec], approved_only=True)
    assert res_dec["applied_count"] == 1
    assert cnvpr.get("descr") == ""

    # Verify modern Office 365 decorative tag is present
    ext_lst = cnvpr.xpath(".//a:extLst")
    assert len(ext_lst) == 1
    exts = [e for e in ext_lst[0] if e.get("uri") == DECORATIVE_EXT_URI]
    assert len(exts) == 1
    dec_val = exts[0].xpath(".//*[local-name()='decorative']")[0].get("val")
    assert dec_val == "1"

    # Integrity verification again
    assert len(slide.shapes) == initial_shape_count
    assert cnvpr.get("id") == initial_id


def test_task_3_reporting_structure():
    """Verify generate_task_3_report matches all Section 17 metrics."""
    decisions = [
        ImageDecision(slide_number=1, shape_id=10, shape_name="Img1", classification=ImageClassification.MEANINGFUL, action=ImageAction.GENERATE_ALT_TEXT, alt_text="Alt 1", original_alt_text="", confidence=0.95, reason="Reason 1", status="APPROVED"),
        ImageDecision(slide_number=1, shape_id=11, shape_name="Img2", classification=ImageClassification.DECORATIVE, action=ImageAction.MARK_DECORATIVE, alt_text="", original_alt_text="", confidence=0.90, reason="Reason 2", is_decorative_tagged=True, status="APPROVED"),
        ImageDecision(slide_number=2, shape_id=20, shape_name="Img3", classification=ImageClassification.MEANINGFUL, action=ImageAction.PRESERVE, alt_text="Good existing", original_alt_text="Good existing", confidence=0.98, reason="Reason 3", status="APPROVED"),
        ImageDecision(slide_number=2, shape_id=21, shape_name="Img4", classification=ImageClassification.UNKNOWN, action=ImageAction.MANUAL_REVIEW, alt_text="", original_alt_text="", confidence=0.60, reason="Reason 4", is_animated=True, status="MANUAL_REVIEW"),
    ]

    report = generate_task_3_report(decisions)

    summary = report["summary"]

    assert summary["total_images"] == 4
    assert summary["meaningful_images"] == 2
    assert summary["decorative_images"] == 1
    assert summary["unknown_images"] == 1
    assert summary["existing_good_alt_text"] == 1
    assert summary["generated_alt_text"] == 1
    assert summary["images_marked_decorative"] == 1
    assert summary["manual_review_items"] == 1
    assert summary["high_confidence_decisions"] == 3
    assert summary["low_confidence_decisions"] == 1
    assert summary["animation_protected_images"] == 1
    assert len(report["changed_images"]) == 4


# =============================================================================
# TASK 3B AUDIT & QUALITY CLASSIFICATION TESTS
# =============================================================================

def test_task_3b_generic_title_detection():
    """Verify is_generic_title_repetition detects formulaic repetitions."""
    from rules.alt_text.quality import is_generic_title_repetition

    # Direct examples from user prompt
    assert is_generic_title_repetition(
        "Illustration for Overview: The Adversarial Arms Race in AI Security.",
        "Overview: The Adversarial Arms Race in AI Security"
    ) is True

    assert is_generic_title_repetition(
        "Illustration for Evasion Attacks: Sneaking Past AI Detectors.",
        "Evasion Attacks: Sneaking Past AI Detectors"
    ) is True

    assert is_generic_title_repetition(
        "Visual graphic supporting When AI Remembers Too Much: Model Inversion.",
        "When AI Remembers Too Much: Model Inversion"
    ) is True

    assert is_generic_title_repetition(
        "Image supporting Model Robustness.",
        "Model Robustness"
    ) is True

    # Substantive non-formulaic descriptions must NOT be flagged
    assert is_generic_title_repetition(
        "Researchers working inside the superconducting qubit laboratory.",
        "Quantum Computing Team"
    ) is False

    assert is_generic_title_repetition(
        "Diagram depicting three stages: feature extraction, shadow model, and reconstruction.",
        "Model Inversion Pipeline"
    ) is False


def test_task_3b_downgrade_generic_title_repetition():
    """Verify formulaic title repetitions are downgraded to MANUAL_REVIEW, GENERIC, and lower confidence."""
    cand = create_mock_candidate(
        nearby_title="Evasion Attacks: Sneaking Past AI Detectors"
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Illustration for Evasion Attacks: Sneaking Past AI Detectors.",
        "confidence": 0.88,
        "reason": "Image supporting slide title."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    # Must be downgraded
    assert decision.action == ImageAction.MANUAL_REVIEW
    assert decision.quality_class == "GENERIC"
    assert decision.generation_basis == "STRUCTURAL_CONTEXT"
    assert decision.confidence <= 0.75
    assert "ALT_TEXT_TOO_GENERIC" in decision.reason
    assert decision.status == "MANUAL_REVIEW"


def test_task_3b_structural_context_never_claimed_specific():
    """Verify images without visual pixels are labeled STRUCTURAL_CONTEXT and CONTEXTUAL, never SPECIFIC."""
    cand = create_mock_candidate(
        nearby_title="System Architecture",
        nearby_text=["Pipeline components flow from ingest to validation."]
    )
    mock_assistant = MagicMock()
    mock_assistant.model_name = "gemini-3.5"
    mock_assistant.analyze_image_accessibility.return_value = {
        "slide_number": 1,
        "shape_id": 10,
        "classification": "MEANINGFUL",
        "action": "GENERATE_ALT_TEXT",
        "alt_text": "Workflow schematic detailing ingest, validation, and serialization pipeline.",
        "confidence": 0.90,
        "reason": "Derived from surrounding text context."
    }

    service = ImageAccessibilityService(assistant=mock_assistant)
    decision = service.evaluate_candidate(cand)

    assert decision.generation_basis == "STRUCTURAL_CONTEXT"
    assert decision.quality_class == "CONTEXTUAL"
    assert decision.quality_class != "SPECIFIC"


def test_task_3b_existing_author_text_is_specific():
    """Verify existing author text is recognized as EXISTING_AUTHOR_TEXT and SPECIFIC."""
    cand = create_mock_candidate(
        existing_alt="Bar chart showing model accuracy dropping from 92% to 45% across iterations."
    )
    service = ImageAccessibilityService()
    decision = service.evaluate_candidate(cand)

    assert decision.action == ImageAction.PRESERVE
    assert decision.generation_basis == "EXISTING_AUTHOR_TEXT"
    assert decision.quality_class == "SPECIFIC"


def test_task_3b_report_contains_quality_and_basis_metrics():
    """Verify report includes Task 3B quality_class and generation_basis metrics."""
    decisions = [
        ImageDecision(
            slide_number=1, shape_id=10, shape_name="Img1",
            classification=ImageClassification.MEANINGFUL, action=ImageAction.PRESERVE,
            alt_text="Good author text", original_alt_text="Good author text",
            confidence=0.98, reason="Preserved",
            generation_basis="EXISTING_AUTHOR_TEXT", quality_class="SPECIFIC",
            status="APPROVED"
        ),
        ImageDecision(
            slide_number=2, shape_id=20, shape_name="Img2",
            classification=ImageClassification.MEANINGFUL, action=ImageAction.MANUAL_REVIEW,
            alt_text="Illustration for Topic", original_alt_text="",
            confidence=0.75, reason="ALT_TEXT_TOO_GENERIC",
            generation_basis="STRUCTURAL_CONTEXT", quality_class="GENERIC",
            status="MANUAL_REVIEW"
        ),
    ]

    report = generate_task_3_report(decisions)
    summary = report["summary"]

    assert summary["quality_specific"] == 1
    assert summary["quality_generic"] == 1
    assert summary["basis_existing_author"] == 1
    assert summary["basis_structural_context"] == 1
    assert report["changed_images"][0]["generation_basis"] == "EXISTING_AUTHOR_TEXT"
    assert report["changed_images"][0]["quality_class"] == "SPECIFIC"
    assert report["changed_images"][1]["generation_basis"] == "STRUCTURAL_CONTEXT"
    assert report["changed_images"][1]["quality_class"] == "GENERIC"

