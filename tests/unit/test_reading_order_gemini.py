"""
Comprehensive unit and integration tests for SlideGuard Reading Order engine
enhanced with Gemini 3.8 Flash High semantic reasoning.

Covers all 10 required test scenarios:
1. Normal title/body slide
2. Two-column slide
3. Repeated cards (card-by-card left-to-right / top-to-bottom sequence)
4. Image + caption
5. Footer exclusion / placement
6. Decorative background exclusion
7. Grouped objects
8. Ambiguous slide (medium confidence / manual review)
9. Low-confidence slide preservation (< 0.70)
10. Animation preservation (timing target IDs untouched)
"""

import json
from pathlib import Path
import pytest
from unittest.mock import MagicMock
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

from rules.reading_order import (
    ReadingOrderRule,
    ReadingOrderResolver,
    ReadingOrderResult,
    SemanticRoleClassifier,
    SemanticRole,
    SemanticItem,
    XMLReadingOrderReorderer,
    LayoutRegionAnalyzer
)
from engine.intelligence.gemini_assistant import GeminiAssistant
from pptx_engine.reader.presentation_reader import PresentationReader
from models.enums import RemediationStatus


@pytest.fixture
def test_pptx_suite(tmp_path) -> Path:
    """Generates a PPTX with targeted slides covering scenarios 1 to 10."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SLIDE 1: Normal Title / Body Slide (Body placed first in XML, Title second) ---
    s1 = prs.slides.add_slide(blank_layout)
    body1 = s1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(4))
    body1.text_frame.text = "This is the main body paragraph describing the project overview."
    title1 = s1.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title1.text_frame.text = "Slide 1: Normal Title"

    # --- SLIDE 2: Two-Column Slide ---
    s2 = prs.slides.add_slide(blank_layout)
    title2 = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title2.text_frame.text = "Two-Column Architecture Comparison"
    # Left Column
    c1_top = s2.shapes.add_textbox(Inches(1), Inches(2.0), Inches(3.8), Inches(2))
    c1_top.text_frame.text = "Left Column Section 1"
    c1_bot = s2.shapes.add_textbox(Inches(1), Inches(4.5), Inches(3.8), Inches(2))
    c1_bot.text_frame.text = "Left Column Section 2"
    # Right Column
    c2_top = s2.shapes.add_textbox(Inches(5.2), Inches(2.0), Inches(3.8), Inches(2))
    c2_top.text_frame.text = "Right Column Section 1"
    c2_bot = s2.shapes.add_textbox(Inches(5.2), Inches(4.5), Inches(3.8), Inches(2))
    c2_bot.text_frame.text = "Right Column Section 2"

    # --- SLIDE 3: Repeated Cards Slide (3 cards horizontally: badge + title + body) ---
    s3 = prs.slides.add_slide(blank_layout)
    title3 = s3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title3.text_frame.text = "Three Core Pillars"
    # Card 1 (left=1.0)
    b1 = s3.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(0.8), Inches(0.8))
    b1.text_frame.text = "01"
    h1 = s3.shapes.add_textbox(Inches(1.0), Inches(2.9), Inches(2.4), Inches(0.8))
    h1.text_frame.text = "First Pillar Header"
    d1 = s3.shapes.add_textbox(Inches(1.0), Inches(3.8), Inches(2.4), Inches(2.5))
    d1.text_frame.text = "Detailed text description for pillar one."
    # Card 2 (left=3.8)
    b2 = s3.shapes.add_textbox(Inches(3.8), Inches(2.0), Inches(0.8), Inches(0.8))
    b2.text_frame.text = "02"
    h2 = s3.shapes.add_textbox(Inches(3.8), Inches(2.9), Inches(2.4), Inches(0.8))
    h2.text_frame.text = "Second Pillar Header"
    d2 = s3.shapes.add_textbox(Inches(3.8), Inches(3.8), Inches(2.4), Inches(2.5))
    d2.text_frame.text = "Detailed text description for pillar two."
    # Card 3 (left=6.6)
    b3 = s3.shapes.add_textbox(Inches(6.6), Inches(2.0), Inches(0.8), Inches(0.8))
    b3.text_frame.text = "03"
    h3 = s3.shapes.add_textbox(Inches(6.6), Inches(2.9), Inches(2.4), Inches(0.8))
    h3.text_frame.text = "Third Pillar Header"
    d3 = s3.shapes.add_textbox(Inches(6.6), Inches(3.8), Inches(2.4), Inches(2.5))
    d3.text_frame.text = "Detailed text description for pillar three."

    # --- SLIDE 4: Image + Caption Slide ---
    s4 = prs.slides.add_slide(blank_layout)
    title4 = s4.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title4.text_frame.text = "System Architecture Diagram"
    fig4 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2), Inches(2.0), Inches(6), Inches(3.0))
    fig4.name = "Architecture_Figure"
    cap4 = s4.shapes.add_textbox(Inches(2), Inches(5.1), Inches(6), Inches(0.5))
    cap4.text_frame.text = "Figure 1: End-to-end processing pipeline overview."

    # --- SLIDE 5: Footer & Legitimate Bottom Content ---
    s5 = prs.slides.add_slide(blank_layout)
    title5 = s5.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title5.text_frame.text = "Financial Summary and Notes"
    bot_content = s5.shapes.add_textbox(Inches(1), Inches(5.5), Inches(8), Inches(1.0))
    bot_content.text_frame.text = "Important calculation notes located near the bottom of slide."
    footer5 = s5.shapes.add_textbox(Inches(1), Inches(6.8), Inches(3), Inches(0.4))
    footer5.text_frame.text = "WILEY © 2026 All Rights Reserved"

    # --- SLIDE 6: Decorative Background Slide ---
    s6 = prs.slides.add_slide(blank_layout)
    bg6 = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(10), Inches(7.5))
    bg6.name = "Background_Rect"
    div6 = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(9), Pt(3))
    div6.name = "Divider_Line"
    title6 = s6.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title6.text_frame.text = "Slide with Decorative Elements"
    body6 = s6.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(4))
    body6.text_frame.text = "Content sitting safely on top of background rectangle."

    # --- SLIDE 7: Grouped Objects Slide ---
    s7 = prs.slides.add_slide(blank_layout)
    title7 = s7.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title7.text_frame.text = "Slide with Grouped Visuals"
    grp = s7.shapes.add_group_shape()
    grp.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(2), Inches(1), Inches(1))
    grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3), Inches(2), Inches(2), Inches(1))

    # --- SLIDE 8: Animation Target Slide ---
    s8 = prs.slides.add_slide(blank_layout)
    title8 = s8.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title8.text_frame.text = "Animated Slide"
    body8 = s8.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(4))
    body8.text_frame.text = "Target animated shape content."
    # Inject a timing XML element into s8
    timing_xml = f"""
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:tnLst>
            <p:par>
                <p:cTn id="1">
                    <p:childTnLst>
                        <p:set>
                            <p:cBhvr>
                                <p:cTn id="2"/>
                                <p:tgtEl>
                                    <p:spTgt spid="{body8.shape_id}"/>
                                </p:tgtEl>
                            </p:cBhvr>
                        </p:set>
                    </p:childTnLst>
                </p:cTn>
            </p:par>
        </p:tnLst>
    </p:timing>
    """
    s8._element.append(parse_xml(timing_xml))

    out = tmp_path / "comprehensive_reading_order_suite.pptx"
    prs.save(str(out))
    return out


# -----------------------------------------------------------------------------
# Test 1: Normal Title / Body Slide Reordering
# -----------------------------------------------------------------------------
def test_normal_title_body_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s1 = pres.slides[0]
    rule = ReadingOrderRule()

    issues = rule.evaluate(s1, pres)
    assert len(issues) == 1
    iss = issues[0]
    extra = iss.extra_data
    assert extra["confidence"] in ("HIGH", "MEDIUM")

    # Body was placed first in XML, title second. Proposed must put title first!
    body_id = s1.shapes[0].shape_id
    title_id = s1.shapes[1].shape_id
    assert extra["original_order"] == [body_id, title_id]
    assert extra["proposed_order"] == [title_id, body_id]

    action = rule.remediate_issue(s1, pres, iss)
    assert action.status == RemediationStatus.APPLIED
    assert action.remediated_state["order"] == [title_id, body_id]


# -----------------------------------------------------------------------------
# Test 2: Two-Column Slide Reading Order
# -----------------------------------------------------------------------------
def test_two_column_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s2 = pres.slides[1]
    rule = ReadingOrderRule()

    issues = rule.evaluate(s2, pres)
    res = rule.resolver.resolve_reading_order(
        slide_number=2,
        items=[rule.classifier.classify_shape(sh._element, sh, pres.slide_height_pt, pres.slide_width_pt, i) for i, sh in enumerate(s2.raw_slide.shapes)],
        slide_w_pt=pres.slide_width_pt,
        slide_h_pt=pres.slide_height_pt
    )

    # Title first, then left column, then right column
    title_item = [sh for sh in s2.shapes if "Two-Column" in sh.text_content][0]
    left_items = [sh for sh in s2.shapes if "Left" in sh.text_content]
    right_items = [sh for sh in s2.shapes if "Right" in sh.text_content]

    # Title is first
    assert res.proposed_order[0] == title_item.shape_id
    # Left column items appear before right column items
    left_positions = [res.proposed_order.index(it.shape_id) for it in left_items]
    right_positions = [res.proposed_order.index(it.shape_id) for it in right_items]
    assert max(left_positions) < min(right_positions)
    assert res.confidence == "HIGH"


# -----------------------------------------------------------------------------
# Test 3: Repeated Cards (Card-by-card sequence, NO zig-zag across columns)
# -----------------------------------------------------------------------------
def test_repeated_cards_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s3 = pres.slides[2]
    rule = ReadingOrderRule()

    raw_shapes = list(s3.raw_slide.shapes)
    items = [rule.classifier.classify_shape(sh._element, sh, pres.slide_height_pt, pres.slide_width_pt, i) for i, sh in enumerate(raw_shapes)]
    res = rule.resolver.resolve_reading_order(
        slide_number=3,
        items=items,
        slide_w_pt=pres.slide_width_pt,
        slide_h_pt=pres.slide_height_pt
    )

    assert res.confidence == "HIGH"
    prop = res.proposed_order

    # Find IDs for Badge 01, Header 1, Body 1
    b1_id = next(s.shape_id for s in s3.shapes if s.text_content.strip() == "01")
    h1_id = next(s.shape_id for s in s3.shapes if "First Pillar" in s.text_content)
    d1_id = next(s.shape_id for s in s3.shapes if "pillar one" in s.text_content)

    b2_id = next(s.shape_id for s in s3.shapes if s.text_content.strip() == "02")
    h2_id = next(s.shape_id for s in s3.shapes if "Second Pillar" in s.text_content)
    d2_id = next(s.shape_id for s in s3.shapes if "pillar two" in s.text_content)

    b3_id = next(s.shape_id for s in s3.shapes if s.text_content.strip() == "03")
    h3_id = next(s.shape_id for s in s3.shapes if "Third Pillar" in s.text_content)
    d3_id = next(s.shape_id for s in s3.shapes if "pillar three" in s.text_content)

    # Within Card 1: Badge 01 -> Header 1 -> Body 1
    assert prop.index(b1_id) < prop.index(h1_id) < prop.index(d1_id)
    # Card 1 completes BEFORE Card 2 starts!
    assert prop.index(d1_id) < prop.index(b2_id)
    # Within Card 2: Badge 02 -> Header 2 -> Body 2
    assert prop.index(b2_id) < prop.index(h2_id) < prop.index(d2_id)
    # Card 2 completes BEFORE Card 3 starts!
    assert prop.index(d2_id) < prop.index(b3_id)
    # Within Card 3: Badge 03 -> Header 3 -> Body 3
    assert prop.index(b3_id) < prop.index(h3_id) < prop.index(d3_id)


# -----------------------------------------------------------------------------
# Test 4: Image + Caption Slide
# -----------------------------------------------------------------------------
def test_image_plus_caption_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s4 = pres.slides[3]
    rule = ReadingOrderRule()

    raw_shapes = list(s4.raw_slide.shapes)
    items = [rule.classifier.classify_shape(sh._element, sh, pres.slide_height_pt, pres.slide_width_pt, i) for i, sh in enumerate(raw_shapes)]
    res = rule.resolver.resolve_reading_order(
        slide_number=4,
        items=items,
        slide_w_pt=pres.slide_width_pt,
        slide_h_pt=pres.slide_height_pt
    )

    fig_id = next(s.shape_id for s in s4.shapes if "Figure" in s.name or "Architecture" in s.name)
    cap_id = next(s.shape_id for s in s4.shapes if "Figure 1:" in s.text_content)

    prop = res.proposed_order
    # Caption appears immediately after figure (or within close proximity)
    assert prop.index(fig_id) < prop.index(cap_id)


# -----------------------------------------------------------------------------
# Test 5: Footer vs Bottom Content Slide
# -----------------------------------------------------------------------------
def test_footer_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s5 = pres.slides[4]
    rule = ReadingOrderRule()

    raw_shapes = list(s5.raw_slide.shapes)
    items = [rule.classifier.classify_shape(sh._element, sh, pres.slide_height_pt, pres.slide_width_pt, i) for i, sh in enumerate(raw_shapes)]
    res = rule.resolver.resolve_reading_order(
        slide_number=5,
        items=items,
        slide_w_pt=pres.slide_width_pt,
        slide_h_pt=pres.slide_height_pt
    )

    wiley_id = next(s.shape_id for s in s5.shapes if "WILEY" in s.text_content)
    notes_id = next(s.shape_id for s in s5.shapes if "Important calculation" in s.text_content)

    prop = res.proposed_order
    # Legitimate notes content is read BEFORE the copyright footer
    assert prop.index(notes_id) < prop.index(wiley_id)
    # Footer is at the very end of content sequence
    assert prop[-1] == wiley_id


# -----------------------------------------------------------------------------
# Test 6: Decorative Background Slide (EXCLUDE_DECORATIVE at start of spTree)
# -----------------------------------------------------------------------------
def test_decorative_background_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s6 = pres.slides[5]
    rule = ReadingOrderRule()

    raw_shapes = list(s6.raw_slide.shapes)
    items = [rule.classifier.classify_shape(sh._element, sh, pres.slide_height_pt, pres.slide_width_pt, i) for i, sh in enumerate(raw_shapes)]
    res = rule.resolver.resolve_reading_order(
        slide_number=6,
        items=items,
        slide_w_pt=pres.slide_width_pt,
        slide_h_pt=pres.slide_height_pt
    )

    bg_id = next(s.shape_id for s in s6.shapes if "Background" in s.name)
    title_id = next(s.shape_id for s in s6.shapes if "Slide with Decorative" in s.text_content)
    body_id = next(s.shape_id for s in s6.shapes if "Content sitting safely" in s.text_content)

    prop = res.proposed_order
    # Background rectangle stays at beginning of XML (background z-order)
    assert prop.index(bg_id) < prop.index(title_id)
    # Title comes before body in reading order
    assert prop.index(title_id) < prop.index(body_id)

    # Check decisions: background is EXCLUDE_DECORATIVE
    bg_dec = next(d for d in res.gemini_decisions if d["shape_id"] == bg_id)
    assert bg_dec["action"] == "EXCLUDE_DECORATIVE"
    assert bg_dec["reading_order"] == 0


# -----------------------------------------------------------------------------
# Test 7: Grouped Objects Slide
# -----------------------------------------------------------------------------
def test_grouped_objects_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s7 = pres.slides[6]
    rule = ReadingOrderRule()

    raw_shapes = list(s7.raw_slide.shapes)
    items = [rule.classifier.classify_shape(sh._element, sh, pres.slide_height_pt, pres.slide_width_pt, i) for i, sh in enumerate(raw_shapes)]
    res = rule.resolver.resolve_reading_order(
        slide_number=7,
        items=items,
        slide_w_pt=pres.slide_width_pt,
        slide_h_pt=pres.slide_height_pt
    )

    assert len(res.proposed_order) == len(res.original_order)
    assert set(res.proposed_order) == set(res.original_order)


# -----------------------------------------------------------------------------
# Test 8: Ambiguous Slide (Medium Confidence Handling)
# -----------------------------------------------------------------------------
def test_ambiguous_slide_manual_review():
    mock_assistant = MagicMock()
    mock_assistant.analyze_reading_order.return_value = {
        "slide_number": 1,
        "confidence": 0.72,  # Medium confidence
        "rationale": "Ambiguous overlapping layout with competing headings.",
        "decisions": [
            {"shape_id": 1, "semantic_role": "TITLE", "reading_order": 2, "confidence": 0.72, "action": "REORDER", "reason": "Probable title"},
            {"shape_id": 2, "semantic_role": "BODY", "reading_order": 1, "confidence": 0.70, "action": "REORDER", "reason": "Ambiguous body"}
        ],

        "warnings": ["Competing headings detected"]
    }

    mock_layout = MagicMock()
    mock_layout.analyze_layout.return_value = MagicMock(header_items=[], body_columns=[1, 2, 3, 4])  # Not strongly confirmed

    resolver = ReadingOrderResolver(layout_analyzer=mock_layout, assistant=mock_assistant)
    items = [
        SemanticItem(shape_id=1, name="Box1", role=SemanticRole.UNKNOWN, confidence=0.7, rationale="", is_footer=False, is_title=False, bbox_pt=(0, 0, 100, 100), z_order=0),
        SemanticItem(shape_id=2, name="Box2", role=SemanticRole.UNKNOWN, confidence=0.7, rationale="", is_footer=False, is_title=False, bbox_pt=(50, 50, 100, 100), z_order=1),
    ]

    res = resolver.resolve_reading_order(1, items, 720.0, 540.0)
    assert res.confidence == "MEDIUM"
    # Medium confidence without strong confirmation preserves original order
    assert res.changes_required is False
    assert res.proposed_order == res.original_order


# -----------------------------------------------------------------------------
# Test 9: Low-Confidence Slide Preservation (< 0.70)
# -----------------------------------------------------------------------------
def test_low_confidence_slide_preservation():
    mock_assistant = MagicMock()
    mock_assistant.analyze_reading_order.return_value = {
        "slide_number": 1,
        "confidence": 0.50,  # LOW confidence
        "rationale": "Layout completely irregular and overlapping.",
        "decisions": [
            {"shape_id": 1, "semantic_role": "UNKNOWN", "reading_order": 2, "confidence": 0.5, "action": "MANUAL_REVIEW", "reason": "Unclear"},
            {"shape_id": 2, "semantic_role": "UNKNOWN", "reading_order": 1, "confidence": 0.5, "action": "MANUAL_REVIEW", "reason": "Unclear"}
        ],
        "warnings": ["Low confidence"]
    }

    resolver = ReadingOrderResolver(assistant=mock_assistant)
    items = [
        SemanticItem(shape_id=1, name="Box1", role=SemanticRole.UNKNOWN, confidence=0.5, rationale="", is_footer=False, is_title=False, bbox_pt=(0, 0, 100, 100), z_order=0),
        SemanticItem(shape_id=2, name="Box2", role=SemanticRole.UNKNOWN, confidence=0.5, rationale="", is_footer=False, is_title=False, bbox_pt=(0, 0, 100, 100), z_order=1),
    ]

    res = resolver.resolve_reading_order(1, items, 720.0, 540.0)
    # Hard safety rule: LOW confidence must preserve existing order!
    assert res.confidence == "LOW"
    assert res.changes_required is False
    assert res.proposed_order == [1, 2]


# -----------------------------------------------------------------------------
# Test 10: Animation-Preserved Slide (Timing target IDs untouched)
# -----------------------------------------------------------------------------
def test_animation_preserved_slide(test_pptx_suite):
    reader = PresentationReader()
    pres = reader.read(test_pptx_suite)
    s8 = pres.slides[7]
    rule = ReadingOrderRule()

    issues = rule.evaluate(s8, pres)
    # In-place XML reordering must preserve animation target references
    slide_elem = s8.raw_slide._element
    timings_before = slide_elem.xpath(".//p:timing")
    assert len(timings_before) > 0
    sp_tgt_before = slide_elem.xpath(".//p:spTgt")[0].get("spid")

    if issues:
        action = rule.remediate_issue(s8, pres, issues[0])
        assert action.status == RemediationStatus.APPLIED

    # Verify animation timing target still points to valid shape on slide
    timings_after = slide_elem.xpath(".//p:timing")
    assert len(timings_after) > 0
    sp_tgt_after = slide_elem.xpath(".//p:spTgt")[0].get("spid")
    assert sp_tgt_before == sp_tgt_after

    # Rule verification method passes
    assert rule._verify_animation_targets(slide_elem) is True


# -----------------------------------------------------------------------------
# Test 11: Gemini Structured Output & Mock Integration Test
# -----------------------------------------------------------------------------
def test_gemini_assistant_structured_call():
    """Validates that GeminiAssistant formats the prompt and parses structured decisions correctly."""
    mock_client = MagicMock()
    mock_response = MagicMock()
    mock_response.text = json.dumps({
        "slide_number": 1,
        "confidence": 0.95,
        "rationale": "Clear hierarchy: Title followed by Body.",
        "decisions": [
            {"shape_id": 10, "semantic_role": "TITLE", "reading_order": 1, "confidence": 0.95, "action": "REORDER", "reason": "Top heading"},
            {"shape_id": 20, "semantic_role": "BODY", "reading_order": 2, "confidence": 0.95, "action": "REORDER", "reason": "Content paragraph"}
        ]
    })
    mock_client.generate_content.return_value = mock_response

    assistant = GeminiAssistant(client=mock_client, model_name="gemini-3.8-flash-high")
    context = {
        "slide_number": 1,
        "shapes": [
            {"shape_id": 10, "name": "Title", "geometry": {"left_pt": 0, "top_pt": 0, "width_pt": 100, "height_pt": 50}},
            {"shape_id": 20, "name": "Body", "geometry": {"left_pt": 0, "top_pt": 60, "width_pt": 100, "height_pt": 100}}
        ]
    }

    result = assistant.analyze_reading_order(context)
    assert mock_client.generate_content.called
    assert result["confidence"] == 0.95
    assert len(result["decisions"]) == 2
    assert result["decisions"][0]["shape_id"] == 10
    assert result["decisions"][0]["action"] == "REORDER"
    assert result["decisions"][1]["shape_id"] == 20
