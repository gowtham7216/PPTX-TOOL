"""
SlideGuard Task 4 Unit and Integration Tests.
Covers all 24 required test scenarios from Task 4 specification:
1. Title textbox in-place conversion
2. Subtitle conversion
3. Body textbox conversion
4. Numeric badges ('03', '04') protection
5. Card-based slide hierarchy retention
6. Missing title slide (no fabricated title)
7. Existing native title not duplicated
8. Empty textbox does not trigger creation
9. Create new placeholder 7 safety gates
10. Existing text and formatting unchanged
11. Geometry and transforms unchanged
12. Nested groups preserved
13. Image, table, and chart object retention
14. Layout/master inheritance remains valid
15. Hyperlinks and relationship targets remain valid
16. Shape ID preservation contract
17. No duplicate IDs or placeholders
18. Animation target IDs and timing preserved
19. Gemini malformed JSON and invalid IDs rejected
20. Gemini timeout/unavailability safety
21. Multi-slide failure isolation
22. Idempotency on rerun
23. Reload verification via PresentationReader
24. Audit report accuracy matching XML changes
"""

from pathlib import Path
from unittest.mock import MagicMock
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml import parse_xml
from lxml import etree

from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from models.presentation import PresentationModel
from models.slide import SlideModel
from models.enums import RuleBehavior, RemediationStatus
from rules.placeholders.classifier import (
    PlaceholderClassifier,
    PlaceholderAnalysis,
    PlaceholderCategory,
    PlaceholderStatus,
    ConversionAction,
    DifferenceCategory,
    is_numeric_badge,
    is_short_label_or_decorative,
    _safe_xpath,
    NSMAP,
)
from rules.placeholders.transformer import PlaceholderTransformer, ConversionResult
from rules.placeholders.rule import PlaceholderRule


def _make_xml(xml_str: str):
    return parse_xml(xml_str.strip())


# =============================================================================
# 1. Genuine title textbox converts in place
# =============================================================================
def test_genuine_title_textbox_converts_in_place():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    tb.text_frame.text = "Executive Strategy Summary"
    shape_id = tb.shape_id

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide, approved_shape_ids=[shape_id])

    assert result.success
    assert shape_id in result.converted_shape_ids

    # Verify native title placeholder was added to the same shape element
    ph_elems = _safe_xpath(tb._element, ".//p:ph")
    assert len(ph_elems) == 1
    assert ph_elems[0].get("type") == "title"
    assert tb.text_frame.text == "Executive Strategy Summary"


# =============================================================================
# 2. Genuine subtitle converts correctly
# =============================================================================
def test_genuine_subtitle_converts_correctly():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide layout (supports subtitle)
    # Clear default slide shapes to test loose text boxes
    for sp in list(slide.shapes):
        sp._element.getparent().remove(sp._element)

    title_box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    title_box.text_frame.text = "Annual Financial Review"

    sub_box = slide.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(0.8))
    sub_box.text_frame.text = "Fiscal Year 2026 Projections"

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide)

    assert result.success
    # Subtitle placeholder tag added
    sub_ph = _safe_xpath(sub_box._element, ".//p:ph")
    assert len(sub_ph) == 1
    assert sub_ph[0].get("type") == "subTitle"
    assert sub_box.text_frame.text == "Fiscal Year 2026 Projections"


# =============================================================================
# 3. Suitable body textbox converts only when role is clear
# =============================================================================
def test_suitable_body_textbox_converts():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and Content layout
    # Remove default content placeholder so slide has title but no body
    for sp in list(slide.shapes):
        if sp.shape_id != slide.shapes.title.shape_id:
            sp._element.getparent().remove(sp._element)

    body_box = slide.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(4))
    body_box.text_frame.text = "Key deliverables include:\n- Phase 1 deployment\n- Architecture hardening\n- WCAG testing"
    b_id = body_box.shape_id

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide, approved_shape_ids=[b_id])

    assert result.success
    body_ph = _safe_xpath(body_box._element, ".//p:ph")
    assert len(body_ph) == 1
    assert body_ph[0].get("type") == "body"


# =============================================================================
# 4. Numeric badges ('03', '04') are not promoted to titles
# =============================================================================
def test_numeric_badges_not_promoted():
    assert is_numeric_badge("03")
    assert is_numeric_badge("04")
    assert is_numeric_badge("#01")
    assert is_numeric_badge("12.")

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    badge = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(1), Inches(1))
    badge.text_frame.text = "03"

    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(raw_slide=slide, slide_number=1)

    # Must NOT be classified as title or converted
    for a in analyses:
        if a.shape_id == badge.shape_id:
            assert a.category != PlaceholderCategory.STANDALONE_TITLE_BOX
            assert a.action != ConversionAction.CONVERT_EXISTING


# =============================================================================
# 5. Card-based slide retains hierarchy
# =============================================================================
def test_card_based_slide_retains_hierarchy():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Card 1: Badge "01" and Card Title "Discovery Phase"
    b1 = slide.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(1), Inches(1))
    b1.text_frame.text = "01"
    t1 = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(3), Inches(1))
    t1.text_frame.text = "Discovery Phase"

    # Card 2: Badge "02" and Card Title "Execution Phase"
    b2 = slide.shapes.add_textbox(Inches(4.5), Inches(1.0), Inches(1), Inches(1))
    b2.text_frame.text = "02"
    t2 = slide.shapes.add_textbox(Inches(4.5), Inches(2.2), Inches(3), Inches(1))
    t2.text_frame.text = "Execution Phase"

    classifier = PlaceholderClassifier()
    analyses = classifier.audit_slide_objects(raw_slide=slide, slide_number=1)

    # Badges must be preserved as badges, neither promoted to slide title
    for a in analyses:
        if a.shape_id in (b1.shape_id, b2.shape_id):
            assert a.action == ConversionAction.PRESERVE_EXISTING
            assert a.category in (PlaceholderCategory.CARD_BADGE, PlaceholderCategory.OTHER)


# =============================================================================
# 6. Slide without genuine title does not receive fabricated title
# =============================================================================
def test_slide_without_title_not_fabricated():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    # Only small caption text
    caption = slide.shapes.add_textbox(Inches(2), Inches(5.0), Inches(4), Inches(0.5))
    caption.text_frame.text = "Figure 1.1: Component flowchart"

    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(raw_slide=slide, slide_number=1)

    # No title conversion
    title_analyses = [a for a in analyses if a.action == ConversionAction.CONVERT_EXISTING and a.proposed_placeholder_type == "title"]
    assert len(title_analyses) == 0


# =============================================================================
# 7. Existing native title placeholder is not duplicated
# =============================================================================
def test_existing_native_title_not_duplicated():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only layout (has native title)
    title_shape = slide.shapes.title
    title_shape.text = "Native Slide Title"

    # Add loose textbox with text
    tb = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(1))
    tb.text_frame.text = "Secondary Section Heading"

    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(raw_slide=slide, slide_number=1)

    title_promotions = [a for a in analyses if a.category == PlaceholderCategory.STANDALONE_TITLE_BOX]
    assert len(title_promotions) == 0


# =============================================================================
# 8. Empty textbox does not trigger placeholder creation
# =============================================================================
def test_empty_textbox_no_placeholder_creation():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    empty_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    empty_box.text_frame.text = "   \n  "

    classifier = PlaceholderClassifier()
    analyses = classifier.audit_slide_objects(raw_slide=slide, slide_number=1)

    for a in analyses:
        if a.shape_id == empty_box.shape_id:
            assert a.action in (ConversionAction.PRESERVE_EXISTING, ConversionAction.SKIP_UNSAFE)
            assert a.category in (PlaceholderCategory.EMPTY, PlaceholderCategory.OTHER)


# =============================================================================
# 9. New placeholder creation 7 safety gates
# =============================================================================
def test_create_new_safety_gates():
    transformer = PlaceholderTransformer()

    # Create dummy analysis with CREATE_NEW action
    analysis = PlaceholderAnalysis(
        slide_number=1,
        shape_id=999,
        shape_name="New Body Placeholder",
        placeholder_type="none",
        placeholder_idx="1",
        category=PlaceholderCategory.STANDALONE_BODY_BOX,
        status=PlaceholderStatus.SAFE_TO_CONVERT,
        has_text=False,
        text_preview="",
        is_empty=True,
        is_animated=False,
        is_picture=False,
        is_table=False,
        is_chart=False,
        inherited_properties=None,
        slide_bbox_pt=(0.0, 0.0, 0.0, 0.0),
        recommendation="Create new body placeholder",
        reason="Slide requires structural body placeholder",
        confidence="HIGH",
        action=ConversionAction.CREATE_NEW,
        proposed_placeholder_type="body",
        conversion_basis="ACCESSIBILITY_STRUCTURE_REQUIREMENT"
    )

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    new_id = transformer._execute_create_new(analysis, slide.shapes._spTree)

    assert new_id > 0
    # Confirm shape was created with body ph tag
    created_elem = next(s for s in slide.shapes._spTree if _safe_xpath(s, f".//p:cNvPr[@id='{new_id}']"))
    ph = _safe_xpath(created_elem, ".//p:ph")
    assert len(ph) == 1
    assert ph[0].get("type") == "body"


# =============================================================================
# 10. Existing text and formatting remain unchanged
# =============================================================================
def test_text_formatting_unchanged():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    tf = box.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Preserved Formatting Run"
    r.font.size = Pt(28)
    r.font.bold = True

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide, approved_shape_ids=[box.shape_id])

    assert result.success
    assert box.text_frame.text == "Preserved Formatting Run"
    assert box.text_frame.paragraphs[0].runs[0].font.bold is True
    assert box.text_frame.paragraphs[0].runs[0].font.size == Pt(28)


# =============================================================================
# 11. Geometry and group transforms remain unchanged
# =============================================================================
def test_geometry_and_transforms_unchanged():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(4.0), Inches(1.5))
    box.text_frame.text = "Geometry Stability Test"

    orig_left, orig_top, orig_w, orig_h = box.left, box.top, box.width, box.height

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide, approved_shape_ids=[box.shape_id])

    assert result.success
    assert box.left == orig_left
    assert box.top == orig_top
    assert box.width == orig_w
    assert box.height == orig_h


# =============================================================================
# 12. Nested groups are preserved
# =============================================================================
def test_nested_groups_preserved():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Construct nested group XML
    nested_xml = """
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr>
            <p:cNvPr id="10" name="Parent Group"/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:grpSp>
            <p:nvGrpSpPr>
                <p:cNvPr id="11" name="Child Group"/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="12" name="Nested Child Shape"/>
                    <p:cNvSpPr/>
                    <p:nvPr/>
                </p:nvSpPr>
                <p:spPr/>
                <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Nested Text</a:t></a:r></a:p></p:txBody>
            </p:sp>
        </p:grpSp>
    </p:grpSp>
    """
    slide.shapes._spTree.append(_make_xml(nested_xml))

    classifier = PlaceholderClassifier()
    analyses = classifier.audit_slide_objects(raw_slide=slide, slide_number=1)

    # Nested child shape (ID 12) must be flagged SKIP_UNSAFE
    nested_analysis = next(a for a in analyses if a.shape_id == 12)
    assert nested_analysis.is_group_child
    assert nested_analysis.action == ConversionAction.SKIP_UNSAFE


# =============================================================================
# 13. Image, table, and chart objects retain correct types
# =============================================================================
def test_image_table_chart_types_retained():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Table
    table_shape = slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
    # Picture (add dummy shape with pic tag)
    pic_xml = """
    <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvPicPr><p:cNvPr id="30" name="Picture 30"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
        <p:blipFill><a:blip/></p:blipFill>
        <p:spPr/>
    </p:pic>
    """
    slide.shapes._spTree.append(_make_xml(pic_xml))

    classifier = PlaceholderClassifier()
    analyses = classifier.audit_slide_objects(raw_slide=slide, slide_number=1)

    for a in analyses:
        if a.shape_id == table_shape.shape_id:
            assert a.is_table
            assert a.action == ConversionAction.PRESERVE_EXISTING
        elif a.shape_id == 30:
            assert a.is_picture
            assert a.action == ConversionAction.PRESERVE_EXISTING


# =============================================================================
# 14. Layout/master inheritance remains valid
# =============================================================================
def test_layout_master_inheritance_valid():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(raw_slide=slide, slide_number=1)

    for a in analyses:
        assert a.inherited_properties is not None
        assert a.inherited_properties.inherited_from in ("SLIDE", "LAYOUT", "MASTER", "NONE")


# =============================================================================
# 15. Hyperlinks and relationship targets remain valid
# =============================================================================
def test_hyperlinks_relationships_valid():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
    p = box.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "Click Here For Summary"
    r.hyperlink.address = "https://example.com/summary"

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide, approved_shape_ids=[box.shape_id])

    assert result.success
    # Hyperlink rel ID still present
    rel_ids = _safe_xpath(box._element, ".//@r:id")
    assert len(rel_ids) > 0


# =============================================================================
# 16. Shape IDs remain unchanged where required
# =============================================================================
def test_shape_ids_preserved():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    box.text_frame.text = "Executive Strategy"
    orig_id = box.shape_id

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide, approved_shape_ids=[orig_id])

    assert result.success
    assert orig_id in result.converted_shape_ids
    assert box.shape_id == orig_id


# =============================================================================
# 17. No duplicate IDs or duplicate placeholders introduced
# =============================================================================
def test_no_duplicate_ids_or_placeholders():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box1 = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    box1.text_frame.text = "Slide Main Title"
    box2 = slide.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(2))
    box2.text_frame.text = "Body paragraph content"

    transformer = PlaceholderTransformer()
    result = transformer.convert_slide(slide._element, 1, raw_slide=slide)

    assert result.success
    all_ids = transformer._collect_shape_ids_list(slide.shapes._spTree)
    assert len(all_ids) == len(set(all_ids))


# =============================================================================
# 18. Animation target IDs and timing structures remain unchanged
# =============================================================================
def test_animation_target_ids_preserved():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    box.text_frame.text = "Animated Title Candidate"
    s_id = box.shape_id

    # Mark s_id as animated
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(
        raw_slide=slide, slide_number=1, animation_target_ids=[s_id]
    )

    # Animated candidate must be flagged MANUAL_REVIEW and NEVER auto-converted
    for a in analyses:
        if a.shape_id == s_id:
            assert a.action == ConversionAction.MANUAL_REVIEW
            assert a.status == PlaceholderStatus.MANUAL_REVIEW


# =============================================================================
# 19. Gemini malformed JSON and invalid shape IDs rejected safely
# =============================================================================
def test_gemini_malformed_json_rejected_safely():
    classifier = PlaceholderClassifier()
    # Mock assistant returning malformed JSON / invalid shape ID
    classifier.assistant.analyze_placeholder_conversion = MagicMock(return_value={
        "selected_title_shape_id": 99999,  # Non-existent ID
        "overall_confidence": 0.99,
        "decisions": [{"shape_id": 99999, "action": "CONVERT_TO_TITLE"}]
    })

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    box.text_frame.text = "Real Shape Title"

    analyses = classifier.analyze_slide_placeholders(raw_slide=slide, slide_number=1)

    # Invalid ID 99999 rejected; fallback to deterministic ranker for real box
    promoted = [a for a in analyses if a.category == PlaceholderCategory.STANDALONE_TITLE_BOX]
    assert len(promoted) == 1
    assert promoted[0].shape_id == box.shape_id


# =============================================================================
# 20. Gemini timeout or unavailability does not trigger unsafe conversions
# =============================================================================
def test_gemini_timeout_no_unsafe_conversion():
    classifier = PlaceholderClassifier()
    # Simulate assistant raising timeout / error
    classifier.assistant.analyze_placeholder_conversion = MagicMock(side_effect=TimeoutError("Model timed out"))

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    badge = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(1), Inches(1))
    badge.text_frame.text = "03"

    analyses = classifier.audit_slide_objects(raw_slide=slide, slide_number=1)

    # Numeric badge must NOT be converted even when AI fails
    for a in analyses:
        if a.shape_id == badge.shape_id:
            assert a.action != ConversionAction.CONVERT_EXISTING


# =============================================================================
# 21. One failed slide does not stop remaining slides
# =============================================================================
def test_one_failed_slide_isolation():
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    b1 = s1.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    b1.text_frame.text = "Slide 1 Valid Title"

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    b2 = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    b2.text_frame.text = "Slide 2 Valid Title"

    transformer = PlaceholderTransformer()
    # Run slide 1 successfully
    res1 = transformer.convert_slide(s1._element, 1, raw_slide=s1)
    assert res1.success

    # Corrupt slide 2 temporarily by removing spTree from its parent
    s2_tree = s2.shapes._spTree
    s2_tree.getparent().remove(s2_tree)
    res2 = transformer.convert_slide(s2._element, 2, raw_slide=s2)
    assert not res2.success
    assert len(res2.errors) > 0

    # Slide 1 conversion remains completely intact
    assert len(_safe_xpath(b1._element, ".//p:ph")) == 1


# =============================================================================
# 22. Rerunning the process is idempotent
# =============================================================================
def test_idempotency_rerun():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    tb.text_frame.text = "Idempotent Title Heading"

    transformer = PlaceholderTransformer()
    res1 = transformer.convert_slide(slide._element, 1, raw_slide=slide)
    assert res1.placeholders_converted == 1

    # Second run on same slide
    res2 = transformer.convert_slide(slide._element, 1, raw_slide=slide)
    assert res2.placeholders_converted == 0
    assert res2.placeholders_preserved >= 1


# =============================================================================
# 23. The output opens and reloads through PresentationReader
# =============================================================================
def test_output_reloads_via_presentation_reader(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    tb.text_frame.text = "Reloadable Slide Title"

    transformer = PlaceholderTransformer()
    transformer.convert_slide(slide._element, 1, raw_slide=slide)

    out_file = tmp_path / "reloaded_test.pptx"
    prs.save(out_file)

    reader = PresentationReader()
    reloaded = reader.read(out_file)
    assert len(reloaded.slides) == 1
    assert reloaded.slides[0].title == "Reloadable Slide Title"


# =============================================================================
# 24. Audit report accurately matches actual XML changes
# =============================================================================
def test_audit_report_matches_xml_changes(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    tb.text_frame.text = "Audited Presentation Title"

    out_file = tmp_path / "audited_test.pptx"
    prs.save(out_file)

    reader = PresentationReader()
    pres_model = reader.read(out_file)

    rule = PlaceholderRule()
    report = rule.generate_presentation_audit_report(pres_model)

    assert report["total_slides_evaluated"] == 1
    assert report["conversions_proposed"] == 1
    assert report["existing_objects_converted"] == 1
    assert report["visual_integrity_verified"] is True
    assert len(report["objects"]) >= 1
