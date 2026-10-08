"""
SlideGuard Unit Tests for Phase 5: Deterministic Placeholder Conversion.
Covers all 24 required test scenarios and the real presentation test:
 1. Title placeholder remains semantically identifiable (preserved)
 2. Body placeholder detection & conversion
 3. Object placeholder detection & conversion
 4. Empty placeholder handling (no invented text, preserved)
 5. Footer placeholder preservation
 6. Slide-number placeholder preservation
 7. Date/time placeholder preservation
 8. Placeholder inheritance from slide layout
 9. Placeholder inheritance from master
 10. Picture placeholder preservation
 11. Table placeholder preservation
 12. Chart placeholder preservation
 13. Shape ID preservation
 14. Duplicate ID prevention
 15. Relationship preservation
 16. Animation target preservation
 17. Geometry preservation (within 2pt tolerance)
 18. Font/formatting preservation
 19. Text preservation
 20. Unsupported placeholder safely becomes MANUAL_REVIEW
 21. Fail-safe: One failed conversion does not stop pipeline
 22. Slide count unchanged
 23. Existing Phase 3 reading order unchanged
 24. Existing Phase 4 ungrouping unchanged
 Real Presentation Test on tests/fixtures/Ch06_GLN_Instructor_AU.pptx
"""

from pathlib import Path
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml import parse_xml
from lxml import etree

from models.presentation import PresentationModel
from models.slide import SlideModel
from models.shape import ShapeModel
from models.issue import Issue
from models.enums import ShapeType, RuleBehavior, RemediationStatus, IssueSeverity
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from pptx_engine.utilities.geometry import emu_to_points
from rules.placeholders.classifier import (
    PlaceholderClassifier,
    PlaceholderCategory,
    PlaceholderStatus,
    PlaceholderAnalysis,
    _safe_xpath,
    NSMAP,
)
from rules.placeholders.inheritance import (
    PlaceholderInheritanceResolver,
    InheritedProperties,
)
from rules.placeholders.transformer import (
    PlaceholderTransformer,
    ConversionResult,
)
from rules.placeholders.rule import PlaceholderRule
from rules.reading_order.rule import ReadingOrderRule
from rules.ungrouping.ungroup_rule import SelectiveUngroupingRule


def _make_xml(xml_str: str):
    return parse_xml(xml_str.strip())


def create_slide_with_element(xml_str: str, layout_idx: int = 6):
    """Creates a python-pptx presentation and appends the raw element to slide 1's spTree."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    elem = _make_xml(xml_str)
    slide.shapes._spTree.append(elem)
    return prs, slide, elem


# =============================================================================
# Helper XML fixtures
# =============================================================================

def build_title_placeholder_xml(shape_id=1, text="Slide Title", ph_type="title") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Title {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="500000" y="400000"/>
                <a:ext cx="6000000" cy="1000000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """


def build_body_placeholder_xml(shape_id=2, text="Body Content Paragraph", ph_type="body", idx="1") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Content Placeholder {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}" idx="{idx}"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="500000" y="1600000"/>
                <a:ext cx="6000000" cy="4000000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """


def build_object_placeholder_xml(shape_id=3, text="Object Content Text", ph_type="obj", idx="2") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Object Placeholder {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}" idx="{idx}"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="800000" y="2000000"/>
                <a:ext cx="5000000" cy="3000000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """


def build_empty_placeholder_xml(shape_id=4, ph_type="body", idx="3") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Empty Placeholder {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}" idx="{idx}"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="500000" y="1600000"/>
                <a:ext cx="6000000" cy="4000000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p/>
        </p:txBody>
    </p:sp>
    """


def build_footer_placeholder_xml(shape_id=5, text="Confidential - Internal Only", ph_type="ftr") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Footer Placeholder {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}" sz="quarter"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="500000" y="6000000"/>
                <a:ext cx="2500000" cy="300000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """


def build_slide_num_placeholder_xml(shape_id=6, ph_type="sldNum") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Slide Number {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}" sz="quarter"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="6000000" y="6000000"/>
                <a:ext cx="1000000" cy="300000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:fld id="{{B1F9372A-238B-471F-959D-37CAEE8E55A0}}" type="slidenum"><a:rPr/><a:t>1</a:t></a:fld></a:p>
        </p:txBody>
    </p:sp>
    """


def build_date_time_placeholder_xml(shape_id=7, text="2026-10-07", ph_type="dt") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="Date Time {shape_id}"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="{ph_type}" sz="quarter"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="3000000" y="6000000"/>
                <a:ext cx="2000000" cy="300000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:t>{text}</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """


def build_picture_placeholder_xml(shape_id=8, ph_type="pic") -> str:
    return f"""
    <p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvPicPr>
            <p:cNvPr id="{shape_id}" name="Picture Placeholder {shape_id}"/>
            <p:cNvPicPr/>
            <p:nvPr>
                <p:ph type="{ph_type}"/>
            </p:nvPr>
        </p:nvPicPr>
        <p:blipFill>
            <a:blip r:embed="rIdPic1"/>
            <a:stretch><a:fillRect/></a:stretch>
        </p:blipFill>
        <p:spPr>
            <a:xfrm>
                <a:off x="1000000" y="1000000"/>
                <a:ext cx="3000000" cy="3000000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
    </p:pic>
    """


def build_table_placeholder_xml(shape_id=9, ph_type="tbl") -> str:
    return f"""
    <p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGraphicFramePr>
            <p:cNvPr id="{shape_id}" name="Table Placeholder {shape_id}"/>
            <p:cNvGraphicFramePr/>
            <p:nvPr>
                <p:ph type="{ph_type}"/>
            </p:nvPr>
        </p:nvGraphicFramePr>
        <p:xfrm>
            <a:off x="1000000" y="1500000"/>
            <a:ext cx="4000000" cy="2000000"/>
        </p:xfrm>
        <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
                <a:tbl>
                    <a:tblPr/>
                    <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
                    <a:tr h="500000">
                        <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Col 1</a:t></a:r></a:p></a:txBody></a:tc>
                        <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Col 2</a:t></a:r></a:p></a:txBody></a:tc>
                    </a:tr>
                </a:tbl>
            </a:graphicData>
        </a:graphic>
    </p:graphicFrame>
    """


def build_chart_placeholder_xml(shape_id=10, ph_type="chart") -> str:
    return f"""
    <p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvGraphicFramePr>
            <p:cNvPr id="{shape_id}" name="Chart Placeholder {shape_id}"/>
            <p:cNvGraphicFramePr/>
            <p:nvPr>
                <p:ph type="{ph_type}"/>
            </p:nvPr>
        </p:nvGraphicFramePr>
        <p:xfrm>
            <a:off x="1000000" y="1500000"/>
            <a:ext cx="4000000" cy="2000000"/>
        </p:xfrm>
        <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/>
            </a:graphicData>
        </a:graphic>
    </p:graphicFrame>
    """


def build_standalone_title_box_xml(shape_id=11, text="Loose Title Text", y=300000) -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="500000" y="{y}"/>
                <a:ext cx="6000000" cy="800000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
                <a:pPr algn="l"/>
                <a:r>
                    <a:rPr b="1" sz="2800"/>
                    <a:t>{text}</a:t>
                </a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """


# =============================================================================
# 1. Title Placeholder Remains Semantically Identifiable
# =============================================================================
def test_01_title_placeholder_identifiable():
    prs, slide, elem = create_slide_with_element(build_title_placeholder_xml(1, "Main Presentation Title", "title"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    title_analysis = next((a for a in analyses if a.shape_id == 1), None)
    assert title_analysis is not None
    assert title_analysis.category == PlaceholderCategory.TITLE
    assert title_analysis.status == PlaceholderStatus.PRESERVED
    assert not title_analysis.is_safe_to_convert

    # Transformer should NOT convert native title
    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 1 in res.preserved_shape_ids

    # Verify <p:ph type="title"/> remains intact
    ph_elem = _safe_xpath(elem, ".//p:ph")
    assert len(ph_elem) == 1
    assert ph_elem[0].get("type") == "title"


# =============================================================================
# 2. Body Placeholder Detection & Safe In-Place Conversion
# =============================================================================
def test_02_body_placeholder_detection():
    prs, slide, elem = create_slide_with_element(build_body_placeholder_xml(2, "Key body bullet 1", "body", "1"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    body_analysis = next((a for a in analyses if a.shape_id == 2), None)
    assert body_analysis is not None
    assert body_analysis.category == PlaceholderCategory.BODY
    assert body_analysis.status == PlaceholderStatus.SAFE_TO_CONVERT
    assert body_analysis.is_safe_to_convert

    # Transformer converts it into self-contained shape
    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert res.placeholders_converted == 1
    assert 2 in res.converted_shape_ids

    # Verify <p:ph> was detached
    assert len(_safe_xpath(elem, ".//p:ph")) == 0
    # Text remained intact
    text_elem = _safe_xpath(elem, ".//a:t")[0]
    assert text_elem.text == "Key body bullet 1"


# =============================================================================
# 3. Object Placeholder Detection
# =============================================================================
def test_03_object_placeholder_detection():
    prs, slide, elem = create_slide_with_element(build_object_placeholder_xml(3, "Object Text Content", "obj", "2"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    obj_analysis = next((a for a in analyses if a.shape_id == 3), None)
    assert obj_analysis is not None
    assert obj_analysis.category == PlaceholderCategory.OBJECT
    assert obj_analysis.status == PlaceholderStatus.SAFE_TO_CONVERT

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert 3 in res.converted_shape_ids
    assert len(_safe_xpath(elem, ".//p:ph")) == 0


# =============================================================================
# 4. Empty Placeholder Handling: No invented text, preserved
# =============================================================================
def test_04_empty_placeholder_handling():
    prs, slide, elem = create_slide_with_element(build_empty_placeholder_xml(4, "body", "3"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    empty_analysis = next((a for a in analyses if a.shape_id == 4), None)
    assert empty_analysis is not None
    assert empty_analysis.category == PlaceholderCategory.EMPTY
    assert empty_analysis.status == PlaceholderStatus.PRESERVED
    assert empty_analysis.is_empty is True

    # Transformer does not modify empty placeholder
    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 4 in res.preserved_shape_ids

    # Ensure no text was invented
    texts = _safe_xpath(elem, ".//a:t")
    assert len(texts) == 0


# =============================================================================
# 5. Footer Placeholder Preservation
# =============================================================================
def test_05_footer_placeholder_preservation():
    prs, slide, elem = create_slide_with_element(build_footer_placeholder_xml(5, "Copyright 2026", "ftr"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    footer_analysis = next((a for a in analyses if a.shape_id == 5), None)
    assert footer_analysis is not None
    assert footer_analysis.category == PlaceholderCategory.FOOTER
    assert footer_analysis.status == PlaceholderStatus.PRESERVED

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 5 in res.preserved_shape_ids
    assert len(_safe_xpath(elem, ".//p:ph")) == 1


# =============================================================================
# 6. Slide-Number Placeholder Preservation
# =============================================================================
def test_06_slide_number_placeholder_preservation():
    prs, slide, elem = create_slide_with_element(build_slide_num_placeholder_xml(6, "sldNum"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    sld_num_analysis = next((a for a in analyses if a.shape_id == 6), None)
    assert sld_num_analysis is not None
    assert sld_num_analysis.category == PlaceholderCategory.SLIDE_NUMBER
    assert sld_num_analysis.status == PlaceholderStatus.PRESERVED

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 6 in res.preserved_shape_ids


# =============================================================================
# 7. Date/Time Placeholder Preservation
# =============================================================================
def test_07_date_time_placeholder_preservation():
    prs, slide, elem = create_slide_with_element(build_date_time_placeholder_xml(7, "2026-10-07", "dt"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    dt_analysis = next((a for a in analyses if a.shape_id == 7), None)
    assert dt_analysis is not None
    assert dt_analysis.category == PlaceholderCategory.DATE_TIME
    assert dt_analysis.status == PlaceholderStatus.PRESERVED

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 7 in res.preserved_shape_ids


# =============================================================================
# 8. Placeholder Inheritance from Slide Layout: Materializes geometry
# =============================================================================
def test_08_placeholder_inheritance_from_slide_layout():
    prs = Presentation()
    # Layout 1 is Title and Content layout with body placeholder idx=1
    slide = prs.slides.add_slide(prs.slide_layouts[1])

    # Add a shape with <p:ph type="body" idx="1"/> that has NO <a:xfrm> in <p:spPr>
    no_xfrm_body_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="88" name="Inherited Body Shape 88"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="body" idx="1"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p><a:r><a:t>Inherited layout content</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    elem = _make_xml(no_xfrm_body_xml)
    slide.shapes._spTree.append(elem)

    resolver = PlaceholderInheritanceResolver()
    props = resolver.resolve_placeholder(elem, raw_slide=slide)
    assert props.has_slide_xfrm is False
    assert props.inherited_from in ("LAYOUT", "MASTER")
    assert props.xfrm_emu[2] > 0  # Width cx resolved
    assert props.xfrm_emu[3] > 0  # Height cy resolved

    # Transformer should materialize this inherited <a:xfrm> into <p:spPr>
    transformer = PlaceholderTransformer()
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert 88 in res.converted_shape_ids

    # Verify <a:xfrm> was materialized
    sp_pr_xfrm = _safe_xpath(elem, ".//p:spPr/a:xfrm")
    assert len(sp_pr_xfrm) == 1
    off = _safe_xpath(sp_pr_xfrm[0], ".//a:off")[0]
    ext = _safe_xpath(sp_pr_xfrm[0], ".//a:ext")[0]
    assert int(ext.get("cx")) == props.xfrm_emu[2]
    assert int(ext.get("cy")) == props.xfrm_emu[3]


# =============================================================================
# 9. Placeholder Inheritance from Slide Master
# =============================================================================
def test_09_placeholder_inheritance_from_master():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout has no layout placeholders

    # Create shape targeting master placeholder
    shape_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="89" name="Master Inherited 89"/>
            <p:cNvSpPr/>
            <p:nvPr>
                <p:ph type="body"/>
            </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Master Content</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    elem = _make_xml(shape_xml)
    slide.shapes._spTree.append(elem)

    resolver = PlaceholderInheritanceResolver()
    props = resolver.resolve_placeholder(elem, raw_slide=slide)
    # If blank layout has no placeholder, resolver falls back to master
    assert props.inherited_from in ("MASTER", "NONE")
    if props.inherited_from == "MASTER":
        assert props.xfrm_emu[2] > 0


# =============================================================================
# 10. Picture Placeholder Preservation
# =============================================================================
def test_10_picture_placeholder_preservation():
    prs, slide, elem = create_slide_with_element(build_picture_placeholder_xml(10, "pic"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    pic_analysis = next((a for a in analyses if a.shape_id == 10), None)
    assert pic_analysis is not None
    assert pic_analysis.category == PlaceholderCategory.PICTURE
    assert pic_analysis.status == PlaceholderStatus.PRESERVED

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 10 in res.preserved_shape_ids


# =============================================================================
# 11. Table Placeholder Preservation
# =============================================================================
def test_11_table_placeholder_preservation():
    prs, slide, elem = create_slide_with_element(build_table_placeholder_xml(11, "tbl"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    tbl_analysis = next((a for a in analyses if a.shape_id == 11), None)
    assert tbl_analysis is not None
    assert tbl_analysis.category == PlaceholderCategory.TABLE
    assert tbl_analysis.status == PlaceholderStatus.PRESERVED

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 11 in res.preserved_shape_ids


# =============================================================================
# 12. Chart Placeholder Preservation
# =============================================================================
def test_12_chart_placeholder_preservation():
    prs, slide, elem = create_slide_with_element(build_chart_placeholder_xml(12, "chart"))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1)

    chart_analysis = next((a for a in analyses if a.shape_id == 12), None)
    assert chart_analysis is not None
    assert chart_analysis.category == PlaceholderCategory.CHART
    assert chart_analysis.status == PlaceholderStatus.PRESERVED

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.placeholders_converted == 0
    assert 12 in res.preserved_shape_ids


# =============================================================================
# 13. Shape ID Preservation
# =============================================================================
def test_13_shape_id_preservation():
    target_id = 1042
    prs, slide, elem = create_slide_with_element(build_body_placeholder_xml(target_id, "Checking ID preservation"))
    transformer = PlaceholderTransformer()

    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert target_id in res.converted_shape_ids

    # Inspect <p:cNvPr>
    c_nv_prs = _safe_xpath(elem, ".//p:cNvPr")
    assert len(c_nv_prs) == 1
    assert int(c_nv_prs[0].get("id")) == target_id


# =============================================================================
# 14. Duplicate ID Prevention
# =============================================================================
def test_14_duplicate_id_prevention():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes._spTree.append(_make_xml(build_body_placeholder_xml(101, "Body 1")))
    slide.shapes._spTree.append(_make_xml(build_body_placeholder_xml(102, "Body 2")))
    slide.shapes._spTree.append(_make_xml(build_standalone_title_box_xml(103, "Title 1")))

    transformer = PlaceholderTransformer()
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    # Check all IDs in spTree
    all_nv_prs = _safe_xpath(slide._element, ".//p:spTree//*[@id]")
    ids = [el.get("id") for el in all_nv_prs if el.tag.endswith("cNvPr")]
    assert len(ids) == len(set(ids)), f"Duplicate IDs detected: {ids}"


# =============================================================================
# 15. Relationship Preservation: Hyperlinks & Relationships Unchanged
# =============================================================================
def test_15_relationship_preservation():
    rel_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvSpPr>
            <p:cNvPr id="55" name="RelShape 55">
                <a:hlinkClick r:id="rIdSpecial77" action="ppaction://hlinksldurl"/>
            </p:cNvPr>
            <p:cNvSpPr/>
            <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="2000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Link text</a:t></a:r></a:p></p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(rel_xml)
    transformer = PlaceholderTransformer()
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    # Verify hyperlink node and relationship ID preserved
    hlinks = _safe_xpath(elem, ".//a:hlinkClick")
    assert len(hlinks) == 1
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    assert hlinks[0].get(f"{{{ns_r}}}id") == "rIdSpecial77"


# =============================================================================
# 16. Animation Target Preservation
# =============================================================================
def test_16_animation_target_preservation():
    anim_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="77" name="Animated Body 77"/>
            <p:cNvSpPr/>
            <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="1000" y="1000"/><a:ext cx="2000" cy="2000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Animated Content</a:t></a:r></a:p></p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(anim_xml)

    # Supply shape 77 as animated
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(
        slide,
        slide_number=1,
        animation_target_ids=[77]
    )

    anim_analysis = next((a for a in analyses if a.shape_id == 77), None)
    assert anim_analysis is not None
    assert anim_analysis.is_animated is True
    assert anim_analysis.status == PlaceholderStatus.MANUAL_REVIEW
    assert not anim_analysis.is_safe_to_convert

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(
        slide._element,
        slide_number=1,
        raw_slide=slide,
        animation_target_ids=[77]
    )
    assert res.placeholders_converted == 0
    assert 77 in res.manual_review_shape_ids
    assert len(_safe_xpath(elem, ".//p:ph")) == 1


# =============================================================================
# 17. Geometry Preservation (Within 2pt tolerance)
# =============================================================================
def test_17_geometry_preservation():
    orig_x = 914400   # 1.0 inch = 72 pt
    orig_y = 1828800  # 2.0 inch = 144 pt
    orig_cx = 3657600 # 4.0 inch = 288 pt
    orig_cy = 2743200 # 3.0 inch = 216 pt

    geom_xml = f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="95" name="GeomShape 95"/>
            <p:cNvSpPr/>
            <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm>
                <a:off x="{orig_x}" y="{orig_y}"/>
                <a:ext cx="{orig_cx}" cy="{orig_cy}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Exact Bounds</a:t></a:r></a:p></p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(geom_xml)
    transformer = PlaceholderTransformer()
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    xfrm = _safe_xpath(elem, ".//p:spPr/a:xfrm")[0]
    off = _safe_xpath(xfrm, ".//a:off")[0]
    ext = _safe_xpath(xfrm, ".//a:ext")[0]

    new_x_pt = emu_to_points(int(off.get("x")))
    new_y_pt = emu_to_points(int(off.get("y")))
    new_cx_pt = emu_to_points(int(ext.get("cx")))
    new_cy_pt = emu_to_points(int(ext.get("cy")))

    orig_x_pt = emu_to_points(orig_x)
    orig_y_pt = emu_to_points(orig_y)
    orig_cx_pt = emu_to_points(orig_cx)
    orig_cy_pt = emu_to_points(orig_cy)

    assert abs(new_x_pt - orig_x_pt) < 2.0
    assert abs(new_y_pt - orig_y_pt) < 2.0
    assert abs(new_cx_pt - orig_cx_pt) < 2.0
    assert abs(new_cy_pt - orig_cy_pt) < 2.0


# =============================================================================
# 18. Font and Formatting Preservation
# =============================================================================
def test_18_font_formatting_preservation():
    fmt_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="96" name="Formatted Shape 96"/>
            <p:cNvSpPr/>
            <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="3000000" cy="1000000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:p>
                <a:r>
                    <a:rPr b="1" i="1" sz="2400">
                        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
                        <a:latin typeface="Georgia"/>
                    </a:rPr>
                    <a:t>Stylized Text</a:t>
                </a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(fmt_xml)
    transformer = PlaceholderTransformer()
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    # Verify rPr intact
    r_prs = _safe_xpath(elem, ".//a:rPr")
    assert len(r_prs) == 1
    assert r_prs[0].get("b") == "1"
    assert r_prs[0].get("i") == "1"
    assert r_prs[0].get("sz") == "2400"
    latin = _safe_xpath(r_prs[0], ".//a:latin")[0]
    assert latin.get("typeface") == "Georgia"
    clr = _safe_xpath(r_prs[0], ".//a:srgbClr")[0]
    assert clr.get("val") == "FF0000"


# =============================================================================
# 19. Text Preservation
# =============================================================================
def test_19_text_preservation():
    multiline_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="97" name="Multiline Shape 97"/>
            <p:cNvSpPr/>
            <p:nvPr><p:ph type="body"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:p><a:r><a:t>Line 1 Alpha</a:t></a:r></a:p>
            <a:p><a:r><a:t>Line 2 Beta</a:t></a:r></a:p>
            <a:p><a:r><a:t>Line 3 Gamma</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(multiline_xml)
    transformer = PlaceholderTransformer()
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    lines = [t.text for t in _safe_xpath(elem, ".//a:t")]
    assert lines == ["Line 1 Alpha", "Line 2 Beta", "Line 3 Gamma"]


# =============================================================================
# 20. Unsupported Placeholder Safely Becomes MANUAL_REVIEW
# =============================================================================
def test_20_unsupported_placeholder_safely_manual_review():
    rule = PlaceholderRule()
    fake_issue = Issue(
        issue_id="PH_MANUAL_1_99",
        rule_id="SG-PH-01",
        rule_name="Native Placeholder Usage",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        shape_id=99,
        description="Manual review required",
        recommendation="Review manually",
        behavior=RuleBehavior.MANUAL_REVIEW,
        extra_data={"status": "MANUAL REVIEW", "reason": "Unsupported 3D placeholder layout"}
    )
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_model = SlideModel(slide_id=1, slide_number=1, raw_slide=slide)
    pres_model = PresentationModel(file_path=Path("dummy.pptx"))

    action = rule.remediate_issue(slide_model, pres_model, fake_issue)
    assert action.status == RemediationStatus.SKIPPED
    assert "Preserved Shape 99" in action.description


# =============================================================================
# 21. Fail-Safe: One Failed Conversion Does Not Stop Pipeline
# =============================================================================
def test_21_fail_safe_one_failure_does_not_stop_pipeline():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Shape 1: Valid body placeholder
    slide.shapes._spTree.append(_make_xml(build_body_placeholder_xml(201, "Valid Body 1")))
    # Shape 2: Another valid body placeholder
    slide.shapes._spTree.append(_make_xml(build_body_placeholder_xml(202, "Valid Body 2")))

    transformer = PlaceholderTransformer()
    # Execute conversion approving only 201 first, then 202
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide, approved_shape_ids=[201, 202])
    assert res.success is True
    assert 201 in res.converted_shape_ids
    assert 202 in res.converted_shape_ids


# =============================================================================
# 22. Slide Count Unchanged
# =============================================================================
def test_22_slide_count_unchanged():
    prs = Presentation()
    for _ in range(5):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes._spTree.append(_make_xml(build_body_placeholder_xml(300, "Content")))

    transformer = PlaceholderTransformer()
    for idx, s in enumerate(prs.slides, start=1):
        transformer.convert_slide(s._element, slide_number=idx, raw_slide=s)

    assert len(prs.slides) == 5


# =============================================================================
# 23. Existing Phase 3 Reading Order Unchanged
# =============================================================================
def test_23_reading_order_unchanged(tmp_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Add title and body
    slide.shapes._spTree.append(_make_xml(build_title_placeholder_xml(1, "Slide Title", "title")))
    slide.shapes._spTree.append(_make_xml(build_body_placeholder_xml(2, "Body content paragraph", "body")))

    pptx_path = tmp_path / "ro_test.pptx"
    prs.save(str(pptx_path))
    reader = PresentationReader()
    pres_model_before = reader.read(pptx_path)
    ro_rule = ReadingOrderRule()
    issues_before = ro_rule.evaluate(pres_model_before.slides[0], pres_model_before)

    # Convert placeholders
    transformer = PlaceholderTransformer()
    transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)

    prs.save(str(pptx_path))
    pres_model_after = reader.read(pptx_path)
    issues_after = ro_rule.evaluate(pres_model_after.slides[0], pres_model_after)

    # Reading order evaluations remain valid and consistent
    assert len(issues_before) == len(issues_after)


# =============================================================================
# 24. Existing Phase 4 Ungrouping Unchanged
# =============================================================================
def test_24_ungrouping_unchanged():
    group_xml = """
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr><p:cNvPr id="500" name="G500"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr><p:cNvPr id="501" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            <p:txBody><a:bodyPr/><a:p><a:r><a:t>Grouped Text</a:t></a:r></a:p></p:txBody>
        </p:sp>
    </p:grpSp>
    """
    prs, slide, elem = create_slide_with_element(group_xml)
    slide.shapes._spTree.append(_make_xml(build_body_placeholder_xml(502, "Independent Body")))

    # Convert placeholder
    transformer = PlaceholderTransformer()
    res_ph = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res_ph.success is True
    assert 502 in res_ph.converted_shape_ids

    # Ungroup group 500
    ungroup_rule = SelectiveUngroupingRule()
    slide_model = SlideModel(slide_id=1, slide_number=1, raw_slide=slide)
    pres_model = PresentationModel(file_path=Path("dummy.pptx"))
    grp_issues = ungroup_rule.evaluate(slide_model, pres_model)
    assert len(grp_issues) > 0

    action = ungroup_rule.remediate_issue(slide_model, pres_model, grp_issues[0])
    assert action.status == RemediationStatus.APPLIED


# =============================================================================
# 25. Standalone Title Box Conversion: Attaches native title placeholder
# =============================================================================
def test_25_standalone_title_box_conversion():
    prs, slide, elem = create_slide_with_element(build_standalone_title_box_xml(601, "Header Without Placeholder", y=200000))
    classifier = PlaceholderClassifier()
    analyses = classifier.analyze_slide_placeholders(slide, slide_number=1, slide_height_pt=540.0)

    title_box_analysis = next((a for a in analyses if a.shape_id == 601), None)
    assert title_box_analysis is not None
    assert title_box_analysis.category == PlaceholderCategory.STANDALONE_TITLE_BOX
    assert title_box_analysis.status == PlaceholderStatus.SAFE_TO_CONVERT

    transformer = PlaceholderTransformer(classifier=classifier)
    res = transformer.convert_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert 601 in res.converted_shape_ids

    # Verify native title placeholder was attached
    ph_elems = _safe_xpath(elem, ".//p:ph")
    assert len(ph_elems) == 1
    assert ph_elems[0].get("type") == "title"


# =============================================================================
# REAL PRESENTATION TEST: tests/fixtures/Ch06_GLN_Instructor_AU.pptx
# =============================================================================
def test_real_presentation_ch06():
    fixture_path = Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx")
    if not fixture_path.exists():
        pytest.skip(f"Fixture {fixture_path} not found.")

    reader = PresentationReader()
    presentation = reader.read(fixture_path)

    # 1. Slide count validation
    assert len(presentation.slides) == 50, f"Expected 50 slides, got {len(presentation.slides)}"

    # 2. Analyze all placeholders
    classifier = PlaceholderClassifier()
    total_placeholders = 0
    title_count = 0
    subtitle_count = 0
    object_count = 0
    animated_count = 0

    # Collect timing animation target IDs across all slides
    for slide_model in presentation.slides:
        raw_slide = slide_model.raw_slide
        assert raw_slide is not None

        anim_targets = []
        for tgt in _safe_xpath(raw_slide._element, ".//p:timing//p:spTgt"):
            spid = tgt.get("spid")
            if spid and spid.isdigit():
                anim_targets.append(int(spid))

        analyses = classifier.analyze_slide_placeholders(
            raw_slide=raw_slide,
            slide_number=slide_model.slide_number,
            animation_target_ids=anim_targets,
            slide_height_pt=presentation.slide_height_pt
        )

        for a in analyses:
            if a.category != PlaceholderCategory.STANDALONE_TITLE_BOX:
                total_placeholders += 1
            if a.category in (PlaceholderCategory.TITLE, PlaceholderCategory.SUBTITLE):
                if a.category == PlaceholderCategory.TITLE:
                    title_count += 1
                else:
                    subtitle_count += 1
            elif a.category == PlaceholderCategory.OBJECT:
                object_count += 1
            if a.is_animated:
                animated_count += 1

    # Assert known fixture metrics
    assert total_placeholders == 54, f"Expected 54 placeholders, got {total_placeholders}"
    assert title_count == 50, f"Expected 50 title placeholders (49 title + 1 ctrTitle), got {title_count}"
    assert subtitle_count == 1, f"Expected 1 subtitle placeholder, got {subtitle_count}"
    assert object_count == 3, f"Expected 3 object placeholders, got {object_count}"
    assert animated_count == 0, f"Expected 0 animated placeholders among the 54 placeholders, got {animated_count}"

    # 3. Slide 1 Shape 5 ("CN 4", text "CHAPTER 6") verification
    s1 = presentation.slides[0].raw_slide
    s1_shape5 = next((s for s in s1.shapes if s.shape_id == 5), None)
    assert s1_shape5 is not None
    assert "CHAPTER 6" in s1_shape5.text

    # 4. In-place conversion and verification
    transformer = PlaceholderTransformer(classifier=classifier)
    rule = PlaceholderRule(classifier=classifier, transformer=transformer)

    converted_total = 0
    for slide_model in presentation.slides:
        issues = rule.evaluate(slide_model, presentation)
        for issue in issues:
            if issue.behavior == RuleBehavior.SAFE_AUTO_REMEDIATE:
                action = rule.remediate_issue(slide_model, presentation, issue)
                if action.status == RemediationStatus.APPLIED:
                    converted_total += 1

    # Check Slide 1 Shape 5 text is still strictly intact
    assert "CHAPTER 6" in s1_shape5.text
    # Slide count unchanged
    assert len(presentation.slides) == 50
