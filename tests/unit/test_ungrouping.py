"""
Comprehensive unit tests for SlideGuard Phase 4: Selective Ungrouping.
Covers all 20 required scenarios:
 1. Simple group (Text + Rectangle + Shape -> all eligible extracted)
 2. Mixed group (Text + Rectangle + Picture -> text/rect extracted, picture protected)
 3. Picture-only group (100% protected)
 4. Nested group (Group -> Group -> Text -> recursively processed)
 5. Nested mixed group (Group -> Text + Group -> Shape + Picture)
 6. Shape ID preservation
 7. Duplicate ID prevention
 8. Relationship preservation
 9. Hyperlink preservation
 10. Coordinate transformation formula
 11. Group scaling
 12. Group translation
 13. Nested group coordinate transformation
 14. Z-order preservation
 15. Animation target ID preservation
 16. Unsupported/ambiguous group fails safely
 17. One failed group does not stop other groups
 18. Original slide count unchanged
 19. Original text unchanged
 20. Original visual properties unchanged
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
from models.enums import ShapeType, RuleBehavior, RemediationStatus
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from rules.ungrouping.classifier import (
    GroupClassifier,
    GroupCategory,
    ChildCategory,
    GroupAnalysis,
    _safe_xpath,
)
from rules.ungrouping.transformer import (
    SelectiveUngroupingTransformer,
    UngroupResult,
)
from rules.ungrouping.ungroup_rule import SelectiveUngroupingRule
from pptx_engine.utilities.geometry import (
    calculate_group_child_slide_coordinates,
    apply_nested_group_transforms,
)


def _make_xml(xml_str: str):
    return parse_xml(xml_str.strip())


# =============================================================================
# Helper XML fixtures
# =============================================================================

def build_simple_group_xml(group_id=10, child1_id=11, child2_id=12, child3_id=13) -> str:
    return f"""
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr>
            <p:cNvPr id="{group_id}" name="Group {group_id}"/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="1000" y="2000"/>
                <a:ext cx="3000" cy="4000"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="300" cy="400"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="{child1_id}" name="Text Box {child1_id}"/>
                <p:cNvSpPr txBox="1"/>
                <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="100" cy="100"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
            <p:txBody>
                <a:bodyPr/>
                <a:lstStyle/>
                <a:p><a:r><a:t>Original Header Text</a:t></a:r></a:p>
            </p:txBody>
        </p:sp>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="{child2_id}" name="Rectangle {child2_id}"/>
                <p:cNvSpPr/>
                <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm>
                    <a:off x="100" y="100"/>
                    <a:ext cx="100" cy="100"/>
                </a:xfrm>
                <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
                <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
            </p:spPr>
        </p:sp>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="{child3_id}" name="Shape {child3_id}"/>
                <p:cNvSpPr/>
                <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm>
                    <a:off x="200" y="200"/>
                    <a:ext cx="100" cy="200"/>
                </a:xfrm>
                <a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>
            </p:spPr>
        </p:sp>
    </p:grpSp>
    """


def build_mixed_group_xml(group_id=20, text_id=21, rect_id=22, pic_id=23, r_id="rId9") -> str:
    return f"""
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvGrpSpPr>
            <p:cNvPr id="{group_id}" name="Mixed Group {group_id}"/>
            <p:cNvGrpSpPr/>
            <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="500" y="500"/>
                <a:ext cx="1000" cy="1000"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="1000" cy="1000"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="{text_id}" name="Text Box {text_id}"/>
                <p:cNvSpPr txBox="1"/>
                <p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="200"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
            <p:txBody>
                <a:bodyPr/><a:lstStyle/>
                <a:p><a:r><a:t>Mixed Text</a:t></a:r></a:p>
            </p:txBody>
        </p:sp>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="{rect_id}" name="Rectangle {rect_id}"/>
                <p:cNvSpPr/><p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm><a:off x="400" y="0"/><a:ext cx="200" cy="200"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
        </p:sp>
        <p:pic>
            <p:nvPicPr>
                <p:cNvPr id="{pic_id}" name="Picture {pic_id}"/>
                <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
                <p:nvPr/>
            </p:nvPicPr>
            <p:blipFill>
                <a:blip r:embed="{r_id}"/>
                <a:stretch><a:fillRect/></a:stretch>
            </p:blipFill>
            <p:spPr>
                <a:xfrm><a:off x="0" y="200"/><a:ext cx="600" cy="800"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
        </p:pic>
    </p:grpSp>
    """


def build_pic_only_group_xml(group_id=30, pic1_id=31, pic2_id=32) -> str:
    return f"""
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvGrpSpPr>
            <p:cNvPr id="{group_id}" name="Picture Group {group_id}"/>
            <p:cNvGrpSpPr/><p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/>
                <a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:pic>
            <p:nvPicPr>
                <p:cNvPr id="{pic1_id}" name="Picture A"/>
                <p:cNvPicPr/><p:nvPr/>
            </p:nvPicPr>
            <p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:pic>
        <p:pic>
            <p:nvPicPr>
                <p:cNvPr id="{pic2_id}" name="Picture B"/>
                <p:cNvPicPr/><p:nvPr/>
            </p:nvPicPr>
            <p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
            <p:spPr><a:xfrm><a:off x="500" y="500"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:pic>
    </p:grpSp>
    """


def build_nested_group_xml(outer_id=40, inner_id=41, text_id=42) -> str:
    return f"""
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr>
            <p:cNvPr id="{outer_id}" name="Outer Group {outer_id}"/>
            <p:cNvGrpSpPr/><p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="1000" y="1000"/>
                <a:ext cx="2000" cy="2000"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="200" cy="200"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:grpSp>
            <p:nvGrpSpPr>
                <p:cNvPr id="{inner_id}" name="Inner Group {inner_id}"/>
                <p:cNvGrpSpPr/><p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="10" y="10"/>
                    <a:ext cx="80" cy="80"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="800" cy="800"/>
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="{text_id}" name="Nested Text {text_id}"/>
                    <p:cNvSpPr txBox="1"/><p:nvPr/>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm><a:off x="100" y="100"/><a:ext cx="200" cy="200"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </p:spPr>
                <p:txBody>
                    <a:bodyPr/><a:lstStyle/>
                    <a:p><a:r><a:t>Deeply Nested Text</a:t></a:r></a:p>
                </p:txBody>
            </p:sp>
        </p:grpSp>
    </p:grpSp>
    """


def build_nested_mixed_group_xml(outer_id=50, text_id=51, inner_id=52, shape_id=53, pic_id=54) -> str:
    return f"""
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvGrpSpPr>
            <p:cNvPr id="{outer_id}" name="Outer Mixed {outer_id}"/>
            <p:cNvGrpSpPr/><p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="100" y="100"/>
                <a:ext cx="1000" cy="1000"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="1000" cy="1000"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="{text_id}" name="Top Text {text_id}"/>
                <p:cNvSpPr txBox="1"/><p:nvPr/>
            </p:nvSpPr>
            <p:spPr>
                <a:xfrm><a:off x="10" y="10"/><a:ext cx="200" cy="50"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </p:spPr>
            <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Top Level Group Text</a:t></a:r></a:p></p:txBody>
        </p:sp>
        <p:grpSp>
            <p:nvGrpSpPr>
                <p:cNvPr id="{inner_id}" name="Inner Mixed {inner_id}"/>
                <p:cNvGrpSpPr/><p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="200" y="200"/>
                    <a:ext cx="500" cy="500"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="500" cy="500"/>
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="{shape_id}" name="Inner Shape {shape_id}"/>
                    <p:cNvSpPr/><p:nvPr/>
                </p:nvSpPr>
                <p:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </p:spPr>
            </p:sp>
            <p:pic>
                <p:nvPicPr>
                    <p:cNvPr id="{pic_id}" name="Inner Pic {pic_id}"/>
                    <p:cNvPicPr/><p:nvPr/>
                </p:nvPicPr>
                <p:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
                <p:spPr><a:xfrm><a:off x="100" y="0"/><a:ext cx="200" cy="200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
            </p:pic>
        </p:grpSp>
    </p:grpSp>
    """


def create_slide_with_element(xml_str: str):
    """Creates a python-pptx presentation and appends the raw element to slide 1's spTree."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    elem = _make_xml(xml_str)
    slide.shapes._spTree.append(elem)
    return prs, slide, elem


# =============================================================================
# 1. Simple Group: Text + Rectangle + Shape -> all eligible children extracted
# =============================================================================
def test_01_simple_group_extraction():
    prs, slide, grp_el = create_slide_with_element(build_simple_group_xml(10, 11, 12, 13))
    transformer = SelectiveUngroupingTransformer()

    result: UngroupResult = transformer.ungroup_slide(slide._element, slide_number=1)

    assert result.success is True
    assert result.groups_processed == 1
    assert result.shapes_extracted == 3
    assert set(result.extracted_shape_ids) == {11, 12, 13}
    assert 10 in result.removed_group_ids

    # Verify no groups remain on slide, only extracted shapes in spTree
    remaining_groups = _safe_xpath(slide._element, ".//p:spTree/p:grpSp")
    assert len(remaining_groups) == 0

    sp_ids = [int(s.xpath("./p:nvSpPr/p:cNvPr")[0].get("id")) for s in _safe_xpath(slide._element, ".//p:spTree/p:sp")]
    assert 11 in sp_ids
    assert 12 in sp_ids
    assert 13 in sp_ids


# =============================================================================
# 2. Mixed Group: Text + Rectangle + Picture -> text/rect extracted, picture protected
# =============================================================================
def test_02_mixed_group_picture_protection():
    prs, slide, grp_el = create_slide_with_element(build_mixed_group_xml(20, 21, 22, 23))
    transformer = SelectiveUngroupingTransformer()

    result: UngroupResult = transformer.ungroup_slide(slide._element, slide_number=1)

    assert result.success is True
    assert result.groups_processed == 1
    assert result.shapes_extracted == 2
    assert result.pictures_protected == 1
    assert set(result.extracted_shape_ids) == {21, 22}
    assert 20 in result.retained_group_ids
    assert 20 not in result.removed_group_ids

    # Group 20 must still exist in spTree and contain Picture 23
    retained_groups = _safe_xpath(slide._element, ".//p:spTree/p:grpSp")
    assert len(retained_groups) == 1
    pic_in_grp = _safe_xpath(retained_groups[0], ".//p:pic")
    assert len(pic_in_grp) == 1
    assert int(pic_in_grp[0].xpath("./p:nvPicPr/p:cNvPr")[0].get("id")) == 23

    # Text and Rect must be at top-level spTree
    top_shapes = _safe_xpath(slide._element, ".//p:spTree/p:sp")
    top_ids = [int(s.xpath("./p:nvSpPr/p:cNvPr")[0].get("id")) for s in top_shapes]
    assert 21 in top_ids
    assert 22 in top_ids


# =============================================================================
# 3. Picture-only Group: Remains 100% untouched and protected
# =============================================================================
def test_03_picture_only_group_protected():
    prs, slide, grp_el = create_slide_with_element(build_pic_only_group_xml(30, 31, 32))
    classifier = GroupClassifier()
    analysis = classifier.analyze_group(grp_el, slide_number=1)

    assert analysis.category == GroupCategory.PICTURE_ONLY
    assert analysis.picture_count == 2
    assert analysis.eligible_count == 0
    assert analysis.is_safe_to_remediate is False

    transformer = SelectiveUngroupingTransformer()
    result = transformer.ungroup_slide(slide._element, slide_number=1)

    # Must NOT touch or dissolve picture-only group
    assert result.groups_processed == 0
    assert result.shapes_extracted == 0
    remaining_groups = _safe_xpath(slide._element, ".//p:spTree/p:grpSp")
    assert len(remaining_groups) == 1
    assert int(remaining_groups[0].xpath("./p:nvGrpSpPr/p:cNvPr")[0].get("id")) == 30


# =============================================================================
# 4. Nested Group: Group -> Group -> Text -> recursively processed safely
# =============================================================================
def test_04_nested_group_recursive_processing():
    prs, slide, grp_el = create_slide_with_element(build_nested_group_xml(40, 41, 42))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)

    assert result.success is True
    assert result.shapes_extracted == 1
    assert 42 in result.extracted_shape_ids
    # Both outer and inner groups dissolved as they contained no pictures
    assert 40 in result.removed_group_ids
    assert 41 in result.removed_group_ids

    # Deeply nested text shape 42 is now on slide spTree
    top_sp = _safe_xpath(slide._element, ".//p:spTree/p:sp")
    assert len(top_sp) == 1
    assert int(top_sp[0].xpath("./p:nvSpPr/p:cNvPr")[0].get("id")) == 42


# =============================================================================
# 5. Nested Mixed Group: Outer Text extracted, Inner Shape extracted, Pic protected
# =============================================================================
def test_05_nested_mixed_group_partial_extraction():
    prs, slide, grp_el = create_slide_with_element(build_nested_mixed_group_xml(50, 51, 52, 53, 54))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)

    assert result.success is True
    # Shapes 51 (outer text) and 53 (inner rect) extracted
    assert 51 in result.extracted_shape_ids
    assert 53 in result.extracted_shape_ids
    assert result.pictures_protected >= 1

    # Picture 54 must remain in the document inside a group
    pics = _safe_xpath(slide._element, ".//p:pic")
    assert len(pics) == 1
    assert int(pics[0].xpath("./p:nvPicPr/p:cNvPr")[0].get("id")) == 54


# =============================================================================
# 6. Shape ID Preservation: IDs never change
# =============================================================================
def test_06_shape_id_preservation():
    prs, slide, grp_el = create_slide_with_element(build_simple_group_xml(10, 101, 102, 103))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)

    for sid in [101, 102, 103]:
        sp_matches = _safe_xpath(slide._element, f".//p:spTree/p:sp[p:nvSpPr/p:cNvPr[@id='{sid}']]")
        assert len(sp_matches) == 1, f"Shape ID {sid} was not preserved!"


# =============================================================================
# 7. Duplicate ID Prevention: No duplicates created
# =============================================================================
def test_07_duplicate_id_prevention():
    prs, slide, grp_el = create_slide_with_element(build_simple_group_xml(10, 11, 12, 13))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Collect all IDs in spTree
    all_id_elements = _safe_xpath(slide._element, ".//p:spTree//*[@id]")
    ids = [el.get("id") for el in all_id_elements if el.tag.endswith("cNvPr")]
    assert len(ids) == len(set(ids)), f"Duplicate IDs detected: {ids}"


# =============================================================================
# 8. Relationship Preservation: r:embed unchanged
# =============================================================================
def test_08_relationship_preservation():
    prs, slide, grp_el = create_slide_with_element(build_mixed_group_xml(20, 21, 22, 23, r_id="rIdSpecial99"))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Check blip embed rel ID
    blip = _safe_xpath(slide._element, ".//a:blip")[0]
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    assert blip.get(f"{{{ns_r}}}embed") == "rIdSpecial99"


# =============================================================================
# 9. Hyperlink Preservation: Hyperlink on shape retained
# =============================================================================
def test_09_hyperlink_preservation():
    xml = """
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvGrpSpPr><p:cNvPr id="60" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr>
                <p:cNvPr id="61" name="LinkShape">
                    <a:hlinkClick r:id="rIdLink7" action="ppaction://hlinksldurl"/>
                </p:cNvPr>
                <p:cNvSpPr/><p:nvPr/>
            </p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="50" cy="50"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:sp>
    </p:grpSp>
    """
    prs, slide, grp_el = create_slide_with_element(xml)
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Verify hlinkClick node still exists on shape 61
    links = _safe_xpath(slide._element, ".//p:sp[p:nvSpPr/p:cNvPr[@id='61']]//a:hlinkClick")
    assert len(links) == 1
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    assert links[0].get(f"{{{ns_r}}}id") == "rIdLink7"


# =============================================================================
# 10. Coordinate Transformation Formula
# =============================================================================
def test_10_coordinate_transformation_formula():
    group_off = (100.0, 200.0)
    group_ext = (400.0, 600.0)
    group_ch_off = (0.0, 0.0)
    group_ch_ext = (200.0, 300.0)
    child_off = (50.0, 50.0)
    child_ext = (100.0, 100.0)

    # scale_x = 400/200 = 2.0; scale_y = 600/300 = 2.0
    # slide_x = 100 + (50 - 0)*2 = 200; slide_y = 200 + (50 - 0)*2 = 300
    # slide_w = 100 * 2 = 200; slide_h = 100 * 2 = 200
    sx, sy, sw, sh = calculate_group_child_slide_coordinates(
        group_off, group_ext, group_ch_off, group_ch_ext, child_off, child_ext
    )

    assert sx == 200.0
    assert sy == 300.0
    assert sw == 200.0
    assert sh == 200.0


# =============================================================================
# 11. Group Scaling: Correctly scales dimensions and offsets
# =============================================================================
def test_11_group_scaling():
    xml = """
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr><p:cNvPr id="70" name="Scaled Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="2000" cy="4000"/>
                <a:chOff x="0" y="0"/>
                <a:chExt cx="1000" cy="2000"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr><p:cNvPr id="71" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:sp>
    </p:grpSp>
    """
    prs, slide, grp_el = create_slide_with_element(xml)
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Scale is 2x in both axes
    # off: (0 + 100*2, 0 + 200*2) = (200, 400)
    # ext: (300*2, 400*2) = (600, 800)
    sp = _safe_xpath(slide._element, ".//p:sp[p:nvSpPr/p:cNvPr[@id='71']]")[0]
    xfrm = sp.xpath("./p:spPr/a:xfrm")[0]
    off = xfrm.xpath("./a:off")[0]
    ext = xfrm.xpath("./a:ext")[0]

    assert int(off.get("x")) == 200
    assert int(off.get("y")) == 400
    assert int(ext.get("cx")) == 600
    assert int(ext.get("cy")) == 800


# =============================================================================
# 12. Group Translation: Offsets shifted accordingly
# =============================================================================
def test_12_group_translation():
    xml = """
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr><p:cNvPr id="80" name="Shift Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr>
            <a:xfrm>
                <a:off x="500000" y="800000"/>
                <a:ext cx="100000" cy="100000"/>
                <a:chOff x="1000" y="1000"/>
                <a:chExt cx="100000" cy="100000"/>
            </a:xfrm>
        </p:grpSpPr>
        <p:sp>
            <p:nvSpPr><p:cNvPr id="81" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="2000" y="3000"/><a:ext cx="5000" cy="5000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:sp>
    </p:grpSp>
    """
    prs, slide, grp_el = create_slide_with_element(xml)
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # scale = 1.0; x = 500000 + (2000 - 1000) = 501000; y = 800000 + (3000 - 1000) = 802000
    sp = _safe_xpath(slide._element, ".//p:sp[p:nvSpPr/p:cNvPr[@id='81']]")[0]
    off = sp.xpath("./p:spPr/a:xfrm/a:off")[0]
    assert int(off.get("x")) == 501000
    assert int(off.get("y")) == 802000


# =============================================================================
# 13. Nested Group Coordinate Transformation
# =============================================================================
def test_13_nested_group_coordinate_transformation():
    # Stack: inner group inside outer group
    # Outer: off=(100, 100), ext=(200, 200), chOff=(0, 0), chExt=(100, 100) -> 2x scale
    # Inner: off=(10, 10), ext=(50, 50), chOff=(0, 0), chExt=(100, 100) -> 0.5x scale
    stack = [
        ((10.0, 10.0), (50.0, 50.0), (0.0, 0.0), (100.0, 100.0)),  # inner
        ((100.0, 100.0), (200.0, 200.0), (0.0, 0.0), (100.0, 100.0)),  # outer
    ]
    child_off = (20.0, 20.0)
    child_ext = (40.0, 40.0)

    # First inner: x = 10 + 20*0.5 = 20; y = 10 + 20*0.5 = 20; w = 40*0.5 = 20; h = 20
    # Then outer: x = 100 + 20*2 = 140; y = 100 + 20*2 = 140; w = 20*2 = 40; h = 40
    sx, sy, sw, sh = apply_nested_group_transforms(stack, child_off, child_ext)
    assert sx == 140.0
    assert sy == 140.0
    assert sw == 40.0
    assert sh == 40.0


# =============================================================================
# 14. Z-Order Preservation: Extracted shapes inserted at group position
# =============================================================================
def test_14_z_order_preservation():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sp_tree = slide.shapes._spTree

    # Append Shape Before (id 1)
    sp_before = _make_xml("""
        <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:nvSpPr><p:cNvPr id="1" name="Before"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:sp>
    """)
    sp_tree.append(sp_before)

    # Append Group (id 10) with children 11, 12
    grp = _make_xml(build_simple_group_xml(10, 11, 12, 13))
    sp_tree.append(grp)

    # Append Shape After (id 2)
    sp_after = _make_xml("""
        <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:nvSpPr><p:cNvPr id="2" name="After"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:sp>
    """)
    sp_tree.append(sp_after)

    transformer = SelectiveUngroupingTransformer()
    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Order in spTree children must be: Shape 1, Shape 11, Shape 12, Shape 13, Shape 2
    child_ids = []
    for child in sp_tree:
        if child.tag.split("}")[-1] in ("nvGrpSpPr", "grpSpPr"):
            continue
        cNvPr = child.xpath(".//p:cNvPr")
        if cNvPr:
            child_ids.append(int(cNvPr[0].get("id")))

    assert child_ids == [1, 11, 12, 13, 2]


# =============================================================================
# 15. Animation Target ID Preservation: Directly animated groups are skipped
# =============================================================================
def test_15_animation_target_preservation():
    prs, slide, grp_el = create_slide_with_element(build_simple_group_xml(90, 91, 92, 93))
    classifier = GroupClassifier()

    # Pass group ID 90 as targeted by animation
    analysis = classifier.analyze_group(grp_el, slide_number=1, animation_target_ids=[90])

    assert analysis.is_animation_target is True
    assert analysis.is_safe_to_remediate is False
    assert "animation" in analysis.skip_reason.lower()

    # Transformer must respect skip and leave group intact
    transformer = SelectiveUngroupingTransformer()
    result = transformer.ungroup_slide(slide._element, slide_number=1, animation_target_ids=[90])

    assert result.groups_processed == 0
    assert 90 in result.skipped_group_ids
    # Group still in spTree
    assert len(_safe_xpath(slide._element, ".//p:spTree/p:grpSp")) == 1


# =============================================================================
# 16. Unsupported/Ambiguous Group Fails Safely
# =============================================================================
def test_16_unsupported_group_fails_safely():
    # Group missing <a:xfrm> transform entirely
    bad_xml = """
    <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvGrpSpPr><p:cNvPr id="999" name="Bad Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
        <p:grpSpPr/>
        <p:sp>
            <p:nvSpPr><p:cNvPr id="1000" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
            <p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        </p:sp>
    </p:grpSp>
    """
    prs, slide, grp_el = create_slide_with_element(bad_xml)
    classifier = GroupClassifier()
    analysis = classifier.analyze_group(grp_el, slide_number=1)

    assert analysis.is_safe_to_remediate is False
    assert "transform" in analysis.skip_reason.lower()

    transformer = SelectiveUngroupingTransformer()
    result = transformer.ungroup_slide(slide._element, slide_number=1)

    # Group must remain safely without crashing
    assert result.groups_processed == 0
    assert 999 in result.skipped_group_ids


# =============================================================================
# 17. One Failed Group Does Not Stop Other Groups
# =============================================================================
def test_17_isolated_failure_does_not_halt_pipeline():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sp_tree = slide.shapes._spTree

    # Group 1: Bad group missing xfrm (ID 800)
    bad_grp = _make_xml("""
        <p:grpSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:nvGrpSpPr><p:cNvPr id="800" name="Bad Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
            <p:grpSpPr/>
            <p:sp><p:nvSpPr><p:cNvPr id="801" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>
        </p:grpSp>
    """)
    sp_tree.append(bad_grp)

    # Group 2: Valid simple group (ID 810)
    good_grp = _make_xml(build_simple_group_xml(810, 811, 812, 813))
    sp_tree.append(good_grp)

    transformer = SelectiveUngroupingTransformer()
    result = transformer.ungroup_slide(slide._element, slide_number=1)

    assert result.success is True
    # Group 800 skipped, Group 810 processed
    assert 800 in result.skipped_group_ids
    assert 810 in result.removed_group_ids
    assert set(result.extracted_shape_ids) == {811, 812, 813}


# =============================================================================
# 18. Original Slide Count Unchanged
# =============================================================================
def test_18_slide_count_unchanged(tmp_path):
    prs = Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    s1.shapes._spTree.append(_make_xml(build_simple_group_xml(10, 11, 12, 13)))
    s2.shapes._spTree.append(_make_xml(build_mixed_group_xml(20, 21, 22, 23)))

    test_file = tmp_path / "test_slide_count.pptx"
    prs.save(test_file)

    reader = PresentationReader()
    model = reader.read(test_file)
    assert model.total_slides == 2

    rule = SelectiveUngroupingRule()
    for s in model.slides:
        issues = rule.evaluate(s, model)
        for issue in issues:
            rule.remediate_issue(s, model, issue)

    out_file = tmp_path / "test_slide_count_remediated.pptx"
    writer = PresentationWriter()
    writer.write(model, out_file)

    rem_model = reader.read(out_file)
    assert rem_model.total_slides == 2


# =============================================================================
# 19. Original Text Unchanged
# =============================================================================
def test_19_text_content_unchanged():
    prs, slide, grp_el = create_slide_with_element(build_simple_group_xml(10, 11, 12, 13))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Check text inside extracted shape 11
    sp = _safe_xpath(slide._element, ".//p:sp[p:nvSpPr/p:cNvPr[@id='11']]")[0]
    texts = [t.text for t in _safe_xpath(sp, ".//a:t")]
    assert "Original Header Text" in texts


# =============================================================================
# 20. Original Visual Properties Unchanged: Fills and Geometries
# =============================================================================
def test_20_visual_properties_preserved():
    prs, slide, grp_el = create_slide_with_element(build_simple_group_xml(10, 11, 12, 13))
    transformer = SelectiveUngroupingTransformer()

    result = transformer.ungroup_slide(slide._element, slide_number=1)
    assert result.success is True

    # Shape 12 had solidFill srgbClr val="FF0000" and prst="roundRect"
    sp12 = _safe_xpath(slide._element, ".//p:sp[p:nvSpPr/p:cNvPr[@id='12']]")[0]
    prst = sp12.xpath(".//a:prstGeom")[0].get("prst")
    assert prst == "roundRect"

    fill = sp12.xpath(".//a:solidFill/a:srgbClr")[0].get("val")
    assert fill == "FF0000"

    # Shape 13 had prst="diamond"
    sp13 = _safe_xpath(slide._element, ".//p:sp[p:nvSpPr/p:cNvPr[@id='13']]")[0]
    assert sp13.xpath(".//a:prstGeom")[0].get("prst") == "diamond"
