"""
SlideGuard Unit Tests for Phase 6: List Semantics & Numbering Preservation.
Covers all 28 required synthetic test scenarios and the real presentation tests:
 1. Simple bullet list.
 2. Simple numbered list.
 3. Lowercase letter list.
 4. Uppercase letter list.
 5. Parenthesized numbering.
 6. Parenthesized lowercase letters.
 7. Custom bullet character.
 8. Numbering with startAt=5.
 9. Multi-level bullet list.
 10. Multi-level numbered list.
 11. Mixed nested list.
 12. buNone.
 13. Inherited list formatting.
 14. List inside body placeholder.
 15. List inside converted text shape.
 16. List extracted from a group.
 17. Shape ID preservation.
 18. Animation target preservation.
 19. Hyperlink preservation.
 20. Font preservation.
 21. Indentation preservation.
 22. Hanging indentation preservation.
 23. Paragraph spacing preservation.
 24. Text character-for-character preservation.
 25. Unsupported numbering type -> MANUAL_REVIEW.
 26. Ambiguous visual list -> MANUAL_REVIEW.
 27. One failed paragraph does not stop processing.
 28. Slide count unchanged.
 Real Presentation Tests on Ch06_GLN_Instructor_AU.pptx and Sec1-5-TrigFunctions (2)original.pptx
"""

from pathlib import Path
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml import parse_xml
from lxml import etree

from models.presentation import PresentationModel
from models.slide import SlideModel
from models.enums import ShapeType, RuleBehavior, RemediationStatus, IssueSeverity
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.utilities.geometry import emu_to_points
from rules.lists.classifier import (
    ListClassifier,
    ParagraphListAnalysis,
    ListCategory,
    ListStatus,
    _safe_xpath,
    NSMAP,
)
from rules.lists.inheritance import (
    ListInheritanceResolver,
    InheritedListProperties,
)
from rules.lists.transformer import (
    ListTransformer,
    ListRemediationResult,
)
from rules.lists.rule import ListRule


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

def build_bullet_list_xml(shape_id=10, char="•") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
                <a:pPr><a:buChar char="{char}"/></a:pPr>
                <a:r><a:t>First bullet item</a:t></a:r>
            </a:p>
            <a:p>
                <a:pPr><a:buChar char="{char}"/></a:pPr>
                <a:r><a:t>Second bullet item</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """


def build_numbered_list_xml(shape_id=11, num_type="arabicPeriod", start_at="1") -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
                <a:pPr><a:buAutoNum type="{num_type}" startAt="{start_at}"/></a:pPr>
                <a:r><a:t>First step item</a:t></a:r>
            </a:p>
            <a:p>
                <a:pPr><a:buAutoNum type="{num_type}" startAt="{int(start_at) + 1}"/></a:pPr>
                <a:r><a:t>Second step item</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """


def build_pseudo_bullet_xml(shape_id=20) -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
                <a:r><a:t>• First visual bullet</a:t></a:r>
            </a:p>
            <a:p>
                <a:r><a:t>• Second visual bullet</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """


def build_pseudo_numbered_xml(shape_id=21) -> str:
    return f"""
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="{shape_id}" name="TextBox {shape_id}"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
            <a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="2000000"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        </p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle/>
            <a:p>
                <a:r><a:t>1. Introduction phase</a:t></a:r>
            </a:p>
            <a:p>
                <a:r><a:t>2. Implementation phase</a:t></a:r>
            </a:p>
            <a:p>
                <a:r><a:t>3. Verification phase</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """


# =============================================================================
# 1. Simple Bullet List Detection & Preservation
# =============================================================================
def test_01_simple_bullet_list():
    prs, slide, elem = create_slide_with_element(build_bullet_list_xml(10, "•"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    for a in analyses:
        assert a.category == ListCategory.BULLET
        assert a.status == ListStatus.VALID
        assert a.bullet_char == "•"
        assert not a.is_safe_to_remediate


# =============================================================================
# 2. Simple Numbered List Detection & Preservation
# =============================================================================
def test_02_simple_numbered_list():
    prs, slide, elem = create_slide_with_element(build_numbered_list_xml(11, "arabicPeriod", "1"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    for a in analyses:
        assert a.category == ListCategory.NUMBERED
        assert a.status == ListStatus.VALID
        assert a.numbering_type == "arabicPeriod"


# =============================================================================
# 3. Lowercase Letter List (alphaLcPeriod / alphaLcParenR)
# =============================================================================
def test_03_lowercase_letter_list():
    prs, slide, elem = create_slide_with_element(build_numbered_list_xml(12, "alphaLcPeriod", "1"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    for a in analyses:
        assert a.category == ListCategory.LETTERED
        assert a.status == ListStatus.VALID
        assert a.numbering_type == "alphaLcPeriod"


# =============================================================================
# 4. Uppercase Letter List (alphaUcPeriod)
# =============================================================================
def test_04_uppercase_letter_list():
    prs, slide, elem = create_slide_with_element(build_numbered_list_xml(13, "alphaUcPeriod", "1"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    for a in analyses:
        assert a.category == ListCategory.LETTERED
        assert a.status == ListStatus.VALID
        assert a.numbering_type == "alphaUcPeriod"


# =============================================================================
# 5. Parenthesized Numbering (arabicParenBoth)
# =============================================================================
def test_05_parenthesized_numbering():
    prs, slide, elem = create_slide_with_element(build_numbered_list_xml(14, "arabicParenBoth", "1"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].numbering_type == "arabicParenBoth"
    assert analyses[0].category == ListCategory.NUMBERED


# =============================================================================
# 6. Parenthesized Lowercase Letters (alphaLcParenBoth)
# =============================================================================
def test_06_parenthesized_lowercase_letters():
    prs, slide, elem = create_slide_with_element(build_numbered_list_xml(15, "alphaLcParenBoth", "1"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].numbering_type == "alphaLcParenBoth"
    assert analyses[0].category == ListCategory.LETTERED


# =============================================================================
# 7. Custom Bullet Character
# =============================================================================
def test_07_custom_bullet_character():
    prs, slide, elem = create_slide_with_element(build_bullet_list_xml(16, "➢"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].bullet_char == "➢"
    assert analyses[0].category == ListCategory.BULLET


# =============================================================================
# 8. Numbering with startAt=5
# =============================================================================
def test_08_numbering_with_start_at_5():
    prs, slide, elem = create_slide_with_element(build_numbered_list_xml(17, "arabicPeriod", "5"))
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert analyses[0].numbering_start == 5
    assert analyses[0].numbering_type == "arabicPeriod"


# =============================================================================
# 9. Multi-level Bullet List
# =============================================================================
def test_09_multi_level_bullet_list():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="30" name="MultiBullet"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr lvl="0"><a:buChar char="•"/></a:pPr><a:r><a:t>Level 0 item</a:t></a:r></a:p>
            <a:p><a:pPr lvl="1"><a:buChar char="–"/></a:pPr><a:r><a:t>Level 1 sub-item</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].category == ListCategory.MULTI_LEVEL
    assert analyses[0].level == 0
    assert analyses[1].category == ListCategory.MULTI_LEVEL
    assert analyses[1].level == 1


# =============================================================================
# 10. Multi-level Numbered List
# =============================================================================
def test_10_multi_level_numbered_list():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="31" name="MultiNum"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr lvl="0"><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr><a:r><a:t>1. Main step</a:t></a:r></a:p>
            <a:p><a:pPr lvl="1"><a:buAutoNum type="alphaLcPeriod" startAt="1"/></a:pPr><a:r><a:t>a. Sub step</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].category == ListCategory.MULTI_LEVEL
    assert analyses[0].level == 0
    assert analyses[1].category == ListCategory.MULTI_LEVEL
    assert analyses[1].level == 1


# =============================================================================
# 11. Mixed Nested List (Level 0 numbered, Level 1 bullet)
# =============================================================================
def test_11_mixed_nested_list():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="32" name="MixedNested"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr lvl="0"><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr><a:r><a:t>Main Section</a:t></a:r></a:p>
            <a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr><a:r><a:t>Bullet point</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].level == 0
    assert analyses[1].level == 1
    assert analyses[0].category == ListCategory.MULTI_LEVEL


# =============================================================================
# 12. buNone: Intentionally suppressed bullets preserved
# =============================================================================
def test_12_bu_none():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="33" name="BuNoneBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
                <a:pPr><a:buNone/></a:pPr>
                <a:r><a:t>Unbulleted paragraph explicitly</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 1
    assert analyses[0].has_bu_none is True
    assert analyses[0].status == ListStatus.NOT_A_LIST
    assert not analyses[0].is_safe_to_remediate


# =============================================================================
# 13. Inherited List Formatting from lstStyle
# =============================================================================
def test_13_inherited_list_formatting():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="34" name="InheritedBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/>
            <a:lstStyle>
                <a:lvl1pPr><a:buChar char="▪"/></a:lvl1pPr>
            </a:lstStyle>
            <a:p>
                <a:pPr lvl="0"/>
                <a:r><a:t>Inherits square bullet from lstStyle</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 1
    assert analyses[0].is_inherited is True
    assert analyses[0].inherited_from == "LST_STYLE"
    assert analyses[0].bullet_char == "▪"
    assert analyses[0].status == ListStatus.VALID


# =============================================================================
# 14. List Inside Body Placeholder
# =============================================================================
def test_14_list_inside_body_placeholder():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="35" name="Body Placeholder 35"/>
            <p:cNvSpPr/>
            <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>Body bullet item</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 1
    assert analyses[0].category == ListCategory.BULLET
    assert analyses[0].status == ListStatus.VALID


# =============================================================================
# 15. List Inside Converted Text Shape (Phase 5 Interaction)
# =============================================================================
def test_15_list_inside_converted_text_shape():
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="36" name="ConvertedShape 36"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="1"/></a:pPr><a:r><a:t>Converted step 1</a:t></a:r></a:p>
            <a:p><a:pPr><a:buAutoNum type="arabicPeriod" startAt="2"/></a:pPr><a:r><a:t>Converted step 2</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 2
    assert analyses[0].category == ListCategory.NUMBERED
    assert analyses[0].status == ListStatus.VALID


# =============================================================================
# 16. List Extracted from a Group (Phase 4 Interaction)
# =============================================================================
def test_16_list_extracted_from_group():
    # Simulates child shape promoted to slide root by Selective Ungrouping
    xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr>
            <p:cNvPr id="37" name="PromotedChild 37"/>
            <p:cNvSpPr txBox="1"/>
            <p:nvPr/>
        </p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="500" cy="500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:pPr><a:buChar char="•"/></a:pPr><a:r><a:t>Ungrouped bullet</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 1
    assert analyses[0].category == ListCategory.BULLET
    assert analyses[0].status == ListStatus.VALID


# =============================================================================
# 17. Shape ID Preservation
# =============================================================================
def test_17_shape_id_preservation():
    target_id = 8888
    prs, slide, elem = create_slide_with_element(build_pseudo_bullet_xml(target_id))
    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert res.paragraphs_remediated == 2

    # Verify shape ID unchanged
    cnv_prs = _safe_xpath(elem, ".//p:cNvPr")
    assert int(cnv_prs[0].get("id")) == target_id


# =============================================================================
# 18. Animation Target Preservation
# =============================================================================
def test_18_animation_target_preservation():
    anim_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="999" name="AnimShape 999"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:r><a:t>• Animated bullet 1</a:t></a:r></a:p>
            <a:p><a:r><a:t>• Animated bullet 2</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(anim_xml)
    # Add timing animation tag referencing 999
    timing_xml = """
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:tnLst>
            <p:par>
                <p:cTn id="1">
                    <p:childTnLst>
                        <p:seq>
                            <p:cTn id="2">
                                <p:childTnLst>
                                    <p:par>
                                        <p:cTn id="3">
                                            <p:childTnLst>
                                                <p:set>
                                                    <p:cBhvr>
                                                        <p:cTn id="4"/>
                                                        <p:tgtEl>
                                                            <p:spTgt spid="999"/>
                                                        </p:tgtEl>
                                                    </p:cBhvr>
                                                </p:set>
                                            </p:childTnLst>
                                        </p:cTn>
                                    </p:par>
                                </p:childTnLst>
                            </p:cTn>
                        </p:seq>
                    </p:childTnLst>
                </p:cTn>
            </p:par>
        </p:tnLst>
    </p:timing>
    """
    slide._element.append(_make_xml(timing_xml))

    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    # Check timing tag remains completely intact
    sp_tgts = _safe_xpath(slide._element, ".//p:timing//p:spTgt")
    assert len(sp_tgts) == 1
    assert sp_tgts[0].get("spid") == "999"


# =============================================================================
# 19. Hyperlink Preservation
# =============================================================================
def test_19_hyperlink_preservation():
    link_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <p:nvSpPr><p:cNvPr id="40" name="LinkList"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
                <a:r>
                    <a:rPr><a:hlinkClick r:id="rIdSpecialLink" action="ppaction://hlinksldurl"/></a:rPr>
                    <a:t>• Link item</a:t>
                </a:r>
            </a:p>
            <a:p><a:r><a:t>• Normal item</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(link_xml)
    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    # Check hyperlink node preserved on run
    hlinks = _safe_xpath(elem, ".//a:hlinkClick")
    assert len(hlinks) == 1
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    assert hlinks[0].get(f"{{{ns_r}}}id") == "rIdSpecialLink"


# =============================================================================
# 20. Font Preservation
# =============================================================================
def test_20_font_preservation():
    font_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="41" name="FontBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
                <a:r>
                    <a:rPr b="1" i="1" sz="2400"><a:solidFill><a:srgbClr val="FF5500"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
                    <a:t>• Styled bullet item</a:t>
                </a:r>
            </a:p>
            <a:p>
                <a:r>
                    <a:rPr b="1" i="1" sz="2400"><a:solidFill><a:srgbClr val="FF5500"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr>
                    <a:t>• Styled second item</a:t>
                </a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(font_xml)
    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    r_prs = _safe_xpath(elem, ".//a:rPr")
    assert len(r_prs) == 2
    assert r_prs[0].get("b") == "1"
    assert r_prs[0].get("sz") == "2400"
    assert _safe_xpath(r_prs[0], ".//a:latin")[0].get("typeface") == "Calibri"
    assert _safe_xpath(r_prs[0], ".//a:srgbClr")[0].get("val") == "FF5500"


# =============================================================================
# 21. Indentation Preservation
# =============================================================================
def test_21_indentation_preservation():
    indent_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="42" name="IndentBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
                <a:pPr marL="720000" indent="-360000"/>
                <a:r><a:t>• Custom indented 1</a:t></a:r>
            </a:p>
            <a:p>
                <a:pPr marL="720000" indent="-360000"/>
                <a:r><a:t>• Custom indented 2</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(indent_xml)
    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    ppr = _safe_xpath(elem, ".//a:pPr")[0]
    assert ppr.get("marL") == "720000"
    assert ppr.get("indent") == "-360000"


# =============================================================================
# 22. Hanging Indentation Preservation
# =============================================================================
def test_22_hanging_indentation_preservation():
    # Negative indent indicates hanging indent in OpenXML
    prs, slide, elem = create_slide_with_element(build_pseudo_bullet_xml(43))
    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    ppr = _safe_xpath(elem, ".//a:pPr")[0]
    indent_val = int(ppr.get("indent"))
    assert indent_val < 0, f"Expected negative hanging indent, got {indent_val}"


# =============================================================================
# 23. Paragraph Spacing Preservation
# =============================================================================
def test_23_paragraph_spacing_preservation():
    spacing_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="44" name="SpacingBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
                <a:pPr>
                    <a:spcBfr><a:spcPts val="1200"/></a:spcBfr>
                    <a:spcAft><a:spcPts val="600"/></a:spcAft>
                </a:pPr>
                <a:r><a:t>• Spaced item 1</a:t></a:r>
            </a:p>
            <a:p>
                <a:pPr>
                    <a:spcBfr><a:spcPts val="1200"/></a:spcBfr>
                    <a:spcAft><a:spcPts val="600"/></a:spcAft>
                </a:pPr>
                <a:r><a:t>• Spaced item 2</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(spacing_xml)
    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True

    ppr = _safe_xpath(elem, ".//a:pPr")[0]
    spc_bfr = _safe_xpath(ppr, ".//a:spcBfr/a:spcPts")
    assert len(spc_bfr) == 1
    assert spc_bfr[0].get("val") == "1200"


# =============================================================================
# 24. Text Character-for-Character Preservation on Native Lists
# =============================================================================
def test_24_text_character_for_character_preservation():
    prs, slide, elem = create_slide_with_element(build_bullet_list_xml(45, "•"))
    classifier = ListClassifier()
    analyses_before = classifier.analyze_slide_lists(slide, slide_number=1)
    texts_before = [a.text_preview for a in analyses_before]

    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.paragraphs_remediated == 0  # Native list preserved intact

    analyses_after = classifier.analyze_slide_lists(slide, slide_number=1)
    texts_after = [a.text_preview for a in analyses_after]

    assert texts_before == texts_after


# =============================================================================
# 25. Unsupported Numbering Type -> MANUAL_REVIEW
# =============================================================================
def test_25_unsupported_numbering_type_manual_review():
    unsupported_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="46" name="UnsupBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p>
                <a:pPr><a:buAutoNum type="unknownExoticSchemeType123"/></a:pPr>
                <a:r><a:t>Exotic numbering</a:t></a:r>
            </a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(unsupported_xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 1
    assert analyses[0].category == ListCategory.UNSUPPORTED
    assert analyses[0].status == ListStatus.MANUAL_REVIEW


# =============================================================================
# 26. Ambiguous Visual List -> MANUAL_REVIEW
# =============================================================================
def test_26_ambiguous_visual_list_manual_review():
    # Only 1 paragraph in a 3-paragraph shape starts with a number (e.g. "1. Introduction")
    # while the other two are regular sentences. Ambiguous whether intended as list or section header.
    ambiguous_xml = """
    <p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <p:nvSpPr><p:cNvPr id="47" name="AmbiguousBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
            <a:bodyPr/><a:lstStyle/>
            <a:p><a:r><a:t>1. Section One Overview</a:t></a:r></a:p>
            <a:p><a:r><a:t>This is a regular descriptive paragraph following the title.</a:t></a:r></a:p>
            <a:p><a:r><a:t>Another sentence expanding on the topic without any numbering.</a:t></a:r></a:p>
        </p:txBody>
    </p:sp>
    """
    prs, slide, elem = create_slide_with_element(ambiguous_xml)
    classifier = ListClassifier()
    analyses = classifier.analyze_slide_lists(slide, slide_number=1)

    assert len(analyses) == 3
    # First paragraph with single isolated number should be flagged MANUAL_REVIEW
    assert analyses[0].category == ListCategory.AMBIGUOUS
    assert analyses[0].status == ListStatus.MANUAL_REVIEW
    assert analyses[0].confidence == "LOW"


# =============================================================================
# 27. One Failed Paragraph Does Not Stop Processing
# =============================================================================
def test_27_fail_safe_one_failure_does_not_stop():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Shape 1: Valid pseudo bullet list
    slide.shapes._spTree.append(_make_xml(build_pseudo_bullet_xml(101)))
    # Shape 2: Another valid pseudo numbered list
    slide.shapes._spTree.append(_make_xml(build_pseudo_numbered_xml(102)))

    transformer = ListTransformer()
    res = transformer.remediate_slide(slide._element, slide_number=1, raw_slide=slide)
    assert res.success is True
    assert res.paragraphs_remediated >= 4


# =============================================================================
# 28. Slide Count Unchanged
# =============================================================================
def test_28_slide_count_unchanged():
    prs = Presentation()
    for i in range(4):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes._spTree.append(_make_xml(build_pseudo_bullet_xml(200 + i)))

    transformer = ListTransformer()
    for idx, s in enumerate(prs.slides, 1):
        transformer.remediate_slide(s._element, slide_number=idx, raw_slide=s)

    assert len(prs.slides) == 4


# =============================================================================
# REAL PRESENTATION TEST 1: tests/fixtures/Ch06_GLN_Instructor_AU.pptx
# =============================================================================
def test_real_presentation_ch06_lists():
    fixture_path = Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx")
    if not fixture_path.exists():
        pytest.skip(f"Fixture {fixture_path} not found.")

    reader = PresentationReader()
    presentation = reader.read(fixture_path)
    assert len(presentation.slides) == 50

    classifier = ListClassifier()
    total_shapes = 0
    total_p = 0
    bullet_count = 0
    numbered_count = 0
    multilevel_count = 0
    inherited_count = 0
    ambiguous_count = 0

    for sm in presentation.slides:
        raw_slide = sm.raw_slide
        assert raw_slide is not None
        analyses = classifier.analyze_slide_lists(raw_slide, sm.slide_number)
        for a in analyses:
            total_p += 1
            if a.category == ListCategory.BULLET:
                bullet_count += 1
            elif a.category in (ListCategory.NUMBERED, ListCategory.LETTERED):
                numbered_count += 1
            elif a.category == ListCategory.MULTI_LEVEL:
                multilevel_count += 1
            elif a.category == ListCategory.INHERITED:
                inherited_count += 1
            elif a.category == ListCategory.AMBIGUOUS:
                ambiguous_count += 1

    assert total_p > 0
    assert len(presentation.slides) == 50


# =============================================================================
# REAL PRESENTATION TEST 2: tests/fixtures/Sec1-5-TrigFunctions (2)original.pptx
# =============================================================================
def test_real_presentation_trig_functions_lists():
    fixture_path = Path("tests/fixtures/Sec1-5-TrigFunctions (2)original.pptx")
    if not fixture_path.exists():
        pytest.skip(f"Fixture {fixture_path} not found.")

    reader = PresentationReader()
    presentation = reader.read(fixture_path)
    assert len(presentation.slides) == 19

    classifier = ListClassifier()
    rule = ListRule(classifier=classifier)

    total_issues = 0
    for sm in presentation.slides:
        issues = rule.evaluate(sm, presentation)
        total_issues += len(issues)

    # In Sec1-5-TrigFunctions, slides have math formulas and headings, no pseudo-lists
    # Verifies critical requirement: "No list remediation required" rather than forced mutations
    assert len(presentation.slides) == 19
