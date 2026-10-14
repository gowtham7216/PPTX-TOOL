"""
tests/unit/test_outline_order.py

Unit tests for Outline Order detection, renumbering, and remediation
so that PowerPoint Outline View order matches Reading Order sequence.
"""

import pytest
from lxml import etree
from pathlib import Path

from rules.reading_order.outline_order import (
    get_slide_ph_idx_sequence,
    is_outline_order_monotonic,
    renumber_slide_outline_order,
    fix_outline_order_on_presentation,
)
from models.presentation import PresentationModel
from models.slide import SlideModel
from models.shape import ShapeModel
from models.job import JobContext
from models.enums import RemediationStatus
from engine.remediation.remediation_service import RemediationService


SLIDE_XML_OUT_OF_ORDER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr/>
            <p:grpSpPr/>
            <!-- Title with NO idx (must remain untouched) -->
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1"/>
                    <p:nvPr>
                        <p:ph type="title"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
            <!-- Body 1 with idx 24 -->
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Content Placeholder 24"/>
                    <p:nvPr>
                        <p:ph idx="24"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
            <!-- Body 2 with idx 25 -->
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Content Placeholder 25"/>
                    <p:nvPr>
                        <p:ph idx="25"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
            <!-- Subtitle or floating block copy-pasted with idx 1 (out of sequence!) -->
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="5" name="Subtitle 1"/>
                    <p:nvPr>
                        <p:ph idx="1"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>
"""

SLIDE_XML_IN_ORDER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Title 1"/>
                    <p:nvPr>
                        <p:ph type="title"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Content Placeholder 1"/>
                    <p:nvPr>
                        <p:ph idx="1"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Content Placeholder 2"/>
                    <p:nvPr>
                        <p:ph idx="2"/>
                    </p:nvPr>
                </p:nvSpPr>
            </p:sp>
        </p:spTree>
    </p:cSld>
</p:sld>
"""


def test_is_outline_order_monotonic():
    assert is_outline_order_monotonic([]) is True
    assert is_outline_order_monotonic([1]) is True
    assert is_outline_order_monotonic([1, 2, 3, 4]) is True
    assert is_outline_order_monotonic([1, 1, 2, 3]) is True
    assert is_outline_order_monotonic([24, 25, 1]) is False
    assert is_outline_order_monotonic([1, 3, 2]) is False


def test_get_slide_ph_idx_sequence():
    elem = etree.fromstring(SLIDE_XML_OUT_OF_ORDER.encode("utf-8"))
    seq = get_slide_ph_idx_sequence(elem)
    # Title has no idx, so seq should be [24, 25, 1]
    assert seq == [24, 25, 1]
    assert is_outline_order_monotonic(seq) is False


def test_renumber_slide_outline_order():
    elem = etree.fromstring(SLIDE_XML_OUT_OF_ORDER.encode("utf-8"))
    count = renumber_slide_outline_order(elem)
    assert count == 3

    new_seq = get_slide_ph_idx_sequence(elem)
    assert new_seq == [1, 2, 3]
    assert is_outline_order_monotonic(new_seq) is True

    # Verify title has no idx attribute added
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    title_ph = elem.xpath(".//p:ph[@type='title']", namespaces=ns)[0]
    assert title_ph.get("idx") is None

    # Verify companion cNvPr names were updated
    cnvprs = elem.xpath(".//p:cNvPr", namespaces=ns)
    names = [c.get("name") for c in cnvprs]
    assert names == ["Title 1", "Content Placeholder 1", "Content Placeholder 2", "Subtitle 3"]


def test_fix_outline_order_on_presentation():
    elem1 = etree.fromstring(SLIDE_XML_OUT_OF_ORDER.encode("utf-8"))
    elem2 = etree.fromstring(SLIDE_XML_IN_ORDER.encode("utf-8"))

    class MockRawSlide:
        def __init__(self, elem):
            self._element = elem

    slide1 = SlideModel(slide_id=1, slide_number=1, title="Slide 1")
    slide1.raw_slide = MockRawSlide(elem1)

    slide2 = SlideModel(slide_id=2, slide_number=2, title="Slide 2")
    slide2.raw_slide = MockRawSlide(elem2)

    pres = PresentationModel(file_path=Path("dummy.pptx"))
    pres.slides = [slide1, slide2]

    summary = fix_outline_order_on_presentation(pres)
    assert summary["slides_checked"] == 2
    assert summary["slides_fixed"] == 1
    assert summary["total_placeholders_renumbered"] == 3

    # Slide 1 is now in order
    assert is_outline_order_monotonic(get_slide_ph_idx_sequence(elem1)) is True
    # Slide 2 remains in order
    assert is_outline_order_monotonic(get_slide_ph_idx_sequence(elem2)) is True


def test_remediation_service_outline_order(tmp_path):
    elem = etree.fromstring(SLIDE_XML_OUT_OF_ORDER.encode("utf-8"))

    class MockRawSlide:
        def __init__(self, elem):
            self._element = elem

    slide = SlideModel(slide_id=1, slide_number=1, title="Slide 1")
    slide.raw_slide = MockRawSlide(elem)

    pres = PresentationModel(file_path=tmp_path / "test.pptx")
    pres.slides = [slide]

    context = JobContext(
        job_id="test_outline_job",
        original_file=tmp_path / "test.pptx",
        workspace_dir=tmp_path,
        working_file=tmp_path / "test.pptx",
    )
    context.user_approved_remediation = True

    service = RemediationService()
    actions = service.remediate_outline_order(context, pres)

    assert len(actions) == 2  # 1 slide action + 1 overall action
    applied = [a for a in actions if a.status == RemediationStatus.APPLIED]
    assert len(applied) == 2
    assert actions[0].rule_id == "PIPELINE_OUTLINE_ORDER"
    assert "renumbered" in actions[0].description
