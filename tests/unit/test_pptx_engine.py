"""
Unit tests for SlideGuard PPTX Engine (Reader, Writer, Geometry, OOXML).
"""

from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from pptx_engine.utilities.geometry import (
    emu_to_points,
    points_to_emu,
    is_in_footer_region,
    boxes_overlap
)
from models.enums import ShapeType


def test_geometry_conversions():
    pt = 72.0
    emu = points_to_emu(pt)
    assert emu == 914400
    assert emu_to_points(emu) == 72.0


def test_footer_region_detection():
    # 7.5 inches slide = 540 points
    slide_height = 540.0
    # Top at 500pt (> 88% of 540 = 475.2pt) is inside footer region
    assert is_in_footer_region(500.0, 20.0, slide_height) is True
    # Top at 200pt is NOT inside footer region
    assert is_in_footer_region(200.0, 20.0, slide_height) is False


def test_boxes_overlap():
    b1 = (10.0, 10.0, 50.0, 50.0)
    b2 = (20.0, 20.0, 50.0, 50.0)
    b3 = (100.0, 100.0, 50.0, 50.0)

    assert boxes_overlap(b1, b2) is True
    assert boxes_overlap(b1, b3) is False


def test_presentation_reader(sample_presentation_path):
    reader = PresentationReader()
    presentation = reader.read(sample_presentation_path)

    assert presentation.total_slides == 6
    assert presentation.slide_width_pt == 720.0
    assert presentation.slide_height_pt == 540.0
    assert presentation.metadata.title == "SlideGuard Test Presentation"

    # Slide 1 has a title
    s1 = presentation.get_slide_by_number(1)
    assert s1 is not None
    assert s1.title == "Introduction to SlideGuard"

    # Slide 3 has footer detected ("WILEY")
    s3 = presentation.get_slide_by_number(3)
    assert s3 is not None
    assert s3.has_footer is True
    assert "WILEY" in (s3.footer_text or "")

    # Slide 4 has table
    s4 = presentation.get_slide_by_number(4)
    assert s4 is not None
    table_shapes = [s for s in s4.shapes if s.shape_type == ShapeType.TABLE]
    assert len(table_shapes) == 1
    assert table_shapes[0].table_row_count == 3
    assert table_shapes[0].table_col_count == 2


def test_presentation_writer(sample_presentation_path, tmp_path):
    reader = PresentationReader()
    writer = PresentationWriter()

    presentation = reader.read(sample_presentation_path)
    output_target = tmp_path / "saved_copy.pptx"

    writer.write(presentation, output_target)
    assert output_target.exists()
    assert output_target.stat().st_size > 0

    # Verify written copy can be reopened
    reopened = reader.read(output_target)
    assert reopened.total_slides == 6
