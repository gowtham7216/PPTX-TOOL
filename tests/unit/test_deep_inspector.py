"""
Unit tests for SlideGuard OOXMLDeepInspector and DrawingML coordinate transformations.
"""

from pathlib import Path
from pptx_engine.utilities.geometry import (
    calculate_group_child_slide_coordinates,
    apply_nested_group_transforms,
)
from pptx_engine.ooxml.deep_inspector import OOXMLDeepInspector
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.verification.verification_service import VerificationService
from engine.comparison.comparison_service import ComparisonService


def test_group_coordinate_transformation_math():
    """
    Validates exact DrawingML coordinate transformation mathematics.
    Formula:
      scale_x = grp_w / ch_w
      scale_y = grp_h / ch_h
      slide_x = grp_x + (child_x - ch_x) * scale_x
      slide_y = grp_y + (child_y - ch_y) * scale_y
      slide_w = child_w * scale_x
      slide_h = child_h * scale_y
    """
    grp_off = (100.0, 200.0)
    grp_ext = (300.0, 400.0)
    grp_ch_off = (50.0, 50.0)
    grp_ch_ext = (150.0, 200.0)
    child_off = (75.0, 100.0)
    child_ext = (50.0, 50.0)

    slide_x, slide_y, slide_w, slide_h = calculate_group_child_slide_coordinates(
        group_off=grp_off,
        group_ext=grp_ext,
        group_ch_off=grp_ch_off,
        group_ch_ext=grp_ch_ext,
        child_off=child_off,
        child_ext=child_ext,
    )

    # Expected: scale_x = 2.0, scale_y = 2.0
    # slide_x = 100 + (75 - 50) * 2 = 150.0
    # slide_y = 200 + (100 - 50) * 2 = 300.0
    # slide_w = 50 * 2 = 100.0
    # slide_h = 50 * 2 = 100.0
    assert slide_x == 150.0
    assert slide_y == 300.0
    assert slide_w == 100.0
    assert slide_h == 100.0


def test_nested_group_transforms():
    """Validates multi-level nested group coordinate transformation."""
    # Outer group
    outer_frame = (
        (100.0, 100.0),  # off
        (400.0, 400.0),  # ext
        (0.0, 0.0),      # chOff
        (200.0, 200.0),  # chExt (scale = 2.0)
    )
    # Inner group (nested inside outer)
    inner_frame = (
        (50.0, 50.0),    # off in outer's child space
        (100.0, 100.0),  # ext in outer's child space
        (0.0, 0.0),      # chOff in inner's child space
        (100.0, 100.0),  # chExt (scale = 1.0)
    )

    child_off = (10.0, 10.0)
    child_ext = (20.0, 20.0)

    # Stack ordered from deepest to outer-most
    stack = [inner_frame, outer_frame]
    slide_x, slide_y, slide_w, slide_h = apply_nested_group_transforms(
        group_transform_stack=stack,
        child_off=child_off,
        child_ext=child_ext
    )

    # Inner transform:
    # child in outer space = (50 + (10 - 0)*1.0, 50 + (10 - 0)*1.0, 20*1.0, 20*1.0) = (60.0, 60.0, 20.0, 20.0)
    # Outer transform:
    # slide = (100 + (60 - 0)*2.0, 100 + (60 - 0)*2.0, 20*2.0, 20*2.0) = (220.0, 220.0, 40.0, 40.0)
    assert slide_x == 220.0
    assert slide_y == 220.0
    assert slide_w == 40.0
    assert slide_h == 40.0


def test_deep_inspector_shape_tree_and_footers(sample_presentation_path):
    inspector = OOXMLDeepInspector()
    report = inspector.inspect_presentation(sample_presentation_path)

    assert report["total_slides"] == 6
    assert len(report["slides"]) == 6

    # Slide 1: Title slide with placeholders
    s1 = report["slides"][0]
    assert s1["slide_number"] == 1
    assert any(n["is_placeholder"] and n["placeholder_type"] in ("title", "ctrTitle") for n in s1["flat_shape_list"])

    # Slide 3: Footer candidate detection ("WILEY")
    s3 = report["slides"][2]
    footers = s3["footer_candidates"]
    assert len(footers) >= 1
    wiley_footer = [f for f in footers if "WILEY" in f.get("text_content", "")][0]
    assert wiley_footer["footer_candidate_type"] == "SPATIAL_BOTTOM_TEXT"

    # Slide 4: Table detection
    s4 = report["slides"][3]
    tables = [n for n in s4["flat_shape_list"] if n["is_table"]]
    assert len(tables) == 1


def test_deep_inspector_real_ch06_presentation():
    """Inspects the real-world 50-slide Ch06 presentation."""
    ch06_path = Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx")
    if not ch06_path.exists():
        return

    inspector = OOXMLDeepInspector()
    report = inspector.inspect_presentation(ch06_path)

    assert report["total_slides"] == 50

    # 1. Animations (<p:timing>) mapped
    anim_inventory = report["global_id_inventory"]
    assert anim_inventory["total_animated_shape_ids"] > 50
    # Shape 48 on Slide 4 is animated
    assert 48 in anim_inventory["animated_shape_ids"]

    # 2. Groups inspected
    total_groups = sum(len(s["group_transforms"]) for s in report["slides"])
    assert total_groups > 50

    # 3. List inspection: Slide 25 has (a), (b) alpha numbering
    s25 = report["slides"][24]
    autonum_items = [li for li in s25["list_items"] if li["autonum_type"] == "alphaLcParenBoth"]
    assert len(autonum_items) >= 2


def test_roundtrip_preservation_real_ch06(tmp_path):
    """
    Verifies that reading and writing the 50-slide real PPTX without modifications
    causes zero structural drift and preserves all shape IDs and animations.
    """
    ch06_path = Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx")
    if not ch06_path.exists():
        return

    reader = PresentationReader()
    writer = PresentationWriter()
    verifier = VerificationService()
    comparer = ComparisonService()

    orig_pres = reader.read(ch06_path)
    out_file = tmp_path / "roundtrip_test.pptx"

    # Write without modifications
    writer.write(orig_pres, out_file)

    # Verify container and slide count
    v_res = verifier.verify_file(out_file, expected_slide_count=50)
    assert v_res.passed is True

    # Read output and compare
    reloaded_pres = reader.read(out_file)
    c_res = comparer.compare(orig_pres, reloaded_pres)

    assert c_res.integrity_preserved is True
    assert c_res.slide_count_matches is True
    assert c_res.unintended_changes_count == 0
