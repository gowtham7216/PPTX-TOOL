"""
Regression tests for SlideGuard Originality & Design Preservation.
Verifies that positions, dimensions, text, and structure remain unchanged unless intentionally modified.
"""

from engine.comparison.comparison_service import ComparisonService
from pptx_engine.reader.presentation_reader import PresentationReader


def test_preservation_of_identical_presentation(sample_presentation_path):
    reader = PresentationReader()
    comparer = ComparisonService()

    orig = reader.read(sample_presentation_path)
    # Re-reading or comparing identical file
    copy = reader.read(sample_presentation_path)

    result = comparer.compare(orig, copy)

    assert result.slide_count_matches is True
    assert result.integrity_preserved is True
    assert result.unintended_changes_count == 0
    assert len(result.differences) == 0


def test_detection_of_unintended_movement(sample_presentation_path):
    reader = PresentationReader()
    comparer = ComparisonService()

    orig = reader.read(sample_presentation_path)
    mutated = reader.read(sample_presentation_path)

    # Artificially alter shape 1 position on slide 1
    s1 = mutated.get_slide_by_number(1)
    target_shape = s1.shapes[0]
    target_shape.left_pt += 15.0  # Shift by 15pt

    result = comparer.compare(orig, mutated)

    assert result.integrity_preserved is False
    assert result.unintended_changes_count >= 1
    diff = [d for d in result.differences if d.attribute == "position"][0]
    assert diff.is_intentional_accessibility_change is False
    assert diff.shape_id == target_shape.shape_id
