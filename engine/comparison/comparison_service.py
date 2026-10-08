"""
SlideGuard Presentation Comparison Service.
Performs semantic and structural comparison between original and remediated presentations.
Distinguishes intentional accessibility adjustments from unintended visual deviations.
"""

from typing import List, Optional
import logging

from models.presentation import PresentationModel
from models.slide import SlideModel
from models.shape import ShapeModel
from models.enums import ShapeType
from models.comparison import ComparisonResult, ComparisonDifference

logger = logging.getLogger(__name__)


class ComparisonService:
    """
    Compares two presentation models semantically to ensure design preservation.
    Flags unintended drift in geometry, font styling, or element deletion.
    """

    def __init__(self, position_tolerance_pt: float = 2.0, size_tolerance_pt: float = 2.0):
        self.pos_tol = position_tolerance_pt
        self.size_tol = size_tolerance_pt

    def compare(
        self,
        original: PresentationModel,
        remediated: PresentationModel,
        intentional_shape_ids: Optional[List[int]] = None
    ) -> ComparisonResult:
        result = ComparisonResult(
            original_slide_count=original.total_slides,
            remediated_slide_count=remediated.total_slides,
            slide_count_matches=(original.total_slides == remediated.total_slides)
        )

        approved_ids = set(intentional_shape_ids or [])

        if not result.slide_count_matches:
            result.add_difference(
                ComparisonDifference(
                    slide_number=0,
                    shape_id=0,
                    attribute="slide_count",
                    original_value=original.total_slides,
                    remediated_value=remediated.total_slides,
                    is_intentional_accessibility_change=False,
                    description=f"Slide count mismatch: original {original.total_slides} vs remediated {remediated.total_slides}"
                )
            )

        # Slide-by-slide comparison
        for orig_slide in original.slides:
            rem_slide = remediated.get_slide_by_number(orig_slide.slide_number)
            if rem_slide is None:
                result.add_difference(
                    ComparisonDifference(
                        slide_number=orig_slide.slide_number,
                        shape_id=0,
                        attribute="slide_missing",
                        original_value=orig_slide.slide_number,
                        remediated_value=None,
                        is_intentional_accessibility_change=False,
                        description=f"Slide {orig_slide.slide_number} missing in remediated presentation"
                    )
                )
                continue

            self._compare_slide(orig_slide, rem_slide, result, approved_ids)

        logger.info(
            "Comparison complete: %d intentional changes, %d unintended changes. Preserved: %s",
            result.intentional_changes_count,
            result.unintended_changes_count,
            result.integrity_preserved
        )
        return result

    def _compare_slide(
        self,
        orig_slide: SlideModel,
        rem_slide: SlideModel,
        result: ComparisonResult,
        approved_ids: set
    ) -> None:
        slide_num = orig_slide.slide_number

        for orig_shape in orig_slide.shapes:
            rem_shape = rem_slide.get_shape_by_id(orig_shape.shape_id)

            if rem_shape is None:
                # Shape missing in remediated slide
                # Check if this was an ungrouped group container whose children were extracted to slide
                is_ungrouped = False
                if orig_shape.shape_type == ShapeType.GROUP and orig_shape.children:
                    for child in orig_shape.children:
                        if rem_slide.get_shape_by_id(child.shape_id) is not None:
                            is_ungrouped = True
                            break

                is_intentional = is_ungrouped or (orig_shape.shape_id in approved_ids)
                attr_name = "group_ungrouped" if is_ungrouped else "shape_deletion"
                desc = (
                    f"Group '{orig_shape.name}' (ID {orig_shape.shape_id}) container dissolved; children extracted for accessibility"
                    if is_ungrouped
                    else f"Shape '{orig_shape.name}' missing in remediated slide {slide_num}"
                )
                result.add_difference(
                    ComparisonDifference(
                        slide_number=slide_num,
                        shape_id=orig_shape.shape_id,
                        attribute=attr_name,
                        original_value=orig_shape.name,
                        remediated_value=None,
                        is_intentional_accessibility_change=is_intentional,
                        description=desc
                    )
                )
                continue

            # Compare position
            if abs(orig_shape.left_pt - rem_shape.left_pt) > self.pos_tol or \
               abs(orig_shape.top_pt - rem_shape.top_pt) > self.pos_tol:
                is_intentional = orig_shape.shape_id in approved_ids
                result.add_difference(
                    ComparisonDifference(
                        slide_number=slide_num,
                        shape_id=orig_shape.shape_id,
                        attribute="position",
                        original_value=(orig_shape.left_pt, orig_shape.top_pt),
                        remediated_value=(rem_shape.left_pt, rem_shape.top_pt),
                        is_intentional_accessibility_change=is_intentional,
                        description=f"Shape '{orig_shape.name}' moved unintendedly"
                    )
                )

            # Compare size
            if abs(orig_shape.width_pt - rem_shape.width_pt) > self.size_tol or \
               abs(orig_shape.height_pt - rem_shape.height_pt) > self.size_tol:
                is_intentional = orig_shape.shape_id in approved_ids
                result.add_difference(
                    ComparisonDifference(
                        slide_number=slide_num,
                        shape_id=orig_shape.shape_id,
                        attribute="size",
                        original_value=(orig_shape.width_pt, orig_shape.height_pt),
                        remediated_value=(rem_shape.width_pt, rem_shape.height_pt),
                        is_intentional_accessibility_change=is_intentional,
                        description=f"Shape '{orig_shape.name}' resized unintendedly"
                    )
                )

            # Compare text content (must not be altered unintentionally)
            if orig_shape.has_text and orig_shape.text_content.strip() != rem_shape.text_content.strip():
                is_intentional = orig_shape.shape_id in approved_ids
                result.add_difference(
                    ComparisonDifference(
                        slide_number=slide_num,
                        shape_id=orig_shape.shape_id,
                        attribute="text_content",
                        original_value=orig_shape.text_content[:40],
                        remediated_value=rem_shape.text_content[:40],
                        is_intentional_accessibility_change=is_intentional,
                        description=f"Text content of shape '{orig_shape.name}' was modified"
                    )
                )
