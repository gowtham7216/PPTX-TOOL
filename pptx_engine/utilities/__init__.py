"""
SlideGuard PPTX Utilities Package.
"""

from pptx_engine.utilities.geometry import (
    EMU_PER_POINT,
    EMU_PER_INCH,
    emu_to_points,
    points_to_emu,
    is_in_footer_region,
    boxes_overlap,
    calculate_group_child_slide_coordinates,
    apply_nested_group_transforms,
)

__all__ = [
    "EMU_PER_POINT",
    "EMU_PER_INCH",
    "emu_to_points",
    "points_to_emu",
    "is_in_footer_region",
    "boxes_overlap",
    "calculate_group_child_slide_coordinates",
    "apply_nested_group_transforms",
]
