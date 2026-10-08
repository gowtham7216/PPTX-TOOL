"""
SlideGuard Geometry & Spatial Utilities.
Provides coordinate transformations, overlap calculation, footer boundary detection,
and DrawingML group coordinate mapping.
"""

from typing import Tuple, List

EMU_PER_POINT = 12700
EMU_PER_INCH = 914400


def emu_to_points(emu_val: int) -> float:
    """Converts English Metric Units (EMU) to Points."""
    if emu_val is None:
        return 0.0
    return float(emu_val) / EMU_PER_POINT


def points_to_emu(pt_val: float) -> int:
    """Converts Points to English Metric Units (EMU)."""
    return int(round(pt_val * EMU_PER_POINT))


def is_in_footer_region(
    top_pt: float,
    height_pt: float,
    slide_height_pt: float,
    footer_threshold_ratio: float = 0.88
) -> bool:
    """
    Checks if a shape lies within the typical bottom footer zone of a slide.
    footer_threshold_ratio specifies the vertical start boundary (default 88% down the slide).
    """
    if slide_height_pt <= 0:
        return False
    threshold_top = slide_height_pt * footer_threshold_ratio
    return top_pt >= threshold_top


def boxes_overlap(
    box1: Tuple[float, float, float, float],
    box2: Tuple[float, float, float, float]
) -> bool:
    """
    Checks whether two bounding boxes overlap.
    Boxes are defined as (left, top, width, height).
    """
    l1, t1, w1, h1 = box1
    l2, t2, w2, h2 = box2

    r1, b1 = l1 + w1, t1 + h1
    r2, b2 = l2 + w2, t2 + h2

    # If one rectangle is on the left side of the other
    if l1 >= r2 or l2 >= r1:
        return False
    # If one rectangle is above the other
    if t1 >= b2 or t2 >= b1:
        return False

    return True


def calculate_group_child_slide_coordinates(
    group_off: Tuple[float, float],
    group_ext: Tuple[float, float],
    group_ch_off: Tuple[float, float],
    group_ch_ext: Tuple[float, float],
    child_off: Tuple[float, float],
    child_ext: Tuple[float, float]
) -> Tuple[float, float, float, float]:
    """
    Maps an OpenXML DrawingML child element's local coordinates within a group
    to the outer slide coordinate space.

    In DrawingML OOXML:
      - group_off: (x, y) offset of the group in parent space
      - group_ext: (cx, cy) bounding box extension of the group in parent space
      - group_ch_off: (chx, chy) origin of the child coordinate system
      - group_ch_ext: (chcx, chcy) extent of the child coordinate system
      - child_off: (cx_off, cy_off) offset of child in group's child coordinate system
      - child_ext: (cx_ext, cy_ext) extent of child in group's child coordinate system

    Transformation formula:
      scale_x = group_ext.cx / group_ch_ext.chcx
      scale_y = group_ext.cy / group_ch_ext.chcy
      slide_x = group_off.x + (child_off.cx_off - group_ch_off.chx) * scale_x
      slide_y = group_off.y + (child_off.cy_off - group_ch_off.chy) * scale_y
      slide_w = child_ext.cx_ext * scale_x
      slide_h = child_ext.cy_ext * scale_y
    """
    grp_x, grp_y = group_off
    grp_w, grp_h = group_ext
    ch_x, ch_y = group_ch_off
    ch_w, ch_h = group_ch_ext
    c_x, c_y = child_off
    c_w, c_h = child_ext

    scale_x = (grp_w / ch_w) if ch_w != 0 else 1.0
    scale_y = (grp_h / ch_h) if ch_h != 0 else 1.0

    slide_x = grp_x + (c_x - ch_x) * scale_x
    slide_y = grp_y + (c_y - ch_y) * scale_y
    slide_w = c_w * scale_x
    slide_h = c_h * scale_y

    return (slide_x, slide_y, slide_w, slide_h)


def apply_nested_group_transforms(
    group_transform_stack: List[Tuple[Tuple[float, float], Tuple[float, float], Tuple[float, float], Tuple[float, float]]],
    child_off: Tuple[float, float],
    child_ext: Tuple[float, float]
) -> Tuple[float, float, float, float]:
    """
    Recursively applies a stack of group transformations from deepest to top-level.
    group_transform_stack is ordered from deepest group to outer-most group.
    Each item is (group_off, group_ext, group_ch_off, group_ch_ext).
    """
    curr_off = child_off
    curr_ext = child_ext

    for grp_off, grp_ext, grp_ch_off, grp_ch_ext in group_transform_stack:
        mapped_x, mapped_y, mapped_w, mapped_h = calculate_group_child_slide_coordinates(
            grp_off, grp_ext, grp_ch_off, grp_ch_ext, curr_off, curr_ext
        )
        curr_off = (mapped_x, mapped_y)
        curr_ext = (mapped_w, mapped_h)

    return (curr_off[0], curr_off[1], curr_ext[0], curr_ext[1])
