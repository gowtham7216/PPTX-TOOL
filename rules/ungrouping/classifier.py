"""
SlideGuard Group Classifier.
Classifies PowerPoint OpenXML groups and their children into semantic categories
(pure text/shape, mixed picture + text/shape, picture-only, nested).
Determines safe eligibility for selective ungrouping while ensuring protected
picture structures are preserved.
"""

from enum import Enum
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Any
import logging

from pptx_engine.utilities.geometry import emu_to_points, calculate_group_child_slide_coordinates

logger = logging.getLogger(__name__)

NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _safe_xpath(elem, xpath_expr: str) -> list:
    """Executes xpath safely across python-pptx BaseOxmlElement and standard lxml Element."""
    try:
        return elem.xpath(xpath_expr)
    except TypeError:
        return elem.xpath(xpath_expr, namespaces=NSMAP)
    except Exception:
        try:
            return elem.xpath(xpath_expr, namespaces=NSMAP)
        except Exception:
            return []


class GroupCategory(str, Enum):
    """Categorization of a PowerPoint group container."""
    SHAPE_TEXT_ONLY = "SHAPE_TEXT_ONLY"  # Contains only text boxes, autoshapes, connectors (0 pictures)
    MIXED = "MIXED"                      # Contains text/autoshapes AND pictures (directly or in nested groups)
    PICTURE_ONLY = "PICTURE_ONLY"        # Contains exclusively pictures (100% protected)
    EMPTY = "EMPTY"                      # Container has 0 children


class ChildCategory(str, Enum):
    """Categorization of an individual child element within a group."""
    ELIGIBLE = "ELIGIBLE"                    # Text box, autoshape, connector, table eligible for extraction
    PROTECTED_PICTURE = "PROTECTED_PICTURE"  # p:pic element; must remain grouped and protected
    NESTED_GROUP = "NESTED_GROUP"            # Nested p:grpSp; requires recursive inspection
    OTHER = "OTHER"                          # Graphic frame or unknown element


@dataclass
class GroupChildInfo:
    """Detailed structural information about an individual child of a group."""
    shape_id: int
    name: str
    xml_tag: str
    category: ChildCategory
    is_picture: bool
    is_group: bool
    has_text: bool
    text_preview: str
    is_animated: bool
    raw_element: Any
    local_off_emu: Tuple[int, int]
    local_ext_emu: Tuple[int, int]
    slide_bbox_emu: Tuple[int, int, int, int]
    children: List["GroupChildInfo"] = field(default_factory=list)

    @property
    def slide_bbox_pt(self) -> Tuple[float, float, float, float]:
        return (
            emu_to_points(self.slide_bbox_emu[0]),
            emu_to_points(self.slide_bbox_emu[1]),
            emu_to_points(self.slide_bbox_emu[2]),
            emu_to_points(self.slide_bbox_emu[3]),
        )


@dataclass
class GroupAnalysis:
    """Complete analysis and classification result for a PowerPoint group shape."""
    group_id: int
    group_name: str
    slide_number: int
    category: GroupCategory
    is_animated: bool
    nesting_depth: int
    parent_group_id: Optional[int]
    total_children: int
    eligible_count: int
    picture_count: int
    nested_group_count: int
    children: List[GroupChildInfo]
    is_safe_to_remediate: bool
    confidence: str  # "HIGH", "MEDIUM", "LOW"
    skip_reason: Optional[str] = None
    group_element: Any = None
    group_off_emu: Tuple[int, int] = (0, 0)
    group_ext_emu: Tuple[int, int] = (0, 0)
    group_ch_off_emu: Tuple[int, int] = (0, 0)
    group_ch_ext_emu: Tuple[int, int] = (0, 0)

    @property
    def is_animation_target(self) -> bool:
        return self.is_animated

    def to_dict(self) -> Dict[str, Any]:
        return {
            "group_id": self.group_id,
            "group_name": self.group_name,
            "slide_number": self.slide_number,
            "category": self.category.value,
            "is_animated": self.is_animated,
            "nesting_depth": self.nesting_depth,
            "parent_group_id": self.parent_group_id,
            "total_children": self.total_children,
            "eligible_count": self.eligible_count,
            "picture_count": self.picture_count,
            "nested_group_count": self.nested_group_count,
            "is_safe_to_remediate": self.is_safe_to_remediate,
            "confidence": self.confidence,
            "skip_reason": self.skip_reason,
            "child_ids": [c.shape_id for c in self.children],
        }


class GroupClassifier:
    """
    Deterministic analyzer and classifier for PowerPoint group elements.
    Recursively inspects group hierarchies, detects embedded pictures and blips,
    and assigns safe remediation classifications.
    """

    def analyze_group(
        self,
        group_elem,
        slide_number: int,
        animation_target_ids: Optional[List[int]] = None,
        nesting_depth: int = 0,
        parent_group_id: Optional[int] = None,
        ancestor_stack: Optional[List[Tuple[Tuple[int, int], Tuple[int, int], Tuple[int, int], Tuple[int, int]]]] = None
    ) -> GroupAnalysis:
        anim_targets = set(animation_target_ids or [])
        stack = list(ancestor_stack or [])

        # Extract group non-visual properties
        cnv_prs = _safe_xpath(group_elem, ".//p:cNvPr")
        group_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
        group_name = cnv_prs[0].get("name", f"Group_{group_id}") if cnv_prs else f"Group_{group_id}"
        is_animated = group_id in anim_targets

        # Extract group transform (a:xfrm)
        off_x, off_y, ext_cx, ext_cy, ch_x, ch_y, ch_cx, ch_cy = self._extract_group_xfrm(group_elem)
        current_grp_transform = ((off_x, off_y), (ext_cx, ext_cy), (ch_x, ch_y), (ch_cx, ch_cy))
        child_ancestor_stack = stack + [current_grp_transform]

        # Inspect immediate children of <p:grpSp>
        child_infos: List[GroupChildInfo] = []
        for child_el in group_elem:
            tag = child_el.tag.split("}")[-1]
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue  # Non-visual and visual properties container, not shapes

            child_info = self._analyze_child_element(
                child_el=child_el,
                tag=tag,
                child_ancestor_stack=child_ancestor_stack,
                anim_targets=anim_targets,
                slide_number=slide_number,
                nesting_depth=nesting_depth + 1,
                parent_group_id=group_id
            )
            if child_info is not None:
                child_infos.append(child_info)

        # Count direct and recursive elements
        total_children = len(child_infos)
        pic_count = self._count_pictures(child_infos)
        eligible_count = self._count_eligible(child_infos)
        nested_group_count = sum(1 for c in child_infos if c.is_group)

        # Determine group category
        if total_children == 0:
            category = GroupCategory.EMPTY
        elif pic_count == 0:
            category = GroupCategory.SHAPE_TEXT_ONLY
        elif eligible_count == 0:
            category = GroupCategory.PICTURE_ONLY
        else:
            category = GroupCategory.MIXED

        # Determine safety and confidence
        is_safe, confidence, skip_reason = self._evaluate_safety(
            group_id=group_id,
            category=category,
            is_animated=is_animated,
            total_children=total_children,
            pic_count=pic_count,
            eligible_count=eligible_count,
            ch_cx=ch_cx,
            ch_cy=ch_cy
        )

        return GroupAnalysis(
            group_id=group_id,
            group_name=group_name,
            slide_number=slide_number,
            category=category,
            is_animated=is_animated,
            nesting_depth=nesting_depth,
            parent_group_id=parent_group_id,
            total_children=total_children,
            eligible_count=eligible_count,
            picture_count=pic_count,
            nested_group_count=nested_group_count,
            children=child_infos,
            is_safe_to_remediate=is_safe,
            confidence=confidence,
            skip_reason=skip_reason,
            group_element=group_elem,
            group_off_emu=(off_x, off_y),
            group_ext_emu=(ext_cx, ext_cy),
            group_ch_off_emu=(ch_x, ch_y),
            group_ch_ext_emu=(ch_cx, ch_cy)
        )

    def _analyze_child_element(
        self,
        child_el,
        tag: str,
        child_ancestor_stack: List[Any],
        anim_targets: set,
        slide_number: int,
        nesting_depth: int,
        parent_group_id: int
    ) -> Optional[GroupChildInfo]:
        """Analyzes an individual child XML node within a group."""
        cnv_prs = _safe_xpath(child_el, ".//p:cNvPr")
        shape_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
        name = cnv_prs[0].get("name", f"Shape_{shape_id}") if cnv_prs else f"Shape_{shape_id}"
        is_child_animated = shape_id in anim_targets

        # Extract child geometry in local coordinates
        c_off_x, c_off_y, c_ext_cx, c_ext_cy = self._extract_shape_xfrm(child_el, tag)

        # Compute slide-level coordinates through ancestor group stack
        slide_x, slide_y, slide_cx, slide_cy = self._compute_slide_coords(
            (c_off_x, c_off_y), (c_ext_cx, c_ext_cy), child_ancestor_stack
        )

        if tag == "grpSp":
            # Nested group: analyze recursively
            nested_analysis = self.analyze_group(
                group_elem=child_el,
                slide_number=slide_number,
                animation_target_ids=list(anim_targets),
                nesting_depth=nesting_depth,
                parent_group_id=parent_group_id,
                ancestor_stack=child_ancestor_stack
            )
            has_text = any(c.has_text for c in nested_analysis.children)
            return GroupChildInfo(
                shape_id=shape_id,
                name=name,
                xml_tag=f"p:{tag}",
                category=ChildCategory.NESTED_GROUP,
                is_picture=False,
                is_group=True,
                has_text=has_text,
                text_preview=f"Nested Group ({len(nested_analysis.children)} children)",
                is_animated=is_child_animated,
                raw_element=child_el,
                local_off_emu=(c_off_x, c_off_y),
                local_ext_emu=(c_ext_cx, c_ext_cy),
                slide_bbox_emu=(slide_x, slide_y, slide_cx, slide_cy),
                children=nested_analysis.children
            )

        elif tag == "pic":
            # Picture element: protected
            return GroupChildInfo(
                shape_id=shape_id,
                name=name,
                xml_tag=f"p:{tag}",
                category=ChildCategory.PROTECTED_PICTURE,
                is_picture=True,
                is_group=False,
                has_text=False,
                text_preview="",
                is_animated=is_child_animated,
                raw_element=child_el,
                local_off_emu=(c_off_x, c_off_y),
                local_ext_emu=(c_ext_cx, c_ext_cy),
                slide_bbox_emu=(slide_x, slide_y, slide_cx, slide_cy)
            )

        elif tag in ("sp", "cxnSp"):
            # Check for embedded picture / blipFill
            blips = _safe_xpath(child_el, ".//a:blip") or _safe_xpath(child_el, ".//a:blipFill")
            if blips:
                # Shape with image texture/fill acting as picture
                return GroupChildInfo(
                    shape_id=shape_id,
                    name=name,
                    xml_tag=f"p:{tag}",
                    category=ChildCategory.PROTECTED_PICTURE,
                    is_picture=True,
                    is_group=False,
                    has_text=False,
                    text_preview="",
                    is_animated=is_child_animated,
                    raw_element=child_el,
                    local_off_emu=(c_off_x, c_off_y),
                    local_ext_emu=(c_ext_cx, c_ext_cy),
                    slide_bbox_emu=(slide_x, slide_y, slide_cx, slide_cy)
                )

            # Auto-shape or Text Box: eligible for extraction
            text_preview = self._extract_text_preview(child_el)
            return GroupChildInfo(
                shape_id=shape_id,
                name=name,
                xml_tag=f"p:{tag}",
                category=ChildCategory.ELIGIBLE,
                is_picture=False,
                is_group=False,
                has_text=bool(text_preview.strip()),
                text_preview=text_preview[:60],
                is_animated=is_child_animated,
                raw_element=child_el,
                local_off_emu=(c_off_x, c_off_y),
                local_ext_emu=(c_ext_cx, c_ext_cy),
                slide_bbox_emu=(slide_x, slide_y, slide_cx, slide_cy)
            )

        elif tag == "graphicFrame":
            # Graphic frame (Table, Chart, Diagram)
            is_tbl = bool(_safe_xpath(child_el, ".//a:tbl"))
            return GroupChildInfo(
                shape_id=shape_id,
                name=name,
                xml_tag=f"p:{tag}",
                category=ChildCategory.ELIGIBLE,
                is_picture=False,
                is_group=False,
                has_text=is_tbl,
                text_preview="[Table Graphic Frame]" if is_tbl else "[Graphic Frame]",
                is_animated=is_child_animated,
                raw_element=child_el,
                local_off_emu=(c_off_x, c_off_y),
                local_ext_emu=(c_ext_cx, c_ext_cy),
                slide_bbox_emu=(slide_x, slide_y, slide_cx, slide_cy)
            )

        return None

    def _extract_group_xfrm(self, grp_el) -> Tuple[int, int, int, int, int, int, int, int]:
        """Extracts off, ext, chOff, and chExt from <p:grpSpPr/a:xfrm>."""
        xfrm = _safe_xpath(grp_el, ".//p:grpSpPr/a:xfrm")
        if not xfrm:
            return 0, 0, 0, 0, 0, 0, 0, 0

        x_node = xfrm[0]
        off = _safe_xpath(x_node, ".//a:off")
        ext = _safe_xpath(x_node, ".//a:ext")
        ch_off = _safe_xpath(x_node, ".//a:chOff")
        ch_ext = _safe_xpath(x_node, ".//a:chExt")

        off_x = int(off[0].get("x", "0")) if off else 0
        off_y = int(off[0].get("y", "0")) if off else 0
        ext_cx = int(ext[0].get("cx", "0")) if ext else 0
        ext_cy = int(ext[0].get("cy", "0")) if ext else 0
        ch_x = int(ch_off[0].get("x", "0")) if ch_off else off_x
        ch_y = int(ch_off[0].get("y", "0")) if ch_off else off_y
        ch_cx = int(ch_ext[0].get("cx", "0")) if ch_ext else ext_cx
        ch_cy = int(ch_ext[0].get("cy", "0")) if ch_ext else ext_cy

        return off_x, off_y, ext_cx, ext_cy, ch_x, ch_y, ch_cx, ch_cy

    def _extract_shape_xfrm(self, shape_el, tag: str) -> Tuple[int, int, int, int]:
        """Extracts off and ext from shape's <a:xfrm> container."""
        if tag == "grpSp":
            xfrm = _safe_xpath(shape_el, ".//p:grpSpPr/a:xfrm")
        elif tag == "graphicFrame":
            xfrm = _safe_xpath(shape_el, ".//p:xfrm") or _safe_xpath(shape_el, ".//a:xfrm")
        else:
            xfrm = _safe_xpath(shape_el, ".//p:spPr/a:xfrm") or _safe_xpath(shape_el, ".//a:xfrm")

        if not xfrm:
            return 0, 0, 0, 0

        x_node = xfrm[0]
        off = _safe_xpath(x_node, ".//a:off")
        ext = _safe_xpath(x_node, ".//a:ext")

        off_x = int(off[0].get("x", "0")) if off else 0
        off_y = int(off[0].get("y", "0")) if off else 0
        ext_cx = int(ext[0].get("cx", "0")) if ext else 0
        ext_cy = int(ext[0].get("cy", "0")) if ext else 0

        return off_x, off_y, ext_cx, ext_cy

    def _compute_slide_coords(
        self,
        child_off: Tuple[int, int],
        child_ext: Tuple[int, int],
        ancestor_stack: List[Tuple[Tuple[int, int], Tuple[int, int], Tuple[int, int], Tuple[int, int]]]
    ) -> Tuple[int, int, int, int]:
        """Maps local coordinates from deepest group to outer slide coordinates in EMU."""
        curr_x, curr_y = child_off
        curr_w, curr_h = child_ext

        # Apply stack from innermost group outwards
        for grp_off, grp_ext, grp_ch_off, grp_ch_ext in reversed(ancestor_stack):
            curr_x, curr_y, curr_w, curr_h = calculate_group_child_slide_coordinates(
                group_off=grp_off,
                group_ext=grp_ext,
                group_ch_off=grp_ch_off,
                group_ch_ext=grp_ch_ext,
                child_off=(curr_x, curr_y),
                child_ext=(curr_w, curr_h)
            )

        return int(round(curr_x)), int(round(curr_y)), int(round(curr_w)), int(round(curr_h))

    def _extract_text_preview(self, elem) -> str:
        """Extracts plain text preview from <p:txBody>."""
        paras = _safe_xpath(elem, ".//a:p")
        lines = []
        for p in paras:
            p_text = "".join(r.text for r in _safe_xpath(p, ".//a:t") if r.text)
            if p_text:
                lines.append(p_text)
        return " ".join(lines)

    def _count_pictures(self, children: List[GroupChildInfo]) -> int:
        """Recursively counts all picture elements in child hierarchy."""
        count = 0
        for c in children:
            if c.is_picture:
                count += 1
            if c.is_group:
                count += self._count_pictures(c.children)
        return count

    def _count_eligible(self, children: List[GroupChildInfo]) -> int:
        """Recursively counts all eligible text/shape elements in child hierarchy."""
        count = 0
        for c in children:
            if c.category == ChildCategory.ELIGIBLE:
                count += 1
            if c.is_group:
                count += self._count_eligible(c.children)
        return count

    def _evaluate_safety(
        self,
        group_id: int,
        category: GroupCategory,
        is_animated: bool,
        total_children: int,
        pic_count: int,
        eligible_count: int,
        ch_cx: int,
        ch_cy: int
    ) -> Tuple[bool, str, Optional[str]]:
        """Determines whether this group is safe to auto-remediate."""
        if category == GroupCategory.PICTURE_ONLY:
            return False, "HIGH", "Picture-only group: preserved and protected from ungrouping."

        if category == GroupCategory.EMPTY:
            return False, "HIGH", "Empty group container."

        if ch_cx == 0 or ch_cy == 0:
            return False, "LOW", "Degenerate coordinate extent (chExt is zero); transform unsafe."

        if is_animated:
            # Animation directly references this group ID
            return (
                False,
                "MEDIUM",
                f"Group ID {group_id} is targeted by animation; container removal requires animation timing remediation (deferred)."
            )

        if eligible_count == 0:
            return False, "HIGH", "No accessible text or auto-shapes found to extract."

        return True, "HIGH", None
