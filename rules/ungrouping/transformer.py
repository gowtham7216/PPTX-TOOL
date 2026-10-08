"""
SlideGuard Selective Ungrouping Transformer.
Executes physical in-place OpenXML node movement within <p:spTree>.
Calculates slide-level coordinates using verified DrawingML transform mathematics.
Extracts eligible text and shape elements while protecting picture groups.
"""

from typing import List, Dict, Tuple, Optional, Any, Set
from dataclasses import dataclass, field
import logging
from lxml import etree

from pptx_engine.utilities.geometry import (
    emu_to_points,
    points_to_emu,
    calculate_group_child_slide_coordinates,
)
from rules.ungrouping.classifier import GroupClassifier, GroupAnalysis, GroupCategory, ChildCategory

logger = logging.getLogger(__name__)

NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _safe_xpath(elem, xpath_expr: str) -> list:
    """Executes xpath query safely across python-pptx BaseOxmlElement and raw lxml Element."""
    try:
        return elem.xpath(xpath_expr)
    except TypeError:
        return elem.xpath(xpath_expr, namespaces=NSMAP)
    except Exception:
        try:
            return elem.xpath(xpath_expr, namespaces=NSMAP)
        except Exception:
            return []


@dataclass
class UngroupResult:
    """Detailed result of an ungrouping operation on a slide or presentation."""
    slide_number: int
    success: bool
    groups_processed: int = 0
    groups_skipped: int = 0
    shapes_extracted: int = 0
    pictures_protected: int = 0
    nested_groups_processed: int = 0
    extracted_shape_ids: List[int] = field(default_factory=list)
    retained_group_ids: List[int] = field(default_factory=list)
    removed_group_ids: List[int] = field(default_factory=list)
    skipped_group_ids: List[int] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


class SelectiveUngroupingTransformer:
    """
    Performs physical in-place OOXML transformations on PowerPoint group trees.
    Never deletes and recreates shapes: detaches existing XML nodes, updates their
    local DrawingML <a:xfrm> to compound slide coordinates, and re-parents them
    in <p:spTree> while preserving all shape IDs, styling, and animation targets.
    """

    def __init__(self, classifier: Optional[GroupClassifier] = None):
        self.classifier = classifier or GroupClassifier()

    def ungroup_slide(
        self,
        slide_elem,
        slide_number: int,
        animation_target_ids: Optional[List[int]] = None,
        approved_group_ids: Optional[List[int]] = None
    ) -> UngroupResult:
        """
        Executes selective ungrouping on all eligible groups within a slide's <p:spTree>.
        """
        result = UngroupResult(slide_number=slide_number, success=True)
        anim_targets = set(animation_target_ids or [])
        approved_set = set(approved_group_ids) if approved_group_ids is not None else None

        sp_tree = self._get_sp_tree(slide_elem)
        if sp_tree is None:
            result.success = False
            result.errors.append(f"Slide {slide_number}: No <p:spTree> found.")
            return result

        # Record original shape IDs for validation
        original_shape_ids = self._collect_all_shape_ids(sp_tree)

        # Find top-level <p:grpSp> elements
        top_groups = [child for child in sp_tree if child.tag.split("}")[-1] == "grpSp"]

        for grp_elem in top_groups:
            cnv_prs = _safe_xpath(grp_elem, ".//p:cNvPr")
            grp_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0

            # Filter if approval list provided
            if approved_set is not None and grp_id not in approved_set:
                result.groups_skipped += 1
                result.skipped_group_ids.append(grp_id)
                continue

            analysis = self.classifier.analyze_group(
                group_elem=grp_elem,
                slide_number=slide_number,
                animation_target_ids=list(anim_targets)
            )

            if not analysis.is_safe_to_remediate:
                result.groups_skipped += 1
                result.skipped_group_ids.append(grp_id)
                result.warnings.append(
                    f"Skipped Group {grp_id} ('{analysis.group_name}'): {analysis.skip_reason}"
                )
                continue

            # Execute transformation on this group
            try:
                self._transform_group(
                    sp_tree=sp_tree,
                    grp_elem=grp_elem,
                    analysis=analysis,
                    result=result
                )
                result.groups_processed += 1
            except Exception as e:
                logger.error("Failed to ungroup Group %d on Slide %d: %s", grp_id, slide_number, e, exc_info=True)
                result.warnings.append(f"Failed to ungroup Group {grp_id}: {str(e)}")
                result.skipped_group_ids.append(grp_id)
                result.groups_skipped += 1

        # Post-remediation verification
        safe, verify_errors = self.verify_integrity(sp_tree, original_shape_ids, anim_targets)
        if not safe:
            result.success = False
            result.errors.extend(verify_errors)

        return result

    def _transform_group(
        self,
        sp_tree,
        grp_elem,
        analysis: GroupAnalysis,
        result: UngroupResult
    ) -> None:
        """
        Transforms an individual group within sp_tree.
        """
        try:
            grp_idx = list(sp_tree).index(grp_elem)
        except ValueError:
            return

        extracted_nodes: List[Any] = []

        # Recursively process children
        self._extract_eligible_children_recursive(
            container_elem=grp_elem,
            ancestor_stack=[],
            extracted_nodes=extracted_nodes,
            result=result
        )

        # Check remaining visual children in grp_elem
        remaining_visual_children = self._get_visual_children(grp_elem)
        has_remaining_content = len(remaining_visual_children) > 0

        # Insert extracted nodes into sp_tree at group's original position in z-order
        for i, node in enumerate(extracted_nodes):
            sp_tree.insert(grp_idx + i, node)
            c_id = self._get_shape_id(node)
            if c_id:
                result.extracted_shape_ids.append(c_id)
            result.shapes_extracted += 1

        if not has_remaining_content:
            # All children extracted: remove now-empty group container
            sp_tree.remove(grp_elem)
            result.removed_group_ids.append(analysis.group_id)
            logger.info("Removed empty group container ID %d", analysis.group_id)
        else:
            # Protected pictures remain in group: keep container
            result.retained_group_ids.append(analysis.group_id)
            result.pictures_protected += len(remaining_visual_children)
            logger.info(
                "Preserved group container ID %d as protected picture group with %d items",
                analysis.group_id, len(remaining_visual_children)
            )

    def _extract_eligible_children_recursive(
        self,
        container_elem,
        ancestor_stack: List[Any],
        extracted_nodes: List[Any],
        result: UngroupResult
    ) -> None:
        """
        Recursively extracts eligible shapes from a group container,
        updating their coordinates to slide-level space.
        """
        # Get container's transform
        c_off_x, c_off_y, c_ext_cx, c_ext_cy, ch_x, ch_y, ch_cx, ch_cy = self._extract_group_xfrm(container_elem)
        current_transform = ((c_off_x, c_off_y), (c_ext_cx, c_ext_cy), (ch_x, ch_y), (ch_cx, ch_cy))
        current_stack = ancestor_stack + [current_transform]

        children = list(container_elem)
        for child_el in children:
            tag = child_el.tag.split("}")[-1]
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue

            if tag == "grpSp":
                # Nested group: recurse
                result.nested_groups_processed += 1
                self._extract_eligible_children_recursive(
                    container_elem=child_el,
                    ancestor_stack=current_stack,
                    extracted_nodes=extracted_nodes,
                    result=result
                )
                # If nested group has no visual children left, remove it from container
                if len(self._get_visual_children(child_el)) == 0:
                    container_elem.remove(child_el)
                    nested_id = self._get_shape_id(child_el)
                    if nested_id:
                        result.removed_group_ids.append(nested_id)

            elif tag == "pic" or self._has_blip_fill(child_el):
                # Protected picture element: stays in group
                continue

            elif tag in ("sp", "cxnSp", "graphicFrame"):
                # Eligible shape: calculate slide coordinates and extract
                self._promote_child_to_slide(child_el, tag, current_stack)
                container_elem.remove(child_el)
                extracted_nodes.append(child_el)

    def _promote_child_to_slide(
        self,
        child_el,
        tag: str,
        ancestor_stack: List[Any]
    ) -> None:
        """
        Rewrites child element's <a:xfrm> to compound slide-level coordinates.
        """
        # Extract child's local coordinates
        off_x, off_y, ext_cx, ext_cy = self._extract_shape_xfrm(child_el, tag)

        # Compute slide coordinates
        slide_x, slide_y, slide_cx, slide_cy = self._compute_compound_coords(
            (off_x, off_y), (ext_cx, ext_cy), ancestor_stack
        )

        # Locate or create <a:xfrm>
        xfrm = self._get_or_create_xfrm(child_el, tag)

        # Update <a:off>
        off_elems = _safe_xpath(xfrm, ".//a:off")
        if off_elems:
            off_elems[0].set("x", str(int(round(slide_x))))
            off_elems[0].set("y", str(int(round(slide_y))))
        else:
            etree.SubElement(xfrm, "{http://schemas.openxmlformats.org/drawingml/2006/main}off", x=str(int(round(slide_x))), y=str(int(round(slide_y))))

        # Update <a:ext>
        ext_elems = _safe_xpath(xfrm, ".//a:ext")
        if ext_elems:
            ext_elems[0].set("cx", str(int(round(slide_cx))))
            ext_elems[0].set("cy", str(int(round(slide_cy))))
        else:
            etree.SubElement(xfrm, "{http://schemas.openxmlformats.org/drawingml/2006/main}ext", cx=str(int(round(slide_cx))), cy=str(int(round(slide_cy))))

    def _compute_compound_coords(
        self,
        child_off: Tuple[int, int],
        child_ext: Tuple[int, int],
        ancestor_stack: List[Any]
    ) -> Tuple[int, int, int, int]:
        """Maps child coordinates from deepest group to outermost slide space in EMU."""
        curr_x, curr_y = child_off
        curr_w, curr_h = child_ext

        # Compound transforms from innermost group outwards
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

    def _get_or_create_xfrm(self, shape_el, tag: str):
        """Locates or creates <a:xfrm> element inside shape's property container."""
        if tag == "graphicFrame":
            xfrm_elems = _safe_xpath(shape_el, ".//p:xfrm")
            if xfrm_elems:
                return xfrm_elems[0]
            return etree.SubElement(shape_el, "{http://schemas.openxmlformats.org/presentationml/2006/main}xfrm")
        else:
            sp_pr_elems = _safe_xpath(shape_el, ".//p:spPr")
            if not sp_pr_elems:
                sp_pr = etree.SubElement(shape_el, "{http://schemas.openxmlformats.org/presentationml/2006/main}spPr")
            else:
                sp_pr = sp_pr_elems[0]

            xfrm_elems = _safe_xpath(sp_pr, ".//a:xfrm")
            if xfrm_elems:
                return xfrm_elems[0]
            return etree.SubElement(sp_pr, "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")

    def _has_blip_fill(self, child_el) -> bool:
        """Checks if a shape contains an image fill."""
        return bool(_safe_xpath(child_el, ".//a:blip") or _safe_xpath(child_el, ".//a:blipFill"))

    def _get_visual_children(self, grp_el) -> List[Any]:
        """Returns non-property child shapes of a group."""
        return [
            c for c in grp_el
            if c.tag.split("}")[-1] not in ("nvGrpSpPr", "grpSpPr")
        ]

    def _get_sp_tree(self, slide_elem):
        """Finds <p:spTree> inside slide element."""
        if slide_elem.tag.split("}")[-1] == "spTree":
            return slide_elem
        trees = _safe_xpath(slide_elem, ".//p:spTree")
        return trees[0] if trees else None

    def _get_shape_id(self, elem) -> Optional[int]:
        """Extracts shape ID from element's non-visual properties."""
        cnv_prs = _safe_xpath(elem, ".//p:cNvPr")
        if cnv_prs and cnv_prs[0].get("id"):
            try:
                return int(cnv_prs[0].get("id"))
            except ValueError:
                pass
        return None

    def _collect_all_shape_ids(self, sp_tree) -> Set[int]:
        """Collects set of all shape IDs within spTree."""
        return set(self._collect_all_shape_ids_list(sp_tree))

    def _collect_all_shape_ids_list(self, sp_tree) -> List[int]:
        """Collects list of all shape IDs within spTree (including duplicates if any exist)."""
        ids = []
        for child in sp_tree:
            if child.tag.split("}")[-1] in ("nvGrpSpPr", "grpSpPr"):
                continue
            for cnv in _safe_xpath(child, ".//p:cNvPr"):
                s_id = cnv.get("id")
                if s_id:
                    try:
                        ids.append(int(s_id))
                    except ValueError:
                        pass
        return ids

    def verify_integrity(
        self,
        sp_tree,
        original_shape_ids: Set[int],
        animation_target_ids: Set[int]
    ) -> Tuple[bool, List[str]]:
        """
        Validates post-remediation integrity:
        1. All animation target IDs still exist.
        2. No duplicate shape IDs were generated.
        """
        errors = []
        current_ids = self._collect_all_shape_ids(sp_tree)

        # 1. Animation target preservation
        for a_id in animation_target_ids:
            if a_id in original_shape_ids and a_id not in current_ids:
                errors.append(
                    f"Animation target shape ID {a_id} was removed or missing after ungrouping!"
                )

        # 2. Duplicate ID detection
        all_ids = self._collect_all_shape_ids_list(sp_tree)
        counts: Dict[int, int] = {}
        for s_id in all_ids:
            counts[s_id] = counts.get(s_id, 0) + 1

        duplicates = [s_id for s_id, count in counts.items() if count > 1]
        if duplicates:
            errors.append(f"Duplicate shape IDs detected after ungrouping: {duplicates}")

        return len(errors) == 0, errors
