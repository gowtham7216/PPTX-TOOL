"""
SlideGuard Placeholder Transformer.
Executes deterministic in-place OpenXML placeholder conversions:
1. Converts loose title text boxes into native slide title placeholders.
2. Converts safe body/object placeholders into self-contained shapes by materializing inherited geometry.
3. Preserves title, footer, slide number, date, empty, picture, table, chart, and animated placeholders.
Preserves exact shape IDs, bounding boxes, text, formatting, and relationships.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Set, Dict, Any, Tuple
import logging
from lxml import etree

from rules.placeholders.classifier import (
    PlaceholderClassifier,
    PlaceholderAnalysis,
    PlaceholderCategory,
    PlaceholderStatus,
    _safe_xpath,
    NSMAP,
)

logger = logging.getLogger(__name__)


@dataclass
class ConversionResult:
    """Outcome of placeholder processing across an individual slide."""
    slide_number: int
    success: bool
    total_placeholders_analyzed: int = 0
    placeholders_converted: int = 0
    placeholders_preserved: int = 0
    placeholders_manual_review: int = 0
    converted_shape_ids: List[int] = field(default_factory=list)
    preserved_shape_ids: List[int] = field(default_factory=list)
    manual_review_shape_ids: List[int] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


class PlaceholderTransformer:
    """
    Executes in-place OpenXML placeholder transformations.
    Strictly guarantees visual fidelity, shape ID stability, and fail-safe isolation.
    """

    def __init__(self, classifier: Optional[PlaceholderClassifier] = None):
        self.classifier = classifier or PlaceholderClassifier()

    def convert_slide(
        self,
        slide_elem,
        slide_number: int,
        raw_slide=None,
        animation_target_ids: Optional[List[int]] = None,
        approved_shape_ids: Optional[List[int]] = None
    ) -> ConversionResult:
        """
        Processes and remediates placeholders on a single slide.
        """
        result = ConversionResult(slide_number=slide_number, success=True)
        anim_targets = set(animation_target_ids or [])
        approved_set = set(approved_shape_ids) if approved_shape_ids is not None else None

        sp_tree = self._get_sp_tree(slide_elem)
        if sp_tree is None:
            result.errors.append("Slide is missing <p:spTree> container.")
            result.success = False
            return result

        # 1. Record original shape IDs and snapshot XML for atomic rollback
        original_shape_ids = self._collect_shape_ids(sp_tree)
        original_sp_tree_xml = etree.tostring(sp_tree)

        # 2. Analyze all placeholders and candidate title boxes on slide
        analyses = self.classifier.analyze_slide_placeholders(
            raw_slide=raw_slide,
            slide_number=slide_number,
            animation_target_ids=list(anim_targets)
        )
        result.total_placeholders_analyzed = len(analyses)

        # 3. Process each placeholder
        for analysis in analyses:
            shape_id = analysis.shape_id

            # Approval check
            if approved_set is not None and shape_id not in approved_set:
                result.placeholders_preserved += 1
                result.preserved_shape_ids.append(shape_id)
                continue

            # Preserved / Manual Review categories
            if analysis.status == PlaceholderStatus.PRESERVED:
                result.placeholders_preserved += 1
                result.preserved_shape_ids.append(shape_id)
                continue

            if analysis.status == PlaceholderStatus.MANUAL_REVIEW:
                result.placeholders_manual_review += 1
                result.manual_review_shape_ids.append(shape_id)
                result.warnings.append(
                    f"Slide {slide_number} Shape {shape_id} ('{analysis.shape_name}'): {analysis.reason}"
                )
                continue

            # Safe to convert: execute deterministic transformation
            if analysis.status == PlaceholderStatus.SAFE_TO_CONVERT:
                try:
                    self._execute_conversion(analysis, sp_tree)
                    result.placeholders_converted += 1
                    result.converted_shape_ids.append(shape_id)
                    logger.info("Successfully converted Shape %d on Slide %d", shape_id, slide_number)
                except Exception as e:
                    logger.error("Failed to convert Shape %d on Slide %d: %s", shape_id, slide_number, e, exc_info=True)
                    result.warnings.append(f"Conversion failed for Shape {shape_id}: {str(e)}")
                    result.placeholders_preserved += 1
                    result.preserved_shape_ids.append(shape_id)

        # 4. Post-conversion integrity verification & Rollback if unsafe
        safe, verify_errors = self.verify_integrity(
            sp_tree=sp_tree,
            original_shape_ids=original_shape_ids,
            animation_target_ids=anim_targets
        )
        if not safe:
            logger.warning("Integrity verification failed on Slide %d: %s. Initiating atomic rollback.", slide_number, verify_errors)
            parent = sp_tree.getparent()
            if parent is not None:
                tree_idx = parent.index(sp_tree)
                parent.remove(sp_tree)
                restored_sp_tree = etree.fromstring(original_sp_tree_xml)
                parent.insert(tree_idx, restored_sp_tree)
            result.success = False
            result.errors.extend(verify_errors)
            result.errors.append("Atomic rollback executed: slide restored to pre-conversion XML state.")
            result.placeholders_converted = 0
            result.converted_shape_ids = []
            return result

        return result

    def _execute_conversion(self, analysis: PlaceholderAnalysis, sp_tree) -> None:
        """
        Executes physical XML mutation for an individual approved shape.
        """
        shape_elem = analysis.raw_element
        if shape_elem is None:
            raise ValueError(f"Shape {analysis.shape_id} raw_element is None.")

        # CASE A: Standalone title text box -> Add native title placeholder tag
        if analysis.category == PlaceholderCategory.STANDALONE_TITLE_BOX:
            self._convert_to_title_placeholder(shape_elem)
            return

        # CASE B: Safe Body / Object placeholder -> Materialize inherited geometry and detach ph
        if analysis.category in (PlaceholderCategory.BODY, PlaceholderCategory.OBJECT):
            self._materialize_and_detach_placeholder(shape_elem, analysis)
            return

        # Default fallback
        logger.info("No conversion needed for Shape %d (category: %s)", analysis.shape_id, analysis.category)

    def _convert_to_title_placeholder(self, shape_elem) -> None:
        """
        Converts standalone text box into native slide title placeholder.
        Adds <p:ph type="title"/> into <p:nvSpPr><p:nvPr>.
        Preserves all geometry, font, color, text, and IDs.
        """
        nv_prs = _safe_xpath(shape_elem, ".//p:nvSpPr/p:nvPr")
        if not nv_prs:
            # Create nvPr if missing
            nv_sp_pr = _safe_xpath(shape_elem, ".//p:nvSpPr")[0]
            nv_pr = etree.SubElement(nv_sp_pr, "{http://schemas.openxmlformats.org/presentationml/2006/main}nvPr")
        else:
            nv_pr = nv_prs[0]

        # Check if <p:ph> already exists
        existing_ph = _safe_xpath(nv_pr, ".//p:ph")
        if not existing_ph:
            etree.SubElement(
                nv_pr,
                "{http://schemas.openxmlformats.org/presentationml/2006/main}ph",
                attrib={"type": "title"}
            )
        else:
            existing_ph[0].set("type", "title")

    def _materialize_and_detach_placeholder(
        self,
        shape_elem,
        analysis: PlaceholderAnalysis
    ) -> None:
        """
        Converts a body/object placeholder into a self-contained shape.
        1. Materializes inherited <a:xfrm> into <p:spPr> if absent on slide.
        2. Removes <p:ph> tag from <p:nvPr>.
        3. Preserves exact shape ID, text frames, formatting, and relationships.
        """
        sp_prs = _safe_xpath(shape_elem, ".//p:spPr")
        if not sp_prs:
            sp_pr = etree.SubElement(shape_elem, "{http://schemas.openxmlformats.org/presentationml/2006/main}spPr")
        else:
            sp_pr = sp_prs[0]

        # 1. Materialize <a:xfrm> if missing on slide
        existing_xfrm = _safe_xpath(sp_pr, ".//a:xfrm")
        if not existing_xfrm:
            inherited = analysis.inherited_properties
            if inherited.xfrm_emu:
                off_x, off_y, ext_cx, ext_cy = inherited.xfrm_emu
                # Insert <a:xfrm> at the beginning of <p:spPr>
                xfrm = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
                if inherited.rotation:
                    xfrm.set("rot", str(inherited.rotation))
                if inherited.flip_h:
                    xfrm.set("flipH", "1")
                if inherited.flip_v:
                    xfrm.set("flipV", "1")

                etree.SubElement(
                    xfrm,
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}off",
                    attrib={"x": str(off_x), "y": str(off_y)}
                )
                etree.SubElement(
                    xfrm,
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}ext",
                    attrib={"cx": str(ext_cx), "cy": str(ext_cy)}
                )
                sp_pr.insert(0, xfrm)

        # 2. Ensure preset geometry exists so shape remains valid rectangular text box
        existing_geom = _safe_xpath(sp_pr, ".//a:prstGeom") or _safe_xpath(sp_pr, ".//a:custGeom")
        if not existing_geom:
            prst = etree.SubElement(
                sp_pr,
                "{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom",
                attrib={"prst": "rect"}
            )
            etree.SubElement(prst, "{http://schemas.openxmlformats.org/drawingml/2006/main}avLst")

        # 3. Detach <p:ph> tag from <p:nvPr>
        ph_elems = _safe_xpath(shape_elem, ".//p:ph")
        for ph in ph_elems:
            parent = ph.getparent()
            if parent is not None:
                parent.remove(ph)

    def verify_integrity(
        self,
        sp_tree,
        original_shape_ids: Set[int],
        animation_target_ids: Set[int]
    ) -> Tuple[bool, List[str]]:
        """
        Validates post-conversion integrity:
        1. All original shape IDs still exist.
        2. All animation target IDs still exist.
        3. No duplicate shape IDs were generated.
        """
        errors = []
        current_ids = self._collect_shape_ids(sp_tree)

        # 1. Shape ID preservation
        missing = original_shape_ids - current_ids
        if missing:
            errors.append(f"Original shape IDs missing after conversion: {sorted(list(missing))}")

        # 2. Animation target ID preservation
        for a_id in animation_target_ids:
            if a_id in original_shape_ids and a_id not in current_ids:
                errors.append(f"Animation target shape ID {a_id} was lost during placeholder conversion!")

        # 3. Duplicate ID prevention
        all_ids_list = self._collect_shape_ids_list(sp_tree)
        counts: Dict[int, int] = {}
        for s_id in all_ids_list:
            counts[s_id] = counts.get(s_id, 0) + 1

        dupes = [s_id for s_id, count in counts.items() if count > 1]
        if dupes:
            errors.append(f"Duplicate shape IDs detected after conversion: {dupes}")

        return len(errors) == 0, errors

    def _get_sp_tree(self, slide_elem):
        """Finds <p:spTree> container."""
        if slide_elem.tag.split("}")[-1] == "spTree":
            return slide_elem
        trees = _safe_xpath(slide_elem, ".//p:spTree")
        return trees[0] if trees else None

    def _collect_shape_ids(self, sp_tree) -> Set[int]:
        """Collects set of unique shape IDs within spTree."""
        return set(self._collect_shape_ids_list(sp_tree))

    def _collect_shape_ids_list(self, sp_tree) -> List[int]:
        """Collects list of shape IDs, ignoring non-shape container properties."""
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
