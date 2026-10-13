"""
SlideGuard Placeholder Transformer.
Executes deterministic in-place OpenXML placeholder conversions:
1. Converts loose title text boxes into native slide title placeholders.
2. Converts loose subtitle text boxes into native slide subtitle placeholders.
3. Converts loose body text boxes into native slide body placeholders.
4. Materializes inherited geometry for safe body/object placeholders without visual drift.
5. Creates new native placeholders only when strictly approved under the 7 safety gates.
6. Preserves title, footer, slide number, date, empty, picture, table, chart, group, and animated placeholders.
Preserves exact shape IDs, bounding boxes, text, formatting, and relationships.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Set, Dict, Any, Tuple
import logging
from lxml import etree

from pptx_engine.utilities.geometry import emu_to_points
from rules.placeholders.classifier import (
    PlaceholderClassifier,
    PlaceholderAnalysis,
    PlaceholderCategory,
    PlaceholderStatus,
    ConversionAction,
    DifferenceCategory,
    _safe_xpath,
    NSMAP,
)
from rules.placeholders.placeholder_engine_v2 import PlaceholderEngineV2

try:
    import ConvertPlaceholder_v2 as cp_v2
except ImportError:
    import sys
    from pathlib import Path
    root_dir = str(Path(__file__).resolve().parents[2])
    if root_dir not in sys.path:
        sys.path.insert(0, root_dir)
    import ConvertPlaceholder_v2 as cp_v2

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
    created_shape_ids: List[int] = field(default_factory=list)
    preserved_shape_ids: List[int] = field(default_factory=list)
    manual_review_shape_ids: List[int] = field(default_factory=list)
    skipped_shape_ids: List[int] = field(default_factory=list)
    verified: bool = False
    verification_category: Optional[str] = None
    differences: List[Dict[str, Any]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    audit_records: List[Dict[str, Any]] = field(default_factory=list)


class PlaceholderTransformer:
    """
    Executes in-place OpenXML placeholder transformations.
    Strictly guarantees visual fidelity, shape ID stability, and fail-safe isolation.
    """

    def __init__(
        self,
        classifier: Optional[PlaceholderClassifier] = None,
        engine: Optional[PlaceholderEngineV2] = None,
    ):
        self.classifier = classifier or PlaceholderClassifier()
        self.engine = engine or PlaceholderEngineV2()

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

        # 1. Record original shape IDs, text, geometry, and snapshot XML for atomic rollback
        original_shape_ids = self._collect_shape_ids(sp_tree)
        original_shape_ids_list = self._collect_shape_ids_list(sp_tree)
        original_sp_tree_xml = etree.tostring(sp_tree)
        original_text_map = self._collect_text_map(sp_tree)
        original_xfrm_map = self._collect_xfrm_map(sp_tree)
        original_group_count = len(_safe_xpath(sp_tree, ".//p:grpSp"))
        original_rel_ids = self._collect_relationship_ids(sp_tree)

        # 2. Analyze all placeholders and candidate shapes on slide
        analyses = self.classifier.analyze_slide_placeholders(
            raw_slide=raw_slide,
            slide_number=slide_number,
            animation_target_ids=list(anim_targets)
        )
        result.total_placeholders_analyzed = len(analyses)

        # 3. Process each candidate
        for analysis in analyses:
            shape_id = analysis.shape_id
            result.audit_records.append(analysis.to_dict())

            # Explicit SKIP_UNSAFE action
            if analysis.action == ConversionAction.SKIP_UNSAFE:
                result.placeholders_preserved += 1
                result.skipped_shape_ids.append(shape_id)
                result.preserved_shape_ids.append(shape_id)
                continue

            # Approval check
            if approved_set is not None and shape_id not in approved_set:
                result.placeholders_preserved += 1
                result.preserved_shape_ids.append(shape_id)
                continue

            # Preserved categories
            if analysis.action == ConversionAction.PRESERVE_EXISTING or analysis.status == PlaceholderStatus.PRESERVED:
                result.placeholders_preserved += 1
                result.preserved_shape_ids.append(shape_id)
                continue

            # Manual Review categories
            if analysis.action == ConversionAction.MANUAL_REVIEW or analysis.status == PlaceholderStatus.MANUAL_REVIEW:
                result.placeholders_manual_review += 1
                result.manual_review_shape_ids.append(shape_id)
                result.warnings.append(
                    f"Slide {slide_number} Shape {shape_id} ('{analysis.shape_name}'): {analysis.reason}"
                )
                continue

            # Safe to convert: execute deterministic transformation
            if analysis.action in (ConversionAction.CONVERT_EXISTING, ConversionAction.CREATE_NEW) or analysis.status == PlaceholderStatus.SAFE_TO_CONVERT:
                try:
                    created_id = self._execute_conversion(analysis, sp_tree, raw_slide=raw_slide)
                    if analysis.action == ConversionAction.CREATE_NEW and created_id:
                        result.created_shape_ids.append(created_id)
                        result.placeholders_converted += 1
                    else:
                        result.placeholders_converted += 1
                        result.converted_shape_ids.append(shape_id)
                    logger.info("Successfully converted Shape %d on Slide %d", shape_id, slide_number)
                except Exception as e:
                    logger.error("Failed to convert Shape %d on Slide %d: %s", shape_id, slide_number, e, exc_info=True)
                    result.warnings.append(f"Conversion failed for Shape {shape_id}: {str(e)}")
                    result.placeholders_preserved += 1
                    result.preserved_shape_ids.append(shape_id)

        safe, verify_errors = self.verify_integrity(
            sp_tree=sp_tree,
            original_shape_ids=original_shape_ids,
            original_shape_ids_list=original_shape_ids_list,
            animation_target_ids=anim_targets,
            original_text_map=original_text_map,
            original_xfrm_map=original_xfrm_map,
            original_group_count=original_group_count,
            original_rel_ids=original_rel_ids,
            allow_created_ids=set(result.created_shape_ids)
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
            result.created_shape_ids = []
            result.verification_category = DifferenceCategory.VERIFICATION_FAILURE.value
            return result

        result.verified = True
        result.verification_category = (
            DifferenceCategory.EXPECTED_ACCESSIBILITY_CHANGE.value
            if result.placeholders_converted > 0
            else "PRESERVED"
        )
        return result

    def _execute_conversion(
        self,
        analysis: PlaceholderAnalysis,
        sp_tree,
        raw_slide=None
    ) -> Optional[int]:
        """
        Executes physical XML mutation for an individual approved shape using ConvertPlaceholder_v2 logic.
        """
        shape_elem = analysis.raw_element
        if shape_elem is None:
            raise ValueError(f"Shape {analysis.shape_id} raw_element is None.")

        # Check if already a placeholder or marked decorative
        if cp_v2.has_ph(shape_elem):
            logger.info("Preserving existing native placeholder for Shape %d", analysis.shape_id)
            return None

        if cp_v2.has_existing_decorative_mark(shape_elem):
            logger.info("Preserving decorative shape %d", analysis.shape_id)
            return None

        # Determine target placeholder type from analysis
        target_type = (analysis.proposed_placeholder_type or "").lower()
        if not target_type or target_type == "none":
            if analysis.is_picture:
                target_type = "pic"
            elif analysis.is_table:
                target_type = "tbl"
            elif analysis.category == PlaceholderCategory.STANDALONE_TITLE_BOX:
                target_type = "title"
            elif analysis.category == PlaceholderCategory.STANDALONE_SUBTITLE_BOX:
                target_type = "subTitle"
            else:
                target_type = "body"

        # Determine non-colliding index
        idx_val = None
        if target_type not in ("title", "ctrtitle"):
            used = set()
            if sp_tree is not None:
                for ph_node in _safe_xpath(sp_tree, ".//p:ph[@idx]"):
                    v = ph_node.get("idx")
                    if v and v.isdigit():
                        used.add(int(v))
            if analysis.placeholder_idx and str(analysis.placeholder_idx).isdigit() and int(analysis.placeholder_idx) not in used:
                idx_val = int(analysis.placeholder_idx)
            else:
                cand = 1
                while cand in used:
                    cand += 1
                idx_val = cand

        tag = etree.QName(shape_elem).localname

        # Conversion for body / picture / table shapes using ConvertPlaceholder_v2 functions
        if target_type in ("body", "pic", "tbl"):
            if target_type == "body" and tag == "sp":
                cp_v2.freeze_paragraph_properties(shape_elem)
                cp_v2.backfill_run_sizes(shape_elem)
            cp_v2.freeze_autofit(shape_elem)
            cp_v2.add_placeholder(shape_elem, target_type, idx_val)
            return None

        elif target_type in ("title", "ctrtitle"):
            if tag == "sp":
                cp_v2.freeze_paragraph_properties(shape_elem)
                cp_v2.freeze_autofit(shape_elem)
                cp_v2.backfill_run_sizes(shape_elem)
            nv = cp_v2.get_nv_container(shape_elem)
            if nv is not None:
                nvpr = nv.find(cp_v2.qn("p", "nvPr"))
                if nvpr is None:
                    nvpr = etree.SubElement(nv, cp_v2.qn("p", "nvPr"))
                if nvpr.find(cp_v2.qn("p", "ph")) is None:
                    ph = etree.Element(cp_v2.qn("p", "ph"), attrib={"type": target_type})
                    nvpr.insert(0, ph)
            return None

        elif target_type == "subtitle":
            if tag == "sp":
                cp_v2.freeze_paragraph_properties(shape_elem)
                cp_v2.freeze_autofit(shape_elem)
                cp_v2.backfill_run_sizes(shape_elem)
            nv = cp_v2.get_nv_container(shape_elem)
            if nv is not None:
                nvpr = nv.find(cp_v2.qn("p", "nvPr"))
                if nvpr is None:
                    nvpr = etree.SubElement(nv, cp_v2.qn("p", "nvPr"))
                if nvpr.find(cp_v2.qn("p", "ph")) is None:
                    ph = etree.Element(cp_v2.qn("p", "ph"), attrib={"type": "subTitle", "idx": str(idx_val)})
                    nvpr.insert(0, ph)
            return None

        elif target_type == "decorative":
            cp_v2.mark_decorative(shape_elem)
            return None

        return None

    def verify_integrity(
        self,
        sp_tree,
        original_shape_ids: Set[int],
        animation_target_ids: Set[int],
        original_shape_ids_list: Optional[List[int]] = None,
        original_text_map: Optional[Dict[int, str]] = None,
        original_xfrm_map: Optional[Dict[int, Tuple[float, float, float, float]]] = None,
        original_group_count: Optional[int] = None,
        original_rel_ids: Optional[Set[str]] = None,
        allow_created_ids: Optional[Set[int]] = None
    ) -> Tuple[bool, List[str]]:
        """
        Validates post-conversion integrity.
        """
        safe, errors, _ = self.verify_integrity_comprehensive(
            sp_tree=sp_tree,
            original_shape_ids=original_shape_ids,
            original_shape_ids_list=original_shape_ids_list,
            original_text_map=original_text_map,
            original_xfrm_map=original_xfrm_map,
            original_group_count=original_group_count,
            original_rel_ids=original_rel_ids,
            animation_target_ids=animation_target_ids,
            allow_created_ids=allow_created_ids
        )
        return safe, errors

    def verify_integrity_comprehensive(
        self,
        sp_tree,
        original_shape_ids: Set[int],
        original_shape_ids_list: Optional[List[int]] = None,
        original_text_map: Optional[Dict[int, str]] = None,
        original_xfrm_map: Optional[Dict[int, Tuple[float, float, float, float]]] = None,
        original_group_count: Optional[int] = None,
        original_rel_ids: Optional[Set[str]] = None,
        animation_target_ids: Optional[Set[int]] = None,
        allow_created_ids: Optional[Set[int]] = None
    ) -> Tuple[bool, List[str], List[Dict[str, Any]]]:
        """
        Deep before/after verification classifying all differences into DifferenceCategory.
        Rejects any UNINTENDED_STRUCTURAL_CHANGE, UNINTENDED_VISUAL_CHANGE, or VERIFICATION_FAILURE.
        """
        errors = []
        differences = []
        anim_targets = animation_target_ids or set()
        created_ids = allow_created_ids or set()
        current_ids = self._collect_shape_ids(sp_tree)

        # 1. Shape ID preservation
        missing = original_shape_ids - current_ids
        if missing:
            err = f"Original shape IDs missing after conversion: {sorted(list(missing))}"
            errors.append(err)
            differences.append({
                "category": DifferenceCategory.UNINTENDED_STRUCTURAL_CHANGE.value,
                "description": err,
                "shape_ids": list(missing)
            })

        # 2. Animation target ID preservation
        for a_id in anim_targets:
            if a_id in original_shape_ids and a_id not in current_ids:
                err = f"Animation target shape ID {a_id} was lost during placeholder conversion!"
                errors.append(err)
                differences.append({
                    "category": DifferenceCategory.UNINTENDED_STRUCTURAL_CHANGE.value,
                    "description": err,
                    "shape_id": a_id
                })

        # 3. Duplicate ID prevention (only newly introduced duplicates)
        orig_counts: Dict[int, int] = {}
        if original_shape_ids_list is not None:
            for s_id in original_shape_ids_list:
                orig_counts[s_id] = orig_counts.get(s_id, 0) + 1

        all_ids_list = self._collect_shape_ids_list(sp_tree)
        counts: Dict[int, int] = {}
        for s_id in all_ids_list:
            counts[s_id] = counts.get(s_id, 0) + 1

        dupes = [s_id for s_id, count in counts.items() if count > 1 and count > orig_counts.get(s_id, 0)]
        if dupes:
            err = f"Duplicate shape IDs detected after conversion: {dupes}"
            errors.append(err)
            differences.append({
                "category": DifferenceCategory.VERIFICATION_FAILURE.value,
                "description": err,
                "shape_ids": dupes
            })

        # 4. Text content equality
        if original_text_map is not None:
            current_text_map = self._collect_text_map(sp_tree)
            for s_id, orig_text in original_text_map.items():
                if s_id in current_ids and s_id not in created_ids:
                    curr_text = current_text_map.get(s_id, "")
                    if orig_text != curr_text:
                        err = f"Shape ID {s_id} text altered: expected '{orig_text[:30]}', got '{curr_text[:30]}'"
                        errors.append(err)
                        differences.append({
                            "category": DifferenceCategory.UNINTENDED_STRUCTURAL_CHANGE.value,
                            "description": err,
                            "shape_id": s_id
                        })

        # 5. Geometry bounding box preservation
        if original_xfrm_map is not None:
            current_xfrm_map = self._collect_xfrm_map(sp_tree)
            for s_id, orig_box in original_xfrm_map.items():
                if s_id in current_xfrm_map and s_id not in created_ids:
                    curr_box = current_xfrm_map[s_id]
                    delta = max(abs(orig_box[i] - curr_box[i]) for i in range(4))
                    if delta > 0.5:
                        err = f"Shape ID {s_id} geometry drifted by {delta:.2f}pt (limit: 0.5pt)"
                        errors.append(err)
                        differences.append({
                            "category": DifferenceCategory.UNINTENDED_VISUAL_CHANGE.value,
                            "description": err,
                            "shape_id": s_id
                        })

        # 6. Group hierarchy preservation
        if original_group_count is not None:
            current_group_count = len(_safe_xpath(sp_tree, ".//p:grpSp"))
            if current_group_count != original_group_count:
                err = f"Group count changed from {original_group_count} to {current_group_count}"
                errors.append(err)
                differences.append({
                    "category": DifferenceCategory.UNINTENDED_STRUCTURAL_CHANGE.value,
                    "description": err
                })

        # 7. Relationship preservation
        if original_rel_ids is not None:
            current_rel_ids = self._collect_relationship_ids(sp_tree)
            missing_rels = original_rel_ids - current_rel_ids
            if missing_rels:
                err = f"Relationship IDs lost after conversion: {missing_rels}"
                errors.append(err)
                differences.append({
                    "category": DifferenceCategory.UNINTENDED_STRUCTURAL_CHANGE.value,
                    "description": err
                })

        # 8. Placeholder index uniqueness check (no duplicate idx on slide)
        used_indices: Dict[str, List[int]] = {}
        for sp in _safe_xpath(sp_tree, ".//p:sp|.//p:pic|.//p:graphicFrame|.//p:cxnSp"):
            ph_nodes = _safe_xpath(sp, ".//p:ph")
            if ph_nodes:
                idx_val = ph_nodes[0].get("idx")
                if idx_val is not None:
                    nv_id = _safe_xpath(sp, ".//p:cNvPr/@id")
                    s_id = int(nv_id[0]) if nv_id and nv_id[0].isdigit() else -1
                    used_indices.setdefault(idx_val, []).append(s_id)
        for idx_val, s_ids in used_indices.items():
            if len(s_ids) > 1:
                err = f"Duplicate placeholder index idx={idx_val} detected on shapes {s_ids}"
                errors.append(err)
                differences.append({
                    "category": DifferenceCategory.VERIFICATION_FAILURE.value,
                    "description": err,
                    "shape_ids": s_ids
                })

        # Record expected accessibility changes for newly converted or created shapes
        for c_id in (current_ids - original_shape_ids):
            if c_id in created_ids:
                differences.append({
                    "category": DifferenceCategory.EXPECTED_ACCESSIBILITY_CHANGE.value,
                    "description": f"New native placeholder created with ID {c_id}",
                    "shape_id": c_id
                })

        return len(errors) == 0, errors, differences

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

    def _collect_text_map(self, sp_tree) -> Dict[int, str]:
        """Maps shape IDs to concatenated plain text."""
        text_map: Dict[int, str] = {}
        for child in sp_tree:
            cnv_prs = _safe_xpath(child, ".//p:cNvPr")
            if not cnv_prs:
                continue
            s_id = cnv_prs[0].get("id")
            if not s_id or not s_id.isdigit():
                continue
            texts = [t.text for t in _safe_xpath(child, ".//a:t") if t.text]
            text_map[int(s_id)] = "".join(texts)
        return text_map

    def _collect_xfrm_map(self, sp_tree) -> Dict[int, Tuple[float, float, float, float]]:
        """Maps shape IDs to bounding boxes (left, top, width, height) in points."""
        xfrm_map: Dict[int, Tuple[float, float, float, float]] = {}
        for child in sp_tree:
            cnv_prs = _safe_xpath(child, ".//p:cNvPr")
            if not cnv_prs:
                continue
            s_id = cnv_prs[0].get("id")
            if not s_id or not s_id.isdigit():
                continue
            off = _safe_xpath(child, ".//p:spPr/a:xfrm/a:off") or _safe_xpath(child, ".//p:picPr/a:xfrm/a:off") or _safe_xpath(child, ".//a:xfrm/a:off")
            ext = _safe_xpath(child, ".//p:spPr/a:xfrm/a:ext") or _safe_xpath(child, ".//p:picPr/a:xfrm/a:ext") or _safe_xpath(child, ".//a:xfrm/a:ext")
            if off and ext:
                try:
                    xfrm_map[int(s_id)] = (
                        emu_to_points(int(off[0].get("x", "0"))),
                        emu_to_points(int(off[0].get("y", "0"))),
                        emu_to_points(int(ext[0].get("cx", "0"))),
                        emu_to_points(int(ext[0].get("cy", "0"))),
                    )
                except Exception:
                    pass
        return xfrm_map

    def _collect_relationship_ids(self, sp_tree) -> Set[str]:
        """Collects all relationship IDs within sp_tree."""
        rel_ids: Set[str] = set()
        try:
            for attr in sp_tree.xpath(".//@r:id", namespaces=NSMAP):
                rel_ids.add(str(attr))
        except Exception:
            pass
        return rel_ids
