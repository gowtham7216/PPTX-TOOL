"""
SlideGuard Reading Order Resolver.
Determines semantic reading order following the strict priority:
  1. Existing meaningful order (preserve if already valid)
  2. Semantic hierarchy (Title -> Subtitle -> Content)
  3. Spatial & column regions (Left column -> Right column)
  4. Footer exclusion (Placeholders / Wiley footers placed at the very end of sequence)
"""

from typing import List, Dict, Any, Optional, Tuple
from dataclasses import dataclass, field
from rules.reading_order.semantic_classifier import SemanticItem, SemanticRole
from rules.reading_order.layout_analyzer import LayoutRegionAnalyzer, SlideLayoutStructure


@dataclass
class ReadingOrderResult:
    """Detailed diagnostic outcome of reading order resolution for a slide."""
    slide_number: int
    original_order: List[int]  # List of shape IDs in current p:spTree order
    proposed_order: List[int]  # List of shape IDs in proposed reading sequence
    changes_required: bool
    confidence: str            # "HIGH", "MEDIUM", "LOW"
    reasoning: List[str] = field(default_factory=list)
    excluded_footer_objects: List[Dict[str, Any]] = field(default_factory=list)
    gemini_decisions: List[Dict[str, Any]] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "original_order": self.original_order,
            "proposed_order": self.proposed_order,
            "changes_required": self.changes_required,
            "confidence": self.confidence,
            "reasoning": self.reasoning,
            "excluded_footer_objects": self.excluded_footer_objects,
            "gemini_decisions": self.gemini_decisions,
            "warnings": self.warnings,
        }


class ReadingOrderResolver:
    """
    Synthesizes semantic classification, Gemini 3.8 Flash High analysis,
    and spatial layout analysis to produce a verifiable, safe reading sequence.
    """

    def __init__(
        self,
        layout_analyzer: Optional[LayoutRegionAnalyzer] = None,
        assistant: Optional[Any] = None
    ):
        self.layout_analyzer = layout_analyzer or LayoutRegionAnalyzer()
        from engine.intelligence.gemini_assistant import GeminiAssistant
        self.assistant = assistant or GeminiAssistant(model_name="gemini-3.8-flash-high")

    def resolve_reading_order(
        self,
        slide_number: int,
        items: List[SemanticItem],
        slide_w_pt: float,
        slide_h_pt: float
    ) -> ReadingOrderResult:
        original_order = [item.shape_id for item in items]
        reasoning = []
        warnings = []
        excluded_footers = []

        if len(items) <= 1:
            return ReadingOrderResult(
                slide_number=slide_number,
                original_order=original_order,
                proposed_order=original_order,
                changes_required=False,
                confidence="HIGH",
                reasoning=["Slide contains 0 or 1 shapes; reading order is already trivial."],
                excluded_footer_objects=[],
                gemini_decisions=[],
                warnings=[]
            )

        # 1. Deterministic Structural Extraction
        slide_context = {
            "slide_number": slide_number,
            "slide_dimensions": {"width_pt": slide_w_pt, "height_pt": slide_h_pt},
            "shapes": [
                {
                    "shape_id": it.shape_id,
                    "name": it.name,
                    "shape_type": it.role.value,
                    "text_content": it.text_preview,
                    "geometry": {
                        "left_pt": it.left,
                        "top_pt": it.top,
                        "width_pt": it.width,
                        "height_pt": it.height
                    },
                    "z_order": it.z_order,
                    "is_placeholder": it.is_placeholder,
                    "placeholder_type": it.placeholder_type,
                    "parent_group_id": it.parent_group_id,
                    "media_type": (
                        "IMAGE" if it.role == SemanticRole.IMAGE
                        else ("TABLE" if it.role == SemanticRole.TABLE
                              else ("CHART" if it.role == SemanticRole.CHART else "TEXT"))
                    ),
                    "surrounding_context": f"top={it.top:.1f}pt, left={it.left:.1f}pt, width={it.width:.1f}pt",
                    "repeated_structure_hint": "CARD_BADGE" if it.role == SemanticRole.CARD_BADGE else None
                }
                for it in items
            ]
        }

        # 2. Gemini 3.8 Flash High Semantic Analysis
        gemini_result = self.assistant.analyze_reading_order(slide_context)
        decisions = gemini_result.get("decisions", [])
        gemini_conf = float(gemini_result.get("confidence", 0.5))
        gemini_rationale = gemini_result.get("rationale", "")
        warnings.extend(gemini_result.get("warnings", []))

        if gemini_rationale:
            reasoning.append(f"Gemini Semantic Analysis: {gemini_rationale}")

        # 3. Validation & Safety Gate
        # Hard Rule: NEVER invent or drop shapes; shape IDs must exactly match
        known_ids = set(original_order)
        dec_map = {d["shape_id"]: d for d in decisions if d["shape_id"] in known_ids}

        # Ensure all existing items have a decision
        for it in items:
            if it.shape_id not in dec_map:
                dec_map[it.shape_id] = {
                    "shape_id": it.shape_id,
                    "semantic_role": it.role.value,
                    "reading_order": 0 if it.role == SemanticRole.DECORATIVE else 999,
                    "confidence": 0.5,
                    "action": "PRESERVE",
                    "reason": "Fallback decision for shape omitted by AI"
                }

        # Separate decorative backgrounds from content sequence
        decorative_ids = [
            it.shape_id for it in items
            if dec_map[it.shape_id]["action"] == "EXCLUDE_DECORATIVE"
            or dec_map[it.shape_id]["semantic_role"] == "DECORATIVE"
            or dec_map[it.shape_id]["reading_order"] == 0
        ]

        # Content elements ordered by reading_order
        content_decisions = [
            dec_map[it.shape_id] for it in items
            if it.shape_id not in decorative_ids
        ]
        content_decisions.sort(key=lambda d: d.get("reading_order", 999))
        content_ids = [d["shape_id"] for d in content_decisions]

        # Analyze visual layout structure
        layout = self.layout_analyzer.analyze_layout(items, slide_w_pt, slide_h_pt)

        # 1. Enforce Timeline / Process Flow horizontal progression (Task 1)
        if layout.detected_layout_type == "TIMELINE_HORIZONTAL":
            # In horizontal timeline/process flows, sort body content strictly by left coordinate
            timeline_items = [
                it for it in items
                if it.shape_id in content_ids
                and not it.is_title
                and it.role not in (SemanticRole.TITLE, SemanticRole.FOOTER, SemanticRole.SLIDE_NUMBER)
                and (it.top < slide_h_pt * 0.88)
            ]
            timeline_sorted_ids = [it.shape_id for it in sorted(timeline_items, key=lambda x: x.left)]
            other_content_ids = [sid for sid in content_ids if sid not in timeline_sorted_ids]
            title_ids = [it.shape_id for it in items if (it.is_title or it.role == SemanticRole.TITLE) and it.shape_id in content_ids]
            footer_ids = [sid for sid in other_content_ids if sid not in title_ids]
            content_ids = title_ids + timeline_sorted_ids + footer_ids
            reasoning.append("Enforced chronological left-to-right horizontal progression for alternating timeline flow.")

        # 2. Enforce Title 1 at index 0 of the content reading order sequence (Task 3)
        title_items = [it for it in items if (it.is_title or it.role == SemanticRole.TITLE) and it.shape_id in content_ids]
        if title_items:
            sorted_titles = sorted(title_items, key=lambda x: (x.top, x.left))
            primary_title_id = sorted_titles[0].shape_id
            content_ids = [primary_title_id] + [sid for sid in content_ids if sid != primary_title_id]

        # 3. Empty Shape Deprioritization & Background Handling with Visual Occlusion Guard
        # Background canvas shapes, banners, or shapes that originally sat behind content they overlap with
        # MUST remain in the background (bg_ids) to prevent visual occlusion (obscuring text).
        orig_idx_map = {sid: idx for idx, sid in enumerate(original_order)}
        bg_ids = []
        other_decorative_ids = []
        content_item_objs = [it for it in items if it.shape_id in content_ids]

        for s_id in decorative_ids:
            it = next((x for x in items if x.shape_id == s_id), None)
            if not it:
                continue

            # a) Full canvas or wide banner background
            is_wide_bg = (it.width >= slide_w_pt * 0.80) or (it.height >= slide_h_pt * 0.80)

            # b) Visual Occlusion Guard:
            # If this decorative/empty shape was originally BEHIND any content shape that it physically overlaps with,
            # moving it after content would "Bring to Front" and cover the content text.
            is_underlay_visual = False
            for c_it in content_item_objs:
                if orig_idx_map.get(s_id, 0) < orig_idx_map.get(c_it.shape_id, 0):
                    if self._shapes_overlap(it, c_it):
                        is_underlay_visual = True
                        break

            if is_wide_bg or is_underlay_visual:
                bg_ids.append(s_id)
            else:
                other_decorative_ids.append(s_id)

        # Preserve original relative Z-order among background shapes
        bg_ids.sort(key=lambda sid: orig_idx_map.get(sid, 0))

        # Place content immediately after background, followed by all other empty decorative shapes
        proposed_order = bg_ids + content_ids + other_decorative_ids

        # 4. Universal Visual Occlusion Guard
        # Guarantee that no shape A that was originally behind text/content shape B and overlaps with B
        # is placed after B in proposed_order (preventing accidental 'Bring to Front' on background visuals).
        proposed_order = self._enforce_visual_occlusion_guard(proposed_order, original_order, items)

        # Track excluded footers
        for it in items:
            role = dec_map[it.shape_id].get("semantic_role", "")
            if role in ("FOOTER", "SLIDE_NUMBER"):
                excluded_footers.append({
                    "shape_id": it.shape_id,
                    "name": it.name,
                    "role": role,
                    "text": it.text_preview,
                    "rationale": dec_map[it.shape_id].get("reason", "")
                })

        # 4. Confidence Gate & Threshold Rules
        # - HIGH >= 0.85 -> safe auto-reorder eligible
        # - MEDIUM 0.70-0.849 -> manual review unless deterministic rules strongly confirm
        # - LOW < 0.70 -> preserve existing order and flag for review

        if gemini_conf >= 0.85:
            confidence_tier = "HIGH"
        elif gemini_conf >= 0.70 or layout.detected_layout_type == "TIMELINE_HORIZONTAL":
            if self._is_layout_strongly_confirmed(items, layout):
                confidence_tier = "HIGH"
                reasoning.append("Deterministic layout analysis strongly confirmed semantic hierarchy.")
            else:
                confidence_tier = "MEDIUM"
                warnings.append("Medium confidence layout; manual review recommended.")
        else:
            confidence_tier = "LOW"
            warnings.append("Low confidence layout (< 0.70); preserving original sequence for review.")

        # 5. Check Priority 1: Is existing author order already valid?
        is_already_acceptable, preservation_reason = self._is_existing_order_acceptable(
            items, original_order, proposed_order, layout
        )
        if is_already_acceptable:
            changes_required = False
            confidence_tier = "HIGH"
            reasoning.insert(0, f"Preserving existing author order: {preservation_reason}")
            final_proposed_order = original_order
        elif confidence_tier == "LOW":
            # Safety Gate: LOW confidence preserves existing order and flags for manual review
            changes_required = False
            reasoning.append("Preserved existing reading order due to LOW confidence layout.")
            final_proposed_order = original_order
        elif confidence_tier == "MEDIUM":
            # Medium confidence without strong confirmation requires manual review
            changes_required = False
            reasoning.append("Preserved existing reading order pending manual review (MEDIUM confidence).")
            final_proposed_order = original_order
        else:
            # Safe reorder
            changes_required = True
            reasoning.append(f"Reordering proposed: {original_order} -> {proposed_order}")
            final_proposed_order = proposed_order

        return ReadingOrderResult(
            slide_number=slide_number,
            original_order=original_order,
            proposed_order=final_proposed_order,
            changes_required=changes_required,
            confidence=confidence_tier,
            reasoning=reasoning,
            excluded_footer_objects=excluded_footers,
            gemini_decisions=list(dec_map.values()),
            warnings=warnings
        )

    def _order_column_items(
        self,
        column_items: List[SemanticItem],
        ordered_output: List[SemanticItem],
        reasoning: List[str]
    ) -> None:
        """Orders items within a column, ensuring paired elements (image + caption) stay together."""
        # Check for image + caption pairs
        used_ids = set()
        for i, item in enumerate(column_items):
            if item.shape_id in used_ids:
                continue

            ordered_output.append(item)
            used_ids.add(item.shape_id)

            # If this is an image, check if a caption immediately follows
            if item.role == SemanticRole.IMAGE:
                for other in column_items:
                    if other.shape_id not in used_ids and other.role == SemanticRole.CAPTION:
                        if 0 <= (other.top - item.bottom) <= 24.0:
                            ordered_output.append(other)
                            used_ids.add(other.shape_id)
                            reasoning.append(f"Paired caption '{other.name}' immediately following image '{item.name}'.")
                            break

    def _is_existing_order_acceptable(
        self,
        items: List[SemanticItem],
        original_order: List[int],
        proposed_order: List[int],
        layout: SlideLayoutStructure
    ) -> Tuple[bool, str]:
        """
        Determines whether the existing shape-tree sequence already provides a logical reading flow.
        """
        if original_order == proposed_order:
            return True, "Original order perfectly matches proposed accessibility sequence."

        # Check: Is title at or near the beginning (first 1-2 shapes)?
        title_ids = {h.shape_id for h in layout.header_items}
        if title_ids:
            first_non_decorative = [i for i in original_order if not self._is_pure_decorative(i, items)]
            if first_non_decorative and first_non_decorative[0] in title_ids:
                # Title is first!
                footer_ids = {f.shape_id for f in layout.footer_items}
                if footer_ids:
                    # Are all footers at the end of original_order?
                    last_n = len(footer_ids)
                    orig_end_ids = set(original_order[-last_n:])
                    if not footer_ids.issubset(orig_end_ids):
                        return False, ""

                # Check if non-decorative body items are in monotonic top-to-bottom sequence
                content_items = [it for it in items if not self._is_pure_decorative(it.shape_id, items) and it.shape_id not in footer_ids]
                item_map = {it.shape_id: it for it in content_items}
                orig_content = [item_map[sid] for sid in original_order if sid in item_map]
                tops = [it.top for it in orig_content]
                if tops == sorted(tops) and not layout.is_multi_column:
                    return True, "Title is early, footers are separated, and body sequence is strictly top-to-bottom."

        return False, ""


    def _is_pure_decorative(self, shape_id: int, items: List[SemanticItem]) -> bool:
        for it in items:
            if it.shape_id == shape_id and it.role == SemanticRole.DECORATIVE:
                return True
        return False

    def _is_layout_strongly_confirmed(
        self,
        items: List[SemanticItem],
        layout: SlideLayoutStructure
    ) -> bool:
        """
        Determines whether deterministic structural rules strongly confirm the reading hierarchy.
        Used to safely promote MEDIUM confidence (0.70 - 0.849) layouts to auto-reorder eligible.
        """
        has_header = bool(layout.header_items)
        content_items = [
            it for it in items
            if it.role not in (SemanticRole.DECORATIVE, SemanticRole.FOOTER, SemanticRole.SLIDE_NUMBER)
        ]
        if not content_items:
            return True
        if has_header and (layout.detected_layout_type == "TIMELINE_HORIZONTAL" or len(layout.body_columns) <= 3):
            return True
        return False

    def _determine_confidence(
        self,
        items: List[SemanticItem],
        layout: SlideLayoutStructure,
        warnings: List[str]
    ) -> str:
        """Calculates confidence score (HIGH, MEDIUM, LOW) based on structural clarity."""
        has_clear_title = bool(layout.header_items)
        has_clear_structure = len(layout.body_columns) <= 2

        # Check for heavy overlap among content items (excluding decorative shapes)
        content_items = [
            it for it in items
            if it.role not in (SemanticRole.DECORATIVE, SemanticRole.FOOTER, SemanticRole.SLIDE_NUMBER)
        ]
        overlaps = 0
        for i in range(len(content_items)):
            for j in range(i + 1, len(content_items)):
                b1 = content_items[i].bbox_pt
                b2 = content_items[j].bbox_pt
                ix = max(0.0, min(b1[0] + b1[2], b2[0] + b2[2]) - max(b1[0], b2[0]))
                iy = max(0.0, min(b1[1] + b1[3], b2[1] + b2[3]) - max(b1[1], b2[1]))
                intersection = ix * iy
                min_area = min(b1[2] * b1[3], b2[2] * b2[3])
                if min_area > 0 and (intersection / min_area) > 0.40:
                    overlaps += 1

        if overlaps > 3:
            warnings.append(f"Detected {overlaps} overlapping shape bounding boxes; spatial reading flow is ambiguous.")
            return "LOW"

        if has_clear_title and has_clear_structure and overlaps == 0:
            return "HIGH"

        return "MEDIUM"

    @staticmethod
    def _shapes_overlap(item1: SemanticItem, item2: SemanticItem, tolerance_pt: float = 2.0) -> bool:
        """Determines if two shapes have overlapping 2D bounding boxes."""
        overlap_w = min(item1.right, item2.right) - max(item1.left, item2.left)
        overlap_h = min(item1.bottom, item2.bottom) - max(item1.top, item2.top)
        return (overlap_w > tolerance_pt) and (overlap_h > tolerance_pt)

    def _enforce_visual_occlusion_guard(
        self,
        proposed_order: List[int],
        original_order: List[int],
        items: List[SemanticItem]
    ) -> List[int]:
        """
        Enforces visual occlusion safety:
        If shape A was originally BEHIND shape B, and shape A overlaps with shape B,
        and shape B contains text or is a content placeholder, shape A must NEVER
        be placed after shape B in the shape tree (which would render shape A on top of B).
        """
        item_map = {it.shape_id: it for it in items}
        orig_idx_map = {sid: idx for idx, sid in enumerate(original_order)}
        adjusted = list(proposed_order)

        changed = True
        iterations = 0
        while changed and iterations < 20:
            changed = False
            iterations += 1
            for i in range(len(adjusted)):
                for j in range(i + 1, len(adjusted)):
                    s_first = adjusted[i]
                    s_second = adjusted[j]

                    it_first = item_map.get(s_first)
                    it_second = item_map.get(s_second)

                    if not it_first or not it_second:
                        continue

                    # If s_second was originally BEHIND s_first, but is now AFTER s_first in adjusted order:
                    if orig_idx_map.get(s_second, 0) < orig_idx_map.get(s_first, 0):
                        has_text = bool((it_first.text_preview or "").strip()) or it_first.is_title or it_first.role in (
                            SemanticRole.TITLE, SemanticRole.SUBTITLE, SemanticRole.BODY_TEXT,
                            SemanticRole.CARD_HEADER, SemanticRole.CARD_BODY, SemanticRole.TABLE
                        )
                        if has_text and self._shapes_overlap(it_second, it_first):
                            adjusted.remove(s_second)
                            adjusted.insert(i, s_second)
                            changed = True
                            break
                if changed:
                    break

        return adjusted

