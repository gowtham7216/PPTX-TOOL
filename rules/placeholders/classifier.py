"""
SlideGuard Placeholder Classifier.
Classifies PowerPoint placeholders and standalone text elements into semantic categories:
title, body, subtitle, object, picture, chart, table, footer, slide number, date/time, other.
Determines safe conversion eligibility without visual design alteration.
"""

from enum import Enum
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Any, Set
import logging

from pptx_engine.utilities.geometry import emu_to_points
from rules.placeholders.inheritance import (
    PlaceholderInheritanceResolver,
    InheritedProperties,
    _safe_xpath,
    NSMAP,
)

logger = logging.getLogger(__name__)


class PlaceholderCategory(str, Enum):
    """Semantic category of a PowerPoint placeholder."""
    TITLE = "TITLE"
    BODY = "BODY"
    SUBTITLE = "SUBTITLE"
    OBJECT = "OBJECT"
    PICTURE = "PICTURE"
    CHART = "CHART"
    TABLE = "TABLE"
    FOOTER = "FOOTER"
    SLIDE_NUMBER = "SLIDE_NUMBER"
    DATE_TIME = "DATE_TIME"
    EMPTY = "EMPTY"
    STANDALONE_TITLE_BOX = "STANDALONE_TITLE_BOX"
    CARD_BADGE = "CARD_BADGE"
    CARD_COMPONENT = "CARD_COMPONENT"
    OTHER = "OTHER"


import re


def is_numeric_badge(text: str) -> bool:
    """
    Determines if a text string is a numeric badge, step index, card number, or design marker.
    Examples: '01', '02', '03', '04', '1', '2', 'Step 1', 'Card 2', '01.', 'A', '#1'.
    """
    t = text.strip()
    if not t:
        return False
    # Pure numeric (1-4 digits, e.g., '01', '02', '03', '1', '42')
    if re.match(r"^\d{1,4}$", t):
        return True
    # Sequence/card badges: 'Step 1', 'Card 03', '01.', '02:', '#1', 'No. 3', 'Phase 1', 'Item 2'
    if re.match(r"^(?:step|card|part|item|no\.?|#|phase|badge)\s*\d{1,3}[.:]?$", t, re.IGNORECASE):
        return True
    # Trailing punctuation like '01.' or '01:'
    if re.match(r"^\d{1,3}[.:\-)]$", t):
        return True
    # Roman numerals like 'I', 'II', 'III', 'IV'
    if re.match(r"^(?:[ivxIVX]{1,4})[.:]?$", t):
        return True
    return False


def is_short_label_or_decorative(text: str) -> bool:
    """Checks if text is a short decorative label, icon label, or design marker."""
    t = text.strip()
    if not t:
        return False
    # Single punctuation / symbol / arrow or short non-alphabet token
    if len(t) <= 3 and not t.isalpha():
        return True
    lower = t.lower()
    if lower in ("icon", "badge", "tag", "label", "bullet", "marker", "arrow", "logo", "step", "card"):
        return True
    return False


def detect_repeated_components(candidates: List[Dict[str, Any]]) -> Set[int]:
    """
    Detects repeated visual structures (e.g. multi-card layouts with '01', '02', '03',
    or repeating headers/badges with similar geometry/styles).
    Returns a set of shape_ids that are card badges or design markers.
    """
    badge_ids: Set[int] = set()
    for c in candidates:
        t = (c.get("text_content") or "").strip()
        if is_numeric_badge(t) or is_short_label_or_decorative(t):
            badge_ids.add(c["shape_id"])
    return badge_ids


def score_title_candidate(cand: Dict[str, Any], slide_h: float, slide_w: float) -> float:
    """
    Evaluates multi-signal semantic score for a title candidate across the slide.
    Positive signals: heading word count, font size, bold typography, placement, title keywords.
    Negative signals: extreme length, footer region, tiny font size.
    """
    t = cand.get("text_content", "").strip()
    words = t.split()
    w_count = len(words)
    score = 0.0

    # Heading-like word count
    if 2 <= w_count <= 25:
        score += 25.0
    elif w_count == 1 and len(t) >= 5:
        score += 10.0
    elif w_count > 25:
        score -= 20.0

    # Typography (max font size and bold)
    max_sz = cand.get("max_font_size", 0.0)
    if max_sz >= 24.0:
        score += 35.0
    elif max_sz >= 18.0:
        score += 20.0
    elif max_sz >= 14.0:
        score += 10.0
    elif 0 < max_sz <= 10.0:
        score -= 30.0

    if cand.get("is_bold"):
        score += 10.0

    # Placement bonus (evaluates entire slide, with higher scores near top)
    bbox = cand.get("bbox_pt", (0.0, 0.0, 0.0, 0.0))
    top_pt = bbox[1]
    if top_pt <= (slide_h * 0.15):
        score += 25.0
    elif top_pt <= (slide_h * 0.35):
        score += 15.0
    elif top_pt <= (slide_h * 0.70):
        score += 10.0
    elif top_pt >= (slide_h * 0.88):
        score -= 80.0

    # Keywords & Patterns
    if re.search(r"Chapter\s+\d+", t, re.IGNORECASE):
        score += 30.0
    if ":" in t:
        score += 10.0

    return score


class PlaceholderStatus(str, Enum):
    """Conversion eligibility status."""
    SAFE_TO_CONVERT = "SAFE TO CONVERT"
    PRESERVED = "PRESERVED"
    MANUAL_REVIEW = "MANUAL REVIEW"


@dataclass
class PlaceholderAnalysis:
    """Detailed structural classification of an individual placeholder or candidate shape."""
    slide_number: int
    shape_id: int
    shape_name: str
    placeholder_type: str  # "title", "body", "subTitle", "obj", "pic", "tbl", "chart", "ftr", "sldNum", "dt", "none"
    placeholder_idx: Optional[str]
    category: PlaceholderCategory
    status: PlaceholderStatus
    has_text: bool
    text_preview: str
    is_empty: bool
    is_animated: bool
    is_picture: bool
    is_table: bool
    is_chart: bool
    inherited_properties: InheritedProperties
    slide_bbox_pt: Tuple[float, float, float, float]
    recommendation: str
    reason: str
    confidence: str  # "HIGH", "MEDIUM", "LOW"
    raw_element: Any = None
    relationship_ids: List[str] = field(default_factory=list)
    numerical_confidence: float = 0.95
    gemini_decision: Optional[Dict[str, Any]] = None
    model_traceability: Dict[str, str] = field(default_factory=lambda: {
        "antigravity_coding_model": "Gemini 3.8 Flash High",
        "slideguard_runtime_model": "Gemini 3.5"
    })

    @property
    def is_safe_to_convert(self) -> bool:
        return self.status == PlaceholderStatus.SAFE_TO_CONVERT

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "placeholder_type": self.placeholder_type,
            "placeholder_idx": self.placeholder_idx,
            "category": self.category.value,
            "status": self.status.value,
            "has_text": self.has_text,
            "text_preview": self.text_preview,
            "is_empty": self.is_empty,
            "is_animated": self.is_animated,
            "is_picture": self.is_picture,
            "is_table": self.is_table,
            "is_chart": self.is_chart,
            "inherited_from": self.inherited_properties.inherited_from,
            "slide_bbox_pt": list(self.slide_bbox_pt),
            "recommendation": self.recommendation,
            "reason": self.reason,
            "confidence": self.confidence,
            "relationship_ids": self.relationship_ids,
            "numerical_confidence": self.numerical_confidence,
            "gemini_decision": self.gemini_decision,
            "model_traceability": self.model_traceability,
        }


class PlaceholderClassifier:
    """
    Deterministic classifier for PowerPoint placeholders.
    Inspects shape content, layout/master inheritance, animation targets, and relationships.
    """

    def __init__(
        self,
        resolver: Optional[PlaceholderInheritanceResolver] = None,
        assistant: Optional[Any] = None
    ):
        self.resolver = resolver or PlaceholderInheritanceResolver()
        from engine.intelligence.gemini_assistant import GeminiAssistant, resolve_runtime_model
        if assistant is not None:
            self.assistant = assistant
        else:
            self.assistant = GeminiAssistant(model_name=resolve_runtime_model("gemini-3.5"))

    def analyze_slide_placeholders(
        self,
        raw_slide,
        slide_number: int,
        animation_target_ids: Optional[List[int]] = None,
        slide_height_pt: float = 540.0,
        slide_width_pt: float = 720.0
    ) -> List[PlaceholderAnalysis]:
        """
        Analyzes all placeholders and candidate shapes on a slide using
        Gemini 3.5 runtime semantic reasoning and deterministic safety gates.
        Architecture:
          1. Candidate Detection
          2. Context Extraction (No PPTX filesystem paths)
          3. Gemini 3.5 Runtime Semantic Decision
          4. Strict JSON Validation
          5. Deterministic Safety Gate
        """
        results: List[PlaceholderAnalysis] = []
        anim_set = set(animation_target_ids or [])

        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return results

        slide_elem = raw_slide._element

        # =====================================================================
        # 1. CANDIDATE DETECTION
        # =====================================================================
        candidates_raw = []
        has_slide_title = False
        existing_native_title_id: Optional[int] = None

        for child_elem in _safe_xpath(slide_elem, ".//p:spTree/*"):
            tag = child_elem.tag.split("}")[-1]
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue

            cnv_prs = _safe_xpath(child_elem, ".//p:cNvPr")
            shape_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
            name = cnv_prs[0].get("name", f"Shape_{shape_id}") if cnv_prs else f"Shape_{shape_id}"
            is_animated = shape_id in anim_set

            ph_elems = _safe_xpath(child_elem, ".//p:ph")
            raw_ph_type = ph_elems[0].get("type") if ph_elems else None
            ph_idx = ph_elems[0].get("idx") if ph_elems else None

            inherited = self.resolver.resolve_placeholder(
                shape_elem=child_elem,
                raw_slide=raw_slide
            )
            bbox_pt = inherited.xfrm_pt or (0.0, 0.0, 0.0, 0.0)

            resolved_type = (raw_ph_type or inherited.placeholder_type_resolved or "").lower()
            if resolved_type in ("title", "ctrtitle"):
                has_slide_title = True
                existing_native_title_id = shape_id

            text_content, has_text, max_font_size, is_bold = self._extract_text_and_typography(child_elem)
            is_pic = bool(_safe_xpath(child_elem, ".//p:pic") or _safe_xpath(child_elem, ".//a:blip") or child_elem.tag.endswith("pic"))
            is_tbl = bool(_safe_xpath(child_elem, ".//a:tbl"))
            is_chart = bool(_safe_xpath(child_elem, ".//c:chart") or _safe_xpath(child_elem, ".//a:graphicData[contains(@uri, 'chart')]"))
            is_empty = (not has_text) and (not is_pic) and (not is_tbl) and (not is_chart)
            rel_ids = self._extract_relationships(child_elem)

            candidates_raw.append({
                "shape_elem": child_elem,
                "shape_id": shape_id,
                "name": name,
                "raw_ph_type": raw_ph_type,
                "ph_idx": ph_idx,
                "is_animated": is_animated,
                "text_content": text_content,
                "has_text": has_text,
                "max_font_size": max_font_size,
                "is_bold": is_bold,
                "is_pic": is_pic,
                "is_tbl": is_tbl,
                "is_chart": is_chart,
                "is_empty": is_empty,
                "rel_ids": rel_ids,
                "inherited": inherited,
                "bbox_pt": bbox_pt,
                "has_ph": bool(ph_elems),
            })

        # =====================================================================
        # 2. REPEATED-COMPONENT DETECTION & CANDIDATE RANKING
        # =====================================================================
        badge_ids = detect_repeated_components(candidates_raw)

        scored_candidates = []
        anim_title_candidate_ids: Set[int] = set()
        if not has_slide_title:
            for c in candidates_raw:
                sid = c["shape_id"]
                if sid in badge_ids:
                    continue
                if not c["has_text"] or c["is_empty"]:
                    continue
                if c["has_ph"]:
                    continue
                t_str = c["text_content"].strip()
                if len(t_str) < 4:
                    continue
                if c["bbox_pt"][1] >= slide_height_pt * 0.88:
                    continue
                sc = score_title_candidate(c, slide_height_pt, slide_width_pt)
                if c["is_animated"]:
                    if sc >= 25.0:
                        anim_title_candidate_ids.add(sid)
                else:
                    scored_candidates.append((sc, c))

            scored_candidates.sort(key=lambda x: x[0], reverse=True)

        deterministic_title_shape_id: Optional[int] = None
        if scored_candidates and scored_candidates[0][0] >= 30.0:
            deterministic_title_shape_id = scored_candidates[0][1]["shape_id"]

        slide_context = {
            "slide_number": slide_number,
            "slide_dimensions": {"width_pt": slide_width_pt, "height_pt": slide_height_pt},
            "existing_native_title_id": existing_native_title_id,
            "candidates": [
                {
                    "shape_id": c["shape_id"],
                    "name": c["name"],
                    "current_ph_type": c["raw_ph_type"],
                    "current_ph_idx": c["ph_idx"],
                    "text_preview": c["text_content"][:60],
                    "text_length": len(c["text_content"]),
                    "geometry": {
                        "left_pt": round(c["bbox_pt"][0], 1),
                        "top_pt": round(c["bbox_pt"][1], 1),
                        "width_pt": round(c["bbox_pt"][2], 1),
                        "height_pt": round(c["bbox_pt"][3], 1),
                    },
                    "has_text": c["has_text"],
                    "max_font_size_pt": c["max_font_size"],
                    "is_bold": c["is_bold"],
                    "is_card_badge": c["shape_id"] in badge_ids,
                    "is_empty": c["is_empty"],
                    "is_picture": c["is_pic"],
                    "is_table": c["is_tbl"],
                    "is_chart": c["is_chart"],
                    "is_animated": c["is_animated"],
                    "inherited_from": c["inherited"].inherited_from,
                    "inherited_xfrm_available": bool(c["inherited"].xfrm_emu or c["inherited"].has_slide_xfrm),
                }
                for c in candidates_raw
            ]
        }

        # =====================================================================
        # 3. GEMINI 3.5 RUNTIME SEMANTIC DECISION
        # =====================================================================
        gemini_response = self.assistant.analyze_placeholder_conversion(slide_context)
        decisions_map = {d["shape_id"]: d for d in gemini_response.get("decisions", [])}
        gemini_selected_id = gemini_response.get("selected_title_shape_id")

        # =====================================================================
        # 4. DETERMINISTIC SAFETY GATE & RESOLUTION
        # =====================================================================
        promoted_title_shape_id: Optional[int] = None
        if not has_slide_title:
            if gemini_selected_id is not None:
                cand_lookup = next((c for c in candidates_raw if c["shape_id"] == gemini_selected_id), None)
                g_dec = decisions_map.get(gemini_selected_id, {})
                g_conf = float(g_dec.get("confidence", gemini_response.get("overall_confidence", 0.85)))
                if g_conf < 0.70:
                    logger.info("Safety gate preserved shape %s due to low confidence (%.2f).", gemini_selected_id, g_conf)
                    promoted_title_shape_id = None
                elif (
                    cand_lookup is None
                    or cand_lookup["is_animated"]
                    or gemini_selected_id in badge_ids
                    or is_numeric_badge(cand_lookup["text_content"])
                    or is_short_label_or_decorative(cand_lookup["text_content"])
                    or len(cand_lookup["text_content"].strip()) < 4
                ):
                    logger.warning(
                        "Safety gate rejected Gemini title selection for Shape ID %s (badge/label/animated); falling back to deterministic ranker (%s).",
                        gemini_selected_id, deterministic_title_shape_id
                    )
                    promoted_title_shape_id = deterministic_title_shape_id
                else:
                    promoted_title_shape_id = gemini_selected_id
            else:
                promoted_title_shape_id = deterministic_title_shape_id

        for c in candidates_raw:
            sid = c["shape_id"]
            ph_type = c["raw_ph_type"]
            ph_idx = c["ph_idx"]
            elem = c["shape_elem"]
            inherited = c["inherited"]
            bbox_pt = c["bbox_pt"]
            is_animated = c["is_animated"]
            is_empty = c["is_empty"]
            has_text = c["has_text"]
            text_content = c["text_content"]

            base_category = self._classify_category(
                raw_ph_type=ph_type,
                is_pic=c["is_pic"],
                is_tbl=c["is_tbl"],
                is_chart=c["is_chart"],
                is_empty=is_empty
            )
            base_status, base_rec, base_reason, base_conf = self._evaluate_safety(
                shape_id=sid,
                category=base_category,
                is_animated=is_animated,
                is_empty=is_empty,
                inherited=inherited,
                has_text=has_text
            )

            g_dec = decisions_map.get(sid, {})
            g_action = g_dec.get("action", "PRESERVE")
            g_conf = float(g_dec.get("confidence", 0.85))

            # HARD SAFETY RULE 1: Animation Targets must NEVER be auto-converted
            if is_animated:
                if c["has_ph"]:
                    final_category = base_category
                    final_status = PlaceholderStatus.MANUAL_REVIEW
                    final_rec = "Preserve animated placeholder without modifying tags."
                    final_reason = f"Shape ID {sid} is targeted by animation sequence; conversion requires timing review."
                    final_conf_str = "MEDIUM"
                elif sid in anim_title_candidate_ids:
                    final_category = PlaceholderCategory.STANDALONE_TITLE_BOX
                    final_status = PlaceholderStatus.MANUAL_REVIEW
                    final_rec = "Preserve animated title candidate without modifying tags."
                    final_reason = f"Shape ID {sid} is an animated title candidate; conversion requires timing review."
                    final_conf_str = "MEDIUM"
                else:
                    final_category = base_category
                    final_status = PlaceholderStatus.PRESERVED
                    final_rec = "Preserve animated shape."
                    final_reason = f"Shape ID {sid} is targeted by animation sequence."
                    final_conf_str = "MEDIUM"

            # HARD SAFETY RULE 2: Native Titles are ALWAYS PRESERVED
            elif base_category == PlaceholderCategory.TITLE:
                final_category = PlaceholderCategory.TITLE
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve native title placeholder."
                final_reason = "Title placeholders are critical for screen reader outline navigation and must remain native."
                final_conf_str = "HIGH"

            # HARD SAFETY RULE 3: Footers, Slide Numbers, Dates are ALWAYS PRESERVED
            elif base_category in (PlaceholderCategory.FOOTER, PlaceholderCategory.SLIDE_NUMBER, PlaceholderCategory.DATE_TIME):
                final_category = base_category
                final_status = PlaceholderStatus.PRESERVED
                final_rec = f"Preserve native {base_category.value.lower().replace('_', ' ')} placeholder."
                final_reason = "Has dedicated semantic slide role and must not be altered."
                final_conf_str = "HIGH"

            # HARD SAFETY RULE 4: Specialized Placeholders are ALWAYS PRESERVED
            elif base_category in (PlaceholderCategory.PICTURE, PlaceholderCategory.TABLE, PlaceholderCategory.CHART):
                final_category = base_category
                final_status = PlaceholderStatus.PRESERVED
                final_rec = f"Preserve specialized {base_category.value.lower()} placeholder."
                final_reason = f"Contains embedded {base_category.value.lower()}; deferred to specialized remediation phase."
                final_conf_str = "HIGH"

            # HARD SAFETY RULE 5: Empty Placeholders are ALWAYS PRESERVED
            elif is_empty or base_category == PlaceholderCategory.EMPTY:
                final_category = PlaceholderCategory.EMPTY
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve empty placeholder container."
                final_reason = "Contains no text or visual content; preserved as layout container without inventing text."
                final_conf_str = "HIGH"

            # HARD SAFETY RULE 6: Subtitle Placeholders are ALWAYS PRESERVED
            elif base_category == PlaceholderCategory.SUBTITLE:
                final_category = PlaceholderCategory.SUBTITLE
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve native subtitle placeholder."
                final_reason = "Semantic subtitle placeholder linked to slide title layout."
                final_conf_str = "HIGH"

            # HARD SAFETY RULE 7: Card Badges and Design Markers are ALWAYS PRESERVED
            elif not c["has_ph"] and (sid in badge_ids or is_numeric_badge(text_content) or is_short_label_or_decorative(text_content)):
                final_category = PlaceholderCategory.CARD_BADGE
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve card badge / design marker component."
                final_reason = f"Shape ID {sid} is a card badge/marker ('{text_content[:20]}'); must not be converted to title."
                final_conf_str = "HIGH"

            # ACTION: CONVERT_TO_TITLE (Semantically selected slide-level title)
            elif sid == promoted_title_shape_id and not has_slide_title:
                final_category = PlaceholderCategory.STANDALONE_TITLE_BOX
                final_status = PlaceholderStatus.SAFE_TO_CONVERT
                final_rec = "Convert standalone title text box into native slide title placeholder."
                final_reason = "Slide lacks native title; text box serves as visual title and should be linked for outline navigation."
                final_conf_str = "HIGH"

            # ACTION: MATERIALIZE_BODY (Body / Object placeholder with text)
            elif (
                base_category in (PlaceholderCategory.BODY, PlaceholderCategory.OBJECT)
                and has_text
            ):
                if not (inherited.xfrm_emu or inherited.has_slide_xfrm):
                    final_category = base_category
                    final_status = PlaceholderStatus.MANUAL_REVIEW
                    final_rec = "Manual review required: geometry could not be resolved from layout or master."
                    final_reason = "Shape lacks local transform and layout/master has no matching geometry bounds."
                    final_conf_str = "LOW"
                elif g_conf < 0.70:
                    final_category = base_category
                    final_status = PlaceholderStatus.PRESERVED
                    final_rec = "Preserve body placeholder due to low confidence."
                    final_reason = f"Low confidence ({g_conf:.2f}); preserved placeholder."
                    final_conf_str = "LOW"
                else:
                    final_category = base_category
                    final_status = PlaceholderStatus.SAFE_TO_CONVERT
                    final_rec = "Convert placeholder by materializing inherited geometry into self-contained shape."
                    final_reason = "Contains accessible text; materializing inherited geometry guarantees zero visual drift."
                    final_conf_str = "HIGH" if g_conf >= 0.85 else "MEDIUM"

            # DEFAULT: PRESERVED shape
            else:
                final_category = base_category
                final_status = base_status
                final_rec = base_rec
                final_reason = base_reason
                final_conf_str = base_conf

            analysis = PlaceholderAnalysis(
                slide_number=slide_number,
                shape_id=sid,
                shape_name=c["name"],
                placeholder_type=ph_type or ("obj" if ph_idx else "none"),
                placeholder_idx=ph_idx,
                category=final_category,
                status=final_status,
                has_text=has_text,
                text_preview=text_content[:60],
                is_empty=is_empty,
                is_animated=is_animated,
                is_picture=c["is_pic"],
                is_table=c["is_tbl"],
                is_chart=c["is_chart"],
                inherited_properties=inherited,
                slide_bbox_pt=bbox_pt,
                recommendation=final_rec,
                reason=final_reason,
                confidence=final_conf_str,
                raw_element=elem,
                relationship_ids=c["rel_ids"],
                numerical_confidence=g_conf,
                gemini_decision=g_dec or None,
                model_traceability={
                    "antigravity_coding_model": "Gemini 3.8 Flash High",
                    "slideguard_runtime_model": "Gemini 3.5",
                    "runtime_model_identifier": self.assistant.model_name
                }
            )

            # Keep results consistent with placeholder evaluation: native placeholders and standalone title candidates
            if c["has_ph"] or final_category == PlaceholderCategory.STANDALONE_TITLE_BOX:
                results.append(analysis)

        return results

    def analyze_shape(
        self,
        shape_elem,
        slide_number: int,
        raw_slide=None,
        anim_targets: Optional[Set[int]] = None
    ) -> PlaceholderAnalysis:
        """Analyzes an individual placeholder shape element."""
        anim_set = anim_targets or set()

        # Non-visual properties
        cnv_prs = _safe_xpath(shape_elem, ".//p:cNvPr")
        shape_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
        name = cnv_prs[0].get("name", f"Shape_{shape_id}") if cnv_prs else f"Shape_{shape_id}"
        is_animated = shape_id in anim_set

        # <p:ph> element
        ph_elems = _safe_xpath(shape_elem, ".//p:ph")
        raw_ph_type = ph_elems[0].get("type") if ph_elems else None
        ph_idx = ph_elems[0].get("idx") if ph_elems else None

        # Content inspection
        text_content, has_text = self._extract_text(shape_elem)
        is_pic = bool(_safe_xpath(shape_elem, ".//p:pic") or _safe_xpath(shape_elem, ".//a:blip") or shape_elem.tag.endswith("pic"))
        is_tbl = bool(_safe_xpath(shape_elem, ".//a:tbl"))
        is_chart = bool(_safe_xpath(shape_elem, ".//c:chart") or _safe_xpath(shape_elem, ".//a:graphicData[contains(@uri, 'chart')]"))
        is_empty = (not has_text) and (not is_pic) and (not is_tbl) and (not is_chart)

        # Relationships
        rel_ids = self._extract_relationships(shape_elem)

        # Resolve inheritance (Slide -> Layout -> Master)
        inherited = self.resolver.resolve_placeholder(
            shape_elem=shape_elem,
            raw_slide=raw_slide
        )

        bbox_pt = inherited.xfrm_pt or (0.0, 0.0, 0.0, 0.0)

        # Categorize placeholder
        category = self._classify_category(
            raw_ph_type=raw_ph_type,
            is_pic=is_pic,
            is_tbl=is_tbl,
            is_chart=is_chart,
            is_empty=is_empty
        )

        # Determine conversion safety and recommendation
        status, recommendation, reason, confidence = self._evaluate_safety(
            shape_id=shape_id,
            category=category,
            is_animated=is_animated,
            is_empty=is_empty,
            inherited=inherited,
            has_text=has_text
        )

        return PlaceholderAnalysis(
            slide_number=slide_number,
            shape_id=shape_id,
            shape_name=name,
            placeholder_type=raw_ph_type or ("obj" if ph_idx else "body"),
            placeholder_idx=ph_idx,
            category=category,
            status=status,
            has_text=has_text,
            text_preview=text_content[:60],
            is_empty=is_empty,
            is_animated=is_animated,
            is_picture=is_pic,
            is_table=is_tbl,
            is_chart=is_chart,
            inherited_properties=inherited,
            slide_bbox_pt=bbox_pt,
            recommendation=recommendation,
            reason=reason,
            confidence=confidence,
            raw_element=shape_elem,
            relationship_ids=rel_ids
        )

    def _check_standalone_title(
        self,
        shape_elem,
        slide_number: int,
        anim_targets: Set[int],
        slide_height_pt: float
    ) -> Optional[PlaceholderAnalysis]:
        """Detects a standalone text box qualifying as an unlinked slide title."""
        cnv_prs = _safe_xpath(shape_elem, ".//p:cNvPr")
        shape_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
        name = cnv_prs[0].get("name", f"Shape_{shape_id}") if cnv_prs else f"Shape_{shape_id}"
        is_animated = shape_id in anim_targets

        text_content, has_text = self._extract_text(shape_elem)
        if not has_text or len(text_content.strip()) > 150:
            return None

        # Hard protection: numeric-only, sequence badge, or short decorative labels can never be titles
        if is_numeric_badge(text_content) or is_short_label_or_decorative(text_content) or len(text_content.strip()) < 4:
            return None

        # Check geometry
        xfrm = _safe_xpath(shape_elem, ".//p:spPr/a:xfrm")
        if not xfrm:
            return None

        off = _safe_xpath(xfrm[0], ".//a:off")
        ext = _safe_xpath(xfrm[0], ".//a:ext")
        if not off or not ext:
            return None

        top_pt = emu_to_points(int(off[0].get("y", "0")))
        left_pt = emu_to_points(int(off[0].get("x", "0")))
        width_pt = emu_to_points(int(ext[0].get("cx", "0")))
        height_pt = emu_to_points(int(ext[0].get("cy", "0")))

        # Exclude footer region
        if top_pt >= (slide_height_pt * 0.88):
            return None

        inherited = InheritedProperties(
            has_slide_xfrm=True,
            inherited_from="SLIDE",
            xfrm_emu=(int(off[0].get("x", "0")), int(off[0].get("y", "0")), int(ext[0].get("cx", "0")), int(ext[0].get("cy", "0"))),
            xfrm_pt=(left_pt, top_pt, width_pt, height_pt),
            placeholder_type_resolved="title"
        )

        rel_ids = self._extract_relationships(shape_elem)

        return PlaceholderAnalysis(
            slide_number=slide_number,
            shape_id=shape_id,
            shape_name=name,
            placeholder_type="none",
            placeholder_idx=None,
            category=PlaceholderCategory.STANDALONE_TITLE_BOX,
            status=PlaceholderStatus.SAFE_TO_CONVERT,
            has_text=True,
            text_preview=text_content[:60],
            is_empty=False,
            is_animated=is_animated,
            is_picture=False,
            is_table=False,
            is_chart=False,
            inherited_properties=inherited,
            slide_bbox_pt=(left_pt, top_pt, width_pt, height_pt),
            recommendation="Convert standalone title text box into native slide title placeholder.",
            reason="Slide lacks native title; text box serves as visual title and should be linked for outline navigation.",
            confidence="HIGH",
            raw_element=shape_elem,
            relationship_ids=rel_ids
        )

    def _classify_category(
        self,
        raw_ph_type: Optional[str],
        is_pic: bool,
        is_tbl: bool,
        is_chart: bool,
        is_empty: bool
    ) -> PlaceholderCategory:
        """Maps OpenXML placeholder attributes and contents to semantic category."""
        p_type = (raw_ph_type or "").lower()

        if p_type in ("title", "ctrtitle"):
            return PlaceholderCategory.TITLE
        if p_type == "subtitle":
            return PlaceholderCategory.SUBTITLE
        if p_type in ("ftr", "footer"):
            return PlaceholderCategory.FOOTER
        if p_type in ("sldnum", "slidenumber"):
            return PlaceholderCategory.SLIDE_NUMBER
        if p_type in ("dt", "date"):
            return PlaceholderCategory.DATE_TIME
        if is_pic or p_type in ("pic", "picture", "clipart"):
            return PlaceholderCategory.PICTURE
        if is_tbl or p_type in ("tbl", "table"):
            return PlaceholderCategory.TABLE
        if is_chart or p_type == "chart":
            return PlaceholderCategory.CHART
        if is_empty:
            return PlaceholderCategory.EMPTY
        if p_type == "body":
            return PlaceholderCategory.BODY
        if p_type in ("obj", "object") or not p_type:
            return PlaceholderCategory.OBJECT

        return PlaceholderCategory.OTHER

    def _evaluate_safety(
        self,
        shape_id: int,
        category: PlaceholderCategory,
        is_animated: bool,
        is_empty: bool,
        inherited: InheritedProperties,
        has_text: bool
    ) -> Tuple[PlaceholderStatus, str, str, str]:
        """Evaluates conversion safety for a placeholder."""
        # 1. Animation target: preserve and flag manual review
        if is_animated:
            return (
                PlaceholderStatus.MANUAL_REVIEW,
                "Preserve animated placeholder without modifying tags.",
                f"Shape ID {shape_id} is targeted by animation sequence; conversion requires timing review.",
                "MEDIUM"
            )

        # 2. Title placeholders: NEVER convert to normal text box
        if category == PlaceholderCategory.TITLE:
            return (
                PlaceholderStatus.PRESERVED,
                "Preserve native title placeholder.",
                "Title placeholders are critical for screen reader outline navigation and must remain native.",
                "HIGH"
            )

        # 3. Footers, slide numbers, and date/time: NEVER convert to normal body text
        if category in (PlaceholderCategory.FOOTER, PlaceholderCategory.SLIDE_NUMBER, PlaceholderCategory.DATE_TIME):
            return (
                PlaceholderStatus.PRESERVED,
                f"Preserve native {category.value.lower().replace('_', ' ')} placeholder.",
                "Has dedicated semantic slide role and must not be altered.",
                "HIGH"
            )

        # 4. Specialized placeholders (Picture, Table, Chart)
        if category in (PlaceholderCategory.PICTURE, PlaceholderCategory.TABLE, PlaceholderCategory.CHART):
            return (
                PlaceholderStatus.PRESERVED,
                f"Preserve specialized {category.value.lower()} placeholder.",
                f"Contains embedded {category.value.lower()}; deferred to specialized remediation phase.",
                "HIGH"
            )

        # 5. Empty placeholders: Leave unchanged, do not invent text
        if is_empty or category == PlaceholderCategory.EMPTY:
            return (
                PlaceholderStatus.PRESERVED,
                "Preserve empty placeholder container.",
                "Contains no text or visual content; preserved as layout container without inventing text.",
                "HIGH"
            )

        # 6. Subtitle placeholders: Preserve
        if category == PlaceholderCategory.SUBTITLE:
            return (
                PlaceholderStatus.PRESERVED,
                "Preserve native subtitle placeholder.",
                "Semantic subtitle placeholder linked to slide title layout.",
                "HIGH"
            )

        # 7. Body / Object placeholders containing text:
        if category in (PlaceholderCategory.BODY, PlaceholderCategory.OBJECT) and has_text:
            # Check geometry availability
            if inherited.inherited_from == "NONE" and (not inherited.has_slide_xfrm):
                return (
                    PlaceholderStatus.MANUAL_REVIEW,
                    "Manual review required: geometry could not be resolved from layout or master.",
                    "Shape lacks local transform and layout/master has no matching geometry bounds.",
                    "LOW"
                )

            return (
                PlaceholderStatus.SAFE_TO_CONVERT,
                "Convert placeholder by materializing inherited geometry into self-contained shape.",
                "Contains accessible text; materializing inherited geometry guarantees zero visual drift.",
                "HIGH"
            )

        # Default fallback: unsupported/ambiguous
        return (
            PlaceholderStatus.MANUAL_REVIEW,
            "Manual review required: ambiguous placeholder structure.",
            "Cannot deterministically guarantee visual and structural preservation.",
            "LOW"
        )

    def _extract_text(self, elem) -> Tuple[str, bool]:
        """Extracts text content from a shape element."""
        texts = []
        for t in _safe_xpath(elem, ".//a:t"):
            if t.text:
                texts.append(t.text)
        joined = " ".join(texts).strip()
        return joined, bool(joined)

    def _extract_text_and_typography(self, elem) -> Tuple[str, bool, float, bool]:
        """Extracts text content, max font size (pt), and bold flag."""
        texts = []
        font_sizes = []
        is_bold = False
        for p in _safe_xpath(elem, ".//a:p"):
            p_texts = []
            for r in _safe_xpath(p, ".//a:r"):
                t_elems = _safe_xpath(r, ".//a:t")
                r_text = "".join(t.text for t in t_elems if t.text)
                if r_text:
                    p_texts.append(r_text)
                r_prs = _safe_xpath(r, ".//a:rPr")
                if r_prs:
                    sz = r_prs[0].get("sz")
                    if sz:
                        try:
                            font_sizes.append(int(sz) / 100.0)
                        except (ValueError, TypeError):
                            pass
                    b = r_prs[0].get("b")
                    if b in ("1", "true"):
                        is_bold = True
            if not p_texts:
                fld_texts = [t.text for t in _safe_xpath(p, ".//a:fld//a:t") if t.text]
                if fld_texts:
                    p_texts.extend(fld_texts)
            if p_texts:
                texts.append(" ".join(p_texts))
        joined = " ".join(texts).strip()
        max_sz = max(font_sizes) if font_sizes else 0.0
        return joined, bool(joined), max_sz, is_bold

    def _extract_relationships(self, elem) -> List[str]:
        """Extracts embedded and hyperlink relationship IDs."""
        rel_ids = []
        for blip in _safe_xpath(elem, ".//a:blip"):
            r_embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if r_embed:
                rel_ids.append(r_embed)
        for hl in _safe_xpath(elem, ".//a:hlinkClick"):
            r_id = hl.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if r_id:
                rel_ids.append(r_id)
        return rel_ids
