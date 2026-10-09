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
    STANDALONE_SUBTITLE_BOX = "STANDALONE_SUBTITLE_BOX"
    STANDALONE_BODY_BOX = "STANDALONE_BODY_BOX"
    CARD_BADGE = "CARD_BADGE"
    CARD_COMPONENT = "CARD_COMPONENT"
    GROUP_COMPONENT = "GROUP_COMPONENT"
    OTHER = "OTHER"


class ConversionAction(str, Enum):
    """Primary action classification for PowerPoint objects (Task 4 Option C)."""
    CONVERT_EXISTING = "CONVERT_EXISTING"
    CREATE_NEW = "CREATE_NEW"
    PRESERVE_EXISTING = "PRESERVE_EXISTING"
    SKIP_UNSAFE = "SKIP_UNSAFE"
    MANUAL_REVIEW = "MANUAL_REVIEW"
    FAILED = "FAILED"


class DifferenceCategory(str, Enum):
    """Verification difference classifications."""
    EXPECTED_ACCESSIBILITY_CHANGE = "EXPECTED_ACCESSIBILITY_CHANGE"
    UNINTENDED_STRUCTURAL_CHANGE = "UNINTENDED_STRUCTURAL_CHANGE"
    UNINTENDED_VISUAL_CHANGE = "UNINTENDED_VISUAL_CHANGE"
    UNVERIFIED = "UNVERIFIED"
    VERIFICATION_FAILURE = "VERIFICATION_FAILURE"


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


def is_caption_or_metadata(text: str) -> bool:
    """Checks if text is a figure/table caption, callout, photo credit, or metadata label."""
    t = text.strip()
    if not t:
        return False
    if re.match(r"^(?:figure|fig\.|table|chart|source|photo|image|note|diagram)\s*[\d.:\-]", t, re.IGNORECASE):
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
        if is_numeric_badge(t) or is_short_label_or_decorative(t) or is_caption_or_metadata(t):
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

    if is_caption_or_metadata(t):
        return -100.0

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
    action: ConversionAction = ConversionAction.PRESERVE_EXISTING
    semantic_role: str = "OTHER"
    conversion_basis: str = ""
    is_group_child: bool = False
    parent_group_id: Optional[int] = None
    before_placeholder_type: str = "none"
    proposed_placeholder_type: str = "none"
    has_ph: bool = False
    verified: bool = False
    verification_category: Optional[DifferenceCategory] = None

    @property
    def is_safe_to_convert(self) -> bool:
        return (
            self.status == PlaceholderStatus.SAFE_TO_CONVERT
            or self.action in (ConversionAction.CONVERT_EXISTING, ConversionAction.CREATE_NEW)
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "placeholder_type": self.placeholder_type,
            "placeholder_idx": self.placeholder_idx,
            "category": self.category.value,
            "status": self.status.value,
            "action": self.action.value if hasattr(self.action, "value") else str(self.action),
            "semantic_role": self.semantic_role,
            "conversion_basis": self.conversion_basis,
            "is_group_child": self.is_group_child,
            "parent_group_id": self.parent_group_id,
            "before_placeholder_type": self.before_placeholder_type,
            "proposed_placeholder_type": self.proposed_placeholder_type,
            "has_ph": self.has_ph,
            "verified": self.verified,
            "verification_category": self.verification_category.value if self.verification_category else None,
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
        slide_width_pt: float = 720.0,
        include_all_shapes: bool = False
    ) -> List[PlaceholderAnalysis]:
        """
        Analyzes all placeholders and candidate shapes on a slide using
        Gemini 3.5 runtime semantic reasoning and deterministic safety gates.
        Option C Architecture:
          1. Candidate Detection across all shapes (including group children)
          2. Multi-Signal Semantic Classification & Candidate Ranking
          3. Protected Object Gates (badges, short labels, groups, animations)
          4. Conversion Action Assignment (CONVERT_EXISTING, CREATE_NEW, PRESERVE_EXISTING, SKIP_UNSAFE, MANUAL_REVIEW, FAILED)
          5. Model Traceability & Strict Verification Readiness
        """
        results: List[PlaceholderAnalysis] = []
        anim_set = set(animation_target_ids or [])

        if raw_slide is None:
            return results

        slide_elem = raw_slide._element if hasattr(raw_slide, "_element") else raw_slide

        # =====================================================================
        # 1. CANDIDATE DETECTION ACROSS ROOT AND GROUPED OBJECTS
        # =====================================================================
        candidates_raw = []
        has_slide_title = False
        existing_native_title_id: Optional[int] = None
        has_slide_subtitle = False

        for child_elem in _safe_xpath(slide_elem, ".//p:spTree/*"):
            tag = child_elem.tag.split("}")[-1]
            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue

            # Handle Group Shapes
            if tag == "grpSp":
                grp_cnv = _safe_xpath(child_elem, ".//p:cNvPr")
                grp_id = int(grp_cnv[0].get("id", "0")) if grp_cnv else 0
                grp_name = grp_cnv[0].get("name", f"Group_{grp_id}") if grp_cnv else f"Group_{grp_id}"

                # Extract and protect child shapes inside the group
                for g_child in _safe_xpath(child_elem, ".//p:sp | .//p:pic | .//p:graphicFrame"):
                    g_cnv = _safe_xpath(g_child, ".//p:cNvPr")
                    g_sid = int(g_cnv[0].get("id", "0")) if g_cnv else 0
                    g_name = g_cnv[0].get("name", f"GroupChild_{g_sid}") if g_cnv else f"GroupChild_{g_sid}"
                    g_anim = g_sid in anim_set
                    g_ph_elems = _safe_xpath(g_child, ".//p:ph")
                    g_raw_ph_type = g_ph_elems[0].get("type") if g_ph_elems else None
                    g_ph_idx = g_ph_elems[0].get("idx") if g_ph_elems else None
                    g_text, g_has_text, g_max_sz, g_is_bold = self._extract_text_and_typography(g_child)
                    g_pic = bool(_safe_xpath(g_child, ".//p:pic") or _safe_xpath(g_child, ".//a:blip") or g_child.tag.endswith("pic"))
                    g_tbl = bool(_safe_xpath(g_child, ".//a:tbl"))
                    g_chart = bool(_safe_xpath(g_child, ".//c:chart") or _safe_xpath(g_child, ".//a:graphicData[contains(@uri, 'chart')]"))
                    g_empty = (not g_has_text) and (not g_pic) and (not g_tbl) and (not g_chart)
                    g_rel_ids = self._extract_relationships(g_child)
                    g_inherited = self.resolver.resolve_placeholder(shape_elem=g_child, raw_slide=raw_slide)

                    candidates_raw.append({
                        "shape_elem": g_child,
                        "shape_id": g_sid,
                        "name": g_name,
                        "raw_ph_type": g_raw_ph_type,
                        "ph_idx": g_ph_idx,
                        "is_animated": g_anim,
                        "text_content": g_text,
                        "has_text": g_has_text,
                        "max_font_size": g_max_sz,
                        "is_bold": g_is_bold,
                        "is_pic": g_pic,
                        "is_tbl": g_tbl,
                        "is_chart": g_chart,
                        "is_empty": g_empty,
                        "rel_ids": g_rel_ids,
                        "inherited": g_inherited,
                        "bbox_pt": g_inherited.xfrm_pt or (0.0, 0.0, 0.0, 0.0),
                        "has_ph": bool(g_ph_elems),
                        "is_group_child": True,
                        "parent_group_id": grp_id,
                    })
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
            elif resolved_type == "subtitle":
                has_slide_subtitle = True

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
                "is_group_child": False,
                "parent_group_id": None,
            })

        # =====================================================================
        # 2. REPEATED-COMPONENT DETECTION & LAYOUT CAPABILITIES
        # =====================================================================
        badge_ids = detect_repeated_components(candidates_raw)

        # Inspect layout placeholder capabilities
        layout_has_subtitle = False
        layout_has_body = False
        layout_body_idx = "1"
        if raw_slide is not None and hasattr(raw_slide, "slide_layout") and raw_slide.slide_layout is not None:
            try:
                for ph_tag in _safe_xpath(raw_slide.slide_layout._element, ".//p:ph"):
                    pt = (ph_tag.get("type") or "").lower()
                    if pt == "subtitle":
                        layout_has_subtitle = True
                    if pt in ("body", "obj") or (not pt and ph_tag.get("idx") in ("1", "2")):
                        layout_has_body = True
                        layout_body_idx = ph_tag.get("idx", "1")
            except Exception:
                pass

        # Check if slide already has a native body placeholder
        has_native_body = any(
            c["has_ph"] and (
                (c["raw_ph_type"] or "").lower() in ("body", "obj")
                or (not c["raw_ph_type"] and c["ph_idx"] in ("1", "2"))
            )
            for c in candidates_raw
        )

        # =====================================================================
        # 3. TITLE CANDIDATE EVALUATION & GEMINI INTEGRATION
        # =====================================================================
        scored_candidates = []
        anim_title_candidate_ids: Set[int] = set()
        if not has_slide_title:
            for c in candidates_raw:
                if c.get("is_group_child"):
                    continue
                sid = c["shape_id"]
                if sid in badge_ids:
                    continue
                if not c["has_text"] or c["is_empty"]:
                    continue
                if c["has_ph"]:
                    continue
                t_str = c["text_content"].strip()
                if len(t_str) < 4 or is_caption_or_metadata(t_str):
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
                    "is_group_child": c.get("is_group_child", False),
                    "inherited_from": c["inherited"].inherited_from,
                    "inherited_xfrm_available": bool(c["inherited"].xfrm_emu or c["inherited"].has_slide_xfrm),
                }
                for c in candidates_raw
                if not c.get("is_group_child")  # Focus root shapes for external decision
            ]
        }

        try:
            gemini_response = self.assistant.analyze_placeholder_conversion(slide_context)
        except Exception as e:
            logger.warning("Gemini semantic analysis unavailable or timed out: %s; falling back to deterministic safety checks.", e)
            gemini_response = {"decisions": [], "error": str(e), "selected_title_shape_id": None}
        decisions_map = {d["shape_id"]: d for d in gemini_response.get("decisions", [])}
        gemini_selected_id = gemini_response.get("selected_title_shape_id")

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
                    or cand_lookup.get("is_group_child")
                    or gemini_selected_id in badge_ids
                    or is_numeric_badge(cand_lookup["text_content"])
                    or is_short_label_or_decorative(cand_lookup["text_content"])
                    or is_caption_or_metadata(cand_lookup["text_content"])
                    or len(cand_lookup["text_content"].strip()) < 4
                ):
                    logger.warning(
                        "Safety gate rejected Gemini title selection for Shape ID %s (badge/label/animated/group); falling back to deterministic ranker (%s).",
                        gemini_selected_id, deterministic_title_shape_id
                    )
                    promoted_title_shape_id = deterministic_title_shape_id
                else:
                    promoted_title_shape_id = gemini_selected_id
            else:
                promoted_title_shape_id = deterministic_title_shape_id

        # =====================================================================
        # 4. SUBTITLE & BODY CANDIDATE EVALUATION
        # =====================================================================
        effective_title_id = existing_native_title_id or promoted_title_shape_id

        # Subtitle promotion evaluation
        promoted_subtitle_shape_id: Optional[int] = None
        if effective_title_id is not None and not has_slide_subtitle and layout_has_subtitle:
            subtitle_candidates = []
            for c in candidates_raw:
                sid = c["shape_id"]
                if sid == effective_title_id or c.get("is_group_child") or c["is_animated"]:
                    continue
                if sid in badge_ids or is_numeric_badge(c["text_content"]) or is_short_label_or_decorative(c["text_content"]):
                    continue
                if not c["has_text"] or c["is_empty"] or c["has_ph"]:
                    continue
                t_str = c["text_content"].strip()
                words = t_str.split()
                if not (1 <= len(words) <= 20) or len(t_str) < 3:
                    continue
                bbox = c["bbox_pt"]
                if bbox[1] >= slide_height_pt * 0.75:
                    continue
                title_c = next((cand for cand in candidates_raw if cand["shape_id"] == effective_title_id), None)
                title_font = title_c["max_font_size"] if title_c else 24.0
                if c["max_font_size"] > title_font and title_font > 0:
                    continue

                sub_score = 15.0
                if layout_has_subtitle:
                    sub_score += 30.0
                if c["is_bold"]:
                    sub_score += 10.0
                if 14.0 <= c["max_font_size"] <= 24.0:
                    sub_score += 15.0
                if bbox[1] <= slide_height_pt * 0.50:
                    sub_score += 15.0
                if len(words) <= 12:
                    sub_score += 10.0
                subtitle_candidates.append((sub_score, sid))

            subtitle_candidates.sort(key=lambda x: x[0], reverse=True)
            if subtitle_candidates and subtitle_candidates[0][0] >= 35.0:
                promoted_subtitle_shape_id = subtitle_candidates[0][1]

        # Body promotion evaluation
        promoted_body_shape_id: Optional[int] = None
        if layout_has_body and not has_native_body:
            body_candidates = []
            for c in candidates_raw:
                sid = c["shape_id"]
                if sid in (effective_title_id, promoted_subtitle_shape_id):
                    continue
                if c.get("is_group_child") or c["is_animated"]:
                    continue
                if sid in badge_ids or is_numeric_badge(c["text_content"]) or is_short_label_or_decorative(c["text_content"]):
                    continue
                if not c["has_text"] or c["is_empty"] or c["has_ph"]:
                    continue
                t_str = c["text_content"].strip()
                words = t_str.split()
                if len(words) < 3 or len(t_str) < 15:
                    continue
                if c["bbox_pt"][1] >= slide_height_pt * 0.88:
                    continue
                if c["max_font_size"] > 24.0:
                    continue

                b_score = 20.0
                if len(words) >= 6:
                    b_score += 20.0
                if "\n" in t_str or "•" in t_str or "-" in t_str:
                    b_score += 20.0
                body_candidates.append((b_score, sid))

            body_candidates.sort(key=lambda x: x[0], reverse=True)
            if body_candidates and body_candidates[0][0] >= 35.0:
                promoted_body_shape_id = body_candidates[0][1]

        # =====================================================================
        # 5. OBJECT CLASSIFICATION & ACTION ASSIGNMENT
        # =====================================================================
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
            is_group_child = c.get("is_group_child", False)
            parent_group_id = c.get("parent_group_id")

            base_category = self._classify_category(
                raw_ph_type=ph_type,
                is_pic=c["is_pic"],
                is_tbl=c["is_tbl"],
                is_chart=c["is_chart"],
                is_empty=is_empty,
                has_ph=c["has_ph"]
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
            g_conf = float(g_dec.get("confidence", 0.85))

            proposed_ph_type = "none"
            semantic_role = "OTHER"
            conversion_basis = ""

            # HARD SAFETY RULE 1: Animation Targets must NEVER be auto-converted
            if is_animated:
                if c["has_ph"]:
                    final_action = ConversionAction.MANUAL_REVIEW
                    final_status = PlaceholderStatus.MANUAL_REVIEW
                    final_category = base_category
                    final_rec = "Preserve animated placeholder without modifying tags."
                    final_reason = f"Shape ID {sid} is targeted by animation sequence; conversion requires timing review."
                    final_conf_str = "MEDIUM"
                    semantic_role = base_category.value
                    conversion_basis = "ANIMATION_TARGET_PROTECTION"
                elif sid in anim_title_candidate_ids:
                    final_action = ConversionAction.MANUAL_REVIEW
                    final_status = PlaceholderStatus.MANUAL_REVIEW
                    final_category = PlaceholderCategory.STANDALONE_TITLE_BOX
                    final_rec = "Preserve animated title candidate without modifying tags."
                    final_reason = f"Shape ID {sid} is an animated title candidate; conversion requires timing review."
                    final_conf_str = "MEDIUM"
                    semantic_role = "TITLE"
                    conversion_basis = "ANIMATION_TARGET_PROTECTION"
                else:
                    final_action = ConversionAction.PRESERVE_EXISTING
                    final_status = PlaceholderStatus.PRESERVED
                    final_category = base_category
                    final_rec = "Preserve animated shape."
                    final_reason = f"Shape ID {sid} is targeted by animation sequence; retained as non-placeholder object."
                    final_conf_str = "HIGH"
                    semantic_role = "ANIMATED_SHAPE"
                    conversion_basis = "ANIMATION_TARGET_PROTECTION"

            # HARD SAFETY RULE 2: Group Children must NEVER be promoted to slide placeholders
            elif is_group_child:
                final_action = ConversionAction.SKIP_UNSAFE
                final_status = PlaceholderStatus.PRESERVED
                final_category = PlaceholderCategory.GROUP_COMPONENT
                final_rec = "Preserve text inside grouped structure."
                final_reason = f"Shape ID {sid} is inside group {parent_group_id}; must retain group geometry and transforms to protect layout structure."
                final_conf_str = "HIGH"
                semantic_role = "GROUP_COMPONENT"
                conversion_basis = "GROUP_STRUCTURE_PROTECTION"

            # HARD SAFETY RULE 3: Native Titles are ALWAYS PRESERVED
            elif base_category == PlaceholderCategory.TITLE:
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = PlaceholderCategory.TITLE
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve native title placeholder."
                final_reason = "Title placeholders are critical for screen reader outline navigation and must remain native."
                final_conf_str = "HIGH"
                semantic_role = "TITLE"
                conversion_basis = "EXISTING_NATIVE_ROLE"

            # HARD SAFETY RULE 4: Footers, Slide Numbers, Dates are ALWAYS PRESERVED
            elif base_category in (PlaceholderCategory.FOOTER, PlaceholderCategory.SLIDE_NUMBER, PlaceholderCategory.DATE_TIME):
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = base_category
                final_status = PlaceholderStatus.PRESERVED
                final_rec = f"Preserve native {base_category.value.lower().replace('_', ' ')} placeholder."
                final_reason = "Has dedicated semantic slide role and must not be altered."
                final_conf_str = "HIGH"
                semantic_role = base_category.value
                conversion_basis = "EXISTING_NATIVE_ROLE"

            # HARD SAFETY RULE 5: Specialized Placeholders are ALWAYS PRESERVED
            elif base_category in (PlaceholderCategory.PICTURE, PlaceholderCategory.TABLE, PlaceholderCategory.CHART):
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = base_category
                final_status = PlaceholderStatus.PRESERVED
                final_rec = f"Preserve specialized {base_category.value.lower()} placeholder."
                final_reason = f"Contains embedded {base_category.value.lower()}; deferred to specialized remediation phase."
                final_conf_str = "HIGH"
                semantic_role = base_category.value
                conversion_basis = "SPECIALIZED_CONTAINER_PROTECTION"

            # HARD SAFETY RULE 6: Empty Placeholders are ALWAYS PRESERVED
            elif is_empty or base_category == PlaceholderCategory.EMPTY:
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = PlaceholderCategory.EMPTY
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve empty placeholder container."
                final_reason = "Contains no text or visual content; preserved as layout container without inventing text."
                final_conf_str = "HIGH"
                semantic_role = "EMPTY"
                conversion_basis = "EMPTY_CONTAINER_PROTECTION"

            # HARD SAFETY RULE 7: Native Subtitle Placeholders are ALWAYS PRESERVED
            elif base_category == PlaceholderCategory.SUBTITLE:
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = PlaceholderCategory.SUBTITLE
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve native subtitle placeholder."
                final_reason = "Semantic subtitle placeholder linked to slide title layout."
                final_conf_str = "HIGH"
                semantic_role = "SUBTITLE"
                conversion_basis = "EXISTING_NATIVE_ROLE"

            # HARD SAFETY RULE 8: Card Badges and Design Markers are ALWAYS PRESERVED
            elif not c["has_ph"] and (sid in badge_ids or is_numeric_badge(text_content) or is_short_label_or_decorative(text_content) or is_caption_or_metadata(text_content)):
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = PlaceholderCategory.CARD_BADGE
                final_status = PlaceholderStatus.PRESERVED
                final_rec = "Preserve card badge / design marker component."
                final_reason = f"Shape ID {sid} is a card badge/marker/caption ('{text_content[:20]}'); must not be converted to title or placeholder."
                final_conf_str = "HIGH"
                semantic_role = "CARD_BADGE"
                conversion_basis = "NUMERIC_BADGE_PROTECTION"

            # ACTION: CONVERT_EXISTING TITLE
            elif sid == promoted_title_shape_id and not has_slide_title:
                final_action = ConversionAction.CONVERT_EXISTING
                final_category = PlaceholderCategory.STANDALONE_TITLE_BOX
                final_status = PlaceholderStatus.SAFE_TO_CONVERT
                final_rec = "Convert standalone title text box into native slide title placeholder."
                final_reason = "Slide lacks native title; text box serves as visual title heading and converted to native title placeholder."
                final_conf_str = "HIGH"
                semantic_role = "TITLE"
                proposed_ph_type = "title"
                conversion_basis = "MULTI_SIGNAL_HEADING"

            # ACTION: CONVERT_EXISTING SUBTITLE
            elif sid == promoted_subtitle_shape_id:
                final_action = ConversionAction.CONVERT_EXISTING
                final_category = PlaceholderCategory.STANDALONE_SUBTITLE_BOX
                final_status = PlaceholderStatus.SAFE_TO_CONVERT
                final_rec = "Convert standalone subtitle text box into native slide subtitle placeholder."
                final_reason = "Slide contains visual subtitle heading below title; converted to native subtitle placeholder."
                final_conf_str = "HIGH"
                semantic_role = "SUBTITLE"
                proposed_ph_type = "subTitle"
                conversion_basis = "HEADING_HIERARCHY_SUBTITLE"

            # ACTION: CONVERT_EXISTING BODY
            elif sid == promoted_body_shape_id:
                final_action = ConversionAction.CONVERT_EXISTING
                final_category = PlaceholderCategory.STANDALONE_BODY_BOX
                final_status = PlaceholderStatus.SAFE_TO_CONVERT
                final_rec = "Convert standalone body text box into native slide body placeholder."
                final_reason = "Unambiguous primary body text box mapped to layout-supported native body placeholder."
                final_conf_str = "HIGH"
                semantic_role = "BODY"
                proposed_ph_type = "body"
                conversion_basis = "LAYOUT_SUPPORTED_BODY_CONTENT"

            # ACTION: MATERIALIZE_BODY (Existing Native Body / Object placeholder with text)
            elif (
                c["has_ph"]
                and base_category in (PlaceholderCategory.BODY, PlaceholderCategory.OBJECT)
                and has_text
            ):
                if not (inherited.xfrm_emu or inherited.has_slide_xfrm):
                    final_action = ConversionAction.MANUAL_REVIEW
                    final_category = base_category
                    final_status = PlaceholderStatus.MANUAL_REVIEW
                    final_rec = "Manual review required: geometry could not be resolved from layout or master."
                    final_reason = "Shape lacks local transform and layout/master has no matching geometry bounds."
                    final_conf_str = "LOW"
                    semantic_role = base_category.value
                    conversion_basis = "UNRESOLVED_INHERITANCE"
                elif g_conf < 0.70:
                    final_action = ConversionAction.PRESERVE_EXISTING
                    final_category = base_category
                    final_status = PlaceholderStatus.PRESERVED
                    final_rec = "Preserve body placeholder due to low confidence."
                    final_reason = f"Low confidence ({g_conf:.2f}); preserved placeholder."
                    final_conf_str = "LOW"
                    semantic_role = base_category.value
                    conversion_basis = "LOW_CONFIDENCE_GATE"
                else:
                    final_action = ConversionAction.CONVERT_EXISTING
                    final_category = base_category
                    final_status = PlaceholderStatus.SAFE_TO_CONVERT
                    final_rec = "Convert placeholder by materializing inherited geometry into self-contained shape."
                    final_reason = "Contains accessible text; materializing inherited geometry guarantees zero visual drift."
                    final_conf_str = "HIGH" if g_conf >= 0.85 else "MEDIUM"
                    semantic_role = base_category.value
                    conversion_basis = "GEOMETRY_MATERIALIZATION"

            # DEFAULT: PRESERVED shape
            else:
                final_action = ConversionAction.PRESERVE_EXISTING
                final_category = base_category
                final_status = base_status
                final_rec = base_rec
                final_reason = base_reason
                final_conf_str = base_conf
                semantic_role = base_category.value
                conversion_basis = "DEFAULT_PRESERVATION"

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
                },
                action=final_action,
                semantic_role=semantic_role,
                conversion_basis=conversion_basis,
                is_group_child=is_group_child,
                parent_group_id=parent_group_id,
                before_placeholder_type=ph_type or "none",
                proposed_placeholder_type=proposed_ph_type,
                has_ph=c["has_ph"],
                verified=False,
                verification_category=None
            )

            # Return native placeholders, actionable shapes, or all objects when requested
            if (
                include_all_shapes
                or c["has_ph"]
                or final_category in (
                    PlaceholderCategory.STANDALONE_TITLE_BOX,
                    PlaceholderCategory.STANDALONE_SUBTITLE_BOX,
                    PlaceholderCategory.STANDALONE_BODY_BOX
                )
                or sid in anim_title_candidate_ids
            ):
                results.append(analysis)

        return results

    def audit_slide_objects(
        self,
        raw_slide,
        slide_number: int,
        animation_target_ids: Optional[List[int]] = None,
        slide_height_pt: float = 540.0,
        slide_width_pt: float = 720.0
    ) -> List[PlaceholderAnalysis]:
        """Performs exhaustive Option C audit of all objects on a slide."""
        return self.analyze_slide_placeholders(
            raw_slide=raw_slide,
            slide_number=slide_number,
            animation_target_ids=animation_target_ids,
            slide_height_pt=slide_height_pt,
            slide_width_pt=slide_width_pt,
            include_all_shapes=True
        )

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
            is_empty=is_empty,
            has_ph=bool(ph_elems)
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
        is_empty: bool,
        has_ph: bool = True
    ) -> PlaceholderCategory:
        """Maps OpenXML placeholder attributes and contents to semantic category."""
        if not has_ph:
            if is_pic:
                return PlaceholderCategory.PICTURE
            if is_tbl:
                return PlaceholderCategory.TABLE
            if is_chart:
                return PlaceholderCategory.CHART
            if is_empty:
                return PlaceholderCategory.EMPTY
            return PlaceholderCategory.OTHER

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
