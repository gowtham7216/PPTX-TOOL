"""
SlideGuard Image Accessibility Service (Task 3).
Coordinates deterministic image candidate extraction, Gemini 3.5 semantic analysis,
confidence gating, and safety gate enforcement for alternative text and decorative images.
"""

import logging
from typing import Dict, Any, List, Optional, Set, Tuple

from engine.intelligence.gemini_assistant import GeminiAssistant, DEFAULT_RUNTIME_MODEL
from rules.alt_text.models import (
    ImageCandidate,
    ImageClassification,
    ImageAction,
    ImageDecision,
    GenerationBasis,
    QualityClass,
)
from rules.alt_text.extractor import ImageExtractor, is_weak_or_generic_alt_text
from rules.alt_text.quality import is_generic_title_repetition, classify_alt_text_quality

logger = logging.getLogger(__name__)


class ImageAccessibilityService:
    """
    Coordinates semantic alternative text generation and decorative image detection.
    Strictly enforces:
      - Runtime Model: Gemini 3.5 (traceable in decision metadata)
      - Existing Alt Text Protection: never overwrites author-provided descriptive alt text
      - Decorative Image Protection: position/size alone never determines decorative status
      - Confidence Gating: High (>= 0.85), Medium (0.70-0.849), Low (< 0.70)
      - Animation & Group Protection: animated images preserved / flagged for manual review
    """

    def __init__(self, assistant: Optional[GeminiAssistant] = None):
        self.assistant = assistant or GeminiAssistant(model_name=DEFAULT_RUNTIME_MODEL)
        self.extractor = ImageExtractor()

    def evaluate_slide_images(
        self,
        raw_slide: Any,
        slide_number: int,
        animation_target_ids: Optional[Set[int]] = None
    ) -> List[ImageDecision]:
        """Evaluates all image candidates on a single slide."""
        candidates = self.extractor.extract_slide_images(
            raw_slide=raw_slide,
            slide_number=slide_number,
            animation_target_ids=animation_target_ids
        )
        return [self.evaluate_candidate(c) for c in candidates]

    def evaluate_presentation_images(
        self,
        presentation: Any,
        animation_target_ids_by_slide: Optional[Dict[int, Set[int]]] = None
    ) -> List[ImageDecision]:
        """Evaluates all image candidates across a presentation with cross-slide repetition analysis."""
        candidates = self.extractor.extract_presentation_images(
            presentation=presentation,
            animation_target_ids_by_slide=animation_target_ids_by_slide
        )
        return [self.evaluate_candidate(c) for c in candidates]

    def evaluate_candidate(self, candidate: ImageCandidate) -> ImageDecision:
        """Evaluates an individual image candidate through deterministic gates and Gemini 3.5."""
        sid = candidate.shape_id
        s_num = candidate.slide_number
        s_name = candidate.shape_name
        orig_alt = candidate.existing_alt_text
        is_anim = candidate.is_animated
        is_generic = candidate.is_generic_alt_text
        dims = candidate.dimensions_pt
        w_pt, h_pt = dims

        # =====================================================================
        # 1. HARD DETERMINISTIC SAFETY GATES (Pre-Gemini)
        # =====================================================================

        # GATE 1: Animation Target Protection
        if is_anim:
            logger.info("Slide %d Shape %d is targeted by animation; preserved for manual review.", s_num, sid)
            return ImageDecision(
                slide_number=s_num,
                shape_id=sid,
                shape_name=s_name,
                classification=ImageClassification.MEANINGFUL,
                action=ImageAction.MANUAL_REVIEW,
                alt_text=orig_alt,
                original_alt_text=orig_alt,
                confidence=0.85,
                reason="Image is targeted by animation sequence; conversion requires timing review.",
                is_animated=True,
                is_grouped=bool(candidate.containing_group_id),
                generation_basis="STRUCTURAL_CONTEXT",
                quality_class="MANUAL_REVIEW",
                status="MANUAL_REVIEW",
                model_traceability=self._build_traceability()
            )

        # GATE 2: Existing Meaningful Alt Text Protection
        if orig_alt and not is_generic:
            logger.info("Slide %d Shape %d already has descriptive alt text; preserving in-place.", s_num, sid)
            return ImageDecision(
                slide_number=s_num,
                shape_id=sid,
                shape_name=s_name,
                classification=ImageClassification.MEANINGFUL,
                action=ImageAction.PRESERVE,
                alt_text=orig_alt,
                original_alt_text=orig_alt,
                confidence=0.98,
                reason=f"Existing author-provided alt text preserved: '{orig_alt[:40]}...'",
                is_animated=False,
                is_grouped=bool(candidate.containing_group_id),
                generation_basis="EXISTING_AUTHOR_TEXT",
                quality_class="SPECIFIC",
                status="APPROVED",
                model_traceability=self._build_traceability()
            )

        # GATE 3: Explicit Author Decorative Tag
        if candidate.is_already_decorative:
            logger.info("Slide %d Shape %d is already marked decorative; preserving decorative state.", s_num, sid)
            return ImageDecision(
                slide_number=s_num,
                shape_id=sid,
                shape_name=s_name,
                classification=ImageClassification.DECORATIVE,
                action=ImageAction.MARK_DECORATIVE,
                alt_text="",
                original_alt_text=orig_alt,
                confidence=0.98,
                reason="Element explicitly designated as decorative by author.",
                is_animated=False,
                is_grouped=bool(candidate.containing_group_id),
                is_decorative_tagged=True,
                generation_basis="DETERMINISTIC",
                quality_class="SPECIFIC",
                status="APPROVED",
                model_traceability=self._build_traceability()
            )

        # GATE 4: Deterministic Separator Line / Divider Image
        if (w_pt <= 5.0 or h_pt <= 5.0) and not candidate.has_hyperlink and not candidate.nearby_caption:
            logger.info("Slide %d Shape %d is a geometric divider line; classified as DECORATIVE.", s_num, sid)
            return ImageDecision(
                slide_number=s_num,
                shape_id=sid,
                shape_name=s_name,
                classification=ImageClassification.DECORATIVE,
                action=ImageAction.MARK_DECORATIVE,
                alt_text="",
                original_alt_text=orig_alt,
                confidence=0.95,
                reason="Geometric separator line serving purely visual aesthetic role.",
                is_animated=False,
                is_grouped=bool(candidate.containing_group_id),
                is_decorative_tagged=True,
                generation_basis="DETERMINISTIC",
                quality_class="SPECIFIC",
                status="APPROVED",
                model_traceability=self._build_traceability()
            )

        # GATE 5: Tiny Decorative Icon / Bullet
        if (w_pt <= 20.0 and h_pt <= 20.0) and not candidate.has_hyperlink and not candidate.nearby_caption:
            logger.info("Slide %d Shape %d is a tiny icon bullet; classified as DECORATIVE.", s_num, sid)
            return ImageDecision(
                slide_number=s_num,
                shape_id=sid,
                shape_name=s_name,
                classification=ImageClassification.DECORATIVE,
                action=ImageAction.MARK_DECORATIVE,
                alt_text="",
                original_alt_text=orig_alt,
                confidence=0.92,
                reason="Small aesthetic icon / bullet marker conveying no standalone information.",
                is_animated=False,
                is_grouped=bool(candidate.containing_group_id),
                is_decorative_tagged=True,
                generation_basis="DETERMINISTIC",
                quality_class="SPECIFIC",
                status="APPROVED",
                model_traceability=self._build_traceability()
            )

        # GATE 6: Repeated Background Pattern Across Multiple Slides
        if candidate.is_repeated_across_slides and not candidate.nearby_caption and not candidate.has_hyperlink:
            logger.info("Slide %d Shape %d is repeated across >=3 slides; classified as DECORATIVE.", s_num, sid)
            return ImageDecision(
                slide_number=s_num,
                shape_id=sid,
                shape_name=s_name,
                classification=ImageClassification.DECORATIVE,
                action=ImageAction.MARK_DECORATIVE,
                alt_text="",
                original_alt_text=orig_alt,
                confidence=0.90,
                reason="Repeated background graphic / ornamentation across presentation.",
                is_animated=False,
                is_grouped=bool(candidate.containing_group_id),
                is_decorative_tagged=True,
                generation_basis="DETERMINISTIC",
                quality_class="SPECIFIC",
                status="APPROVED",
                model_traceability=self._build_traceability()
            )

        # =====================================================================
        # 2. GEMINI 3.5 RUNTIME SEMANTIC DECISION
        # =====================================================================
        context_payload = candidate.to_context_dict()
        ai_res = self.assistant.analyze_image_accessibility(context_payload)

        raw_class = ai_res.get("classification", "UNKNOWN").upper()
        raw_act = ai_res.get("action", "MANUAL_REVIEW").upper()
        ai_alt = str(ai_res.get("alt_text", "")).strip()
        ai_conf = float(ai_res.get("confidence", 0.70))
        ai_reason = str(ai_res.get("reason", "")).strip()

        # Map to enums safely
        try:
            classification = ImageClassification(raw_class)
        except ValueError:
            classification = ImageClassification.UNKNOWN

        try:
            action = ImageAction(raw_act)
        except ValueError:
            action = ImageAction.MANUAL_REVIEW

        # =====================================================================
        # 3. DETERMINISTIC SAFETY GATE & CONFIDENCE GATING (Post-Gemini)
        # =====================================================================

        # Quality & basis tracking
        has_visual_analysis = False  # SlideGuard context-based runtime does not inspect pixels directly
        generation_basis = GenerationBasis.STRUCTURAL_CONTEXT.value

        # HARD SAFETY GATE 1: Existing good alt text cannot be overwritten by Gemini
        if orig_alt and not is_generic:
            action = ImageAction.PRESERVE
            classification = ImageClassification.MEANINGFUL
            ai_alt = orig_alt
            ai_conf = 0.98
            ai_reason = "Existing author alt text protected by safety gate."
            generation_basis = GenerationBasis.EXISTING_AUTHOR_TEXT.value
            quality_class = QualityClass.SPECIFIC.value
        else:
            # Check for generic title repetition (ALT_TEXT_TOO_GENERIC)
            is_generic_title = is_generic_title_repetition(ai_alt, candidate.nearby_title)
            if is_generic_title:
                quality_class = QualityClass.GENERIC.value
                # Per Rule 6 & 7: Downgrade generic title-only descriptions to MANUAL_REVIEW
                action = ImageAction.MANUAL_REVIEW
                ai_conf = min(ai_conf, 0.75)
                ai_reason = f"Alt text is too generic (ALT_TEXT_TOO_GENERIC: merely repeats slide title '{candidate.nearby_title}'). Visual pixels not analyzed; queued for manual review."
            elif classification == ImageClassification.DECORATIVE:
                quality_class = QualityClass.SPECIFIC.value
                generation_basis = GenerationBasis.DETERMINISTIC.value
            elif ai_alt:
                # Contextual description based on slide context
                quality_class = QualityClass.CONTEXTUAL.value
            else:
                quality_class = QualityClass.MANUAL_REVIEW.value

        # HARD SAFETY GATE 2: Confidence Gating
        # HIGH: >= 0.85 -> action permitted
        # MEDIUM: 0.70-0.849 -> MANUAL_REVIEW
        # LOW: < 0.70 -> MANUAL_REVIEW
        if ai_conf < 0.70:
            action = ImageAction.MANUAL_REVIEW
            if not ai_reason:
                ai_reason = "Low AI confidence (<0.70); requires human verification."
        elif 0.70 <= ai_conf < 0.85:
            action = ImageAction.MANUAL_REVIEW
            if not ai_reason:
                ai_reason = "Medium AI confidence (0.70-0.849); queued for manual review."

        # HARD SAFETY GATE 3: If GENERATE_ALT_TEXT has empty alt text, force MANUAL_REVIEW
        if action == ImageAction.GENERATE_ALT_TEXT and not ai_alt:
            action = ImageAction.MANUAL_REVIEW
            classification = ImageClassification.UNKNOWN
            ai_conf = min(ai_conf, 0.65)
            quality_class = QualityClass.MANUAL_REVIEW.value
            ai_reason = "Proposed alt text was empty; manual review required."

        # HARD SAFETY GATE 4: If DECORATIVE, alt_text must be empty
        if classification == ImageClassification.DECORATIVE:
            ai_alt = ""
            if action not in (ImageAction.MANUAL_REVIEW, ImageAction.PRESERVE):
                action = ImageAction.MARK_DECORATIVE

        # Determine review status
        status = "PENDING_REVIEW"
        if action == ImageAction.MANUAL_REVIEW or ai_conf < 0.85:
            status = "MANUAL_REVIEW"
        elif action in (ImageAction.PRESERVE, ImageAction.MARK_DECORATIVE, ImageAction.GENERATE_ALT_TEXT) and ai_conf >= 0.85:
            status = "PENDING_REVIEW"

        return ImageDecision(
            slide_number=s_num,
            shape_id=sid,
            shape_name=s_name,
            classification=classification,
            action=action,
            alt_text=ai_alt,
            original_alt_text=orig_alt,
            confidence=ai_conf,
            reason=ai_reason,
            is_animated=False,
            is_grouped=bool(candidate.containing_group_id),
            is_decorative_tagged=(action == ImageAction.MARK_DECORATIVE),
            generation_basis=generation_basis,
            quality_class=quality_class,
            status=status,
            model_traceability=self._build_traceability()
        )

    def _build_traceability(self) -> Dict[str, str]:
        return {
            "antigravity_coding_model": "Gemini 3.8 Flash High",
            "slideguard_runtime_model": "Gemini 3.5",
            "runtime_model_identifier": self.assistant.model_name
        }

    def generate_report(
        self,
        decisions: List[ImageDecision],
        actions_applied: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """Generates structured Task 3 report dictionary."""
        from rules.alt_text.reporter import generate_task_3_report
        return generate_task_3_report(decisions, actions_applied=actions_applied)

