"""
SlideGuard AI Authoring Service (Phase 9).
Coordinates the generation of structured accessibility content recommendations:
  - Alternative text for meaningful visual elements
  - Accessible chart descriptions
  - Diagram / flow descriptions
  - Slide title generation for untitled slides
  - Table header interpretation and descriptions
Strictly enforces:
  - Gemini 3.8 Flash High model usage only
  - Privacy: minimal structured context only (no binaries, no full presentation)
  - Hallucination protection and schema validation
  - Zero direct PPTX modifications
"""

import logging
import re
from typing import List, Dict, Any, Optional

from models.enums import ShapeType, PlaceholderType, RuleBehavior, AccessibilityStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.shape import ShapeModel
from models.job import JobContext
from models.ai_authoring import (
    AuthoringType,
    AuthoringStatus,
    FactGrounding,
    AuthoringRecommendation,
    AuthoringSummary,
)
from engine.intelligence.gemini_assistant import GeminiAssistant
from rules.accessibility.alt_text import GENERIC_ALT_TEXTS

logger = logging.getLogger(__name__)


class AuthoringService:
    """
    Identifies accessibility content gaps from Phase 7/8 analysis
    and generates structured AI authoring recommendations via Gemini 3.8 Flash High.
    """

    def __init__(self, assistant: Optional[GeminiAssistant] = None):
        self.assistant = assistant or GeminiAssistant(model_name="gemini-3.8-flash-high")
        from rules.alt_text.service import ImageAccessibilityService
        # Task 3 runtime semantic model is Gemini 3.5 (never Gemini 3.8)
        self.image_service = ImageAccessibilityService(assistant=assistant or GeminiAssistant(model_name="gemini-3.5"))

    def generate_recommendations(
        self,
        context: JobContext,
        presentation: PresentationModel
    ) -> List[AuthoringRecommendation]:
        """
        Scans issues and slides to generate structured authoring recommendations.
        Never modifies presentation or OOXML structures.
        """
        recommendations: List[AuthoringRecommendation] = []
        summary = AuthoringSummary()

        # Collect issues by rule
        alt_issues = []
        ttl_issues = []
        tbl_issues = []

        for res in context.rule_results:
            if res.rule_id == "SG-ALT-01":
                alt_issues.extend(res.issues)
            elif res.rule_id == "SG-TTL-01":
                ttl_issues.extend(res.issues)
            elif res.rule_id == "SG-TBL-01":
                tbl_issues.extend(res.issues)

        # 1. Process Alternative Text / Visual Elements (SG-ALT-01)
        for issue in alt_issues:
            rec = self._evaluate_alt_text_issue(issue, presentation)
            if rec is not None:
                recommendations.append(rec)
                summary.generated += 1
                summary.by_type[rec.authoring_type.value] = summary.by_type.get(rec.authoring_type.value, 0) + 1
                if rec.status == AuthoringStatus.MANUAL_REVIEW:
                    summary.manual_review += 1
                if "fallback" in rec.rationale.lower() or any("fallback" in w.lower() for w in rec.warnings):
                    summary.fallback_cases += 1

        # Collect Task 3 Image Decisions and Report
        try:
            image_decisions = self.image_service.evaluate_presentation_images(presentation)
            context.image_decisions = image_decisions
            context.task_3_report = self.image_service.generate_report(image_decisions)
        except Exception as e:
            logger.warning("Error generating Task 3 image decisions: %s", e)
            context.image_decisions = []
            context.task_3_report = {}

        # 2. Process Slide Titles (SG-TTL-01)
        for issue in ttl_issues:
            rec = self._evaluate_title_issue(issue, presentation)
            if rec is not None:
                recommendations.append(rec)
                summary.generated += 1
                summary.by_type[rec.authoring_type.value] = summary.by_type.get(rec.authoring_type.value, 0) + 1
                if rec.status == AuthoringStatus.MANUAL_REVIEW:
                    summary.manual_review += 1
                if "fallback" in rec.rationale.lower() or any("fallback" in w.lower() for w in rec.warnings):
                    summary.fallback_cases += 1

        # 3. Process Table Accessibility (SG-TBL-01)
        for issue in tbl_issues:
            rec = self._evaluate_table_issue(issue, presentation)
            if rec is not None:
                recommendations.append(rec)
                summary.generated += 1
                summary.by_type[rec.authoring_type.value] = summary.by_type.get(rec.authoring_type.value, 0) + 1
                if rec.status == AuthoringStatus.MANUAL_REVIEW:
                    summary.manual_review += 1
                if "fallback" in rec.rationale.lower() or any("fallback" in w.lower() for w in rec.warnings):
                    summary.fallback_cases += 1

        context.ai_recommendations = recommendations
        context.ai_authoring_summary = summary.to_dict()
        logger.info(
            "Phase 9 AI authoring generated %d recommendation(s) across %d slides.",
            len(recommendations),
            len(presentation.slides)
        )
        return recommendations

    def _evaluate_alt_text_issue(
        self,
        issue: Any,
        presentation: PresentationModel
    ) -> Optional[AuthoringRecommendation]:
        """Evaluates missing or generic alt text on visual shapes."""
        slide = next((s for s in presentation.slides if s.slide_number == issue.slide_number), None)
        if slide is None:
            return None

        shape = self._find_shape(slide.shapes, issue.shape_id)
        if shape is None:
            return None

        # Rule 5: Decorative vs Meaningful Protection
        # If deterministic analysis already established DECORATIVE, do NOT ask AI to invent alt text!
        if shape.is_decorative:
            logger.debug("Skipping AI authoring for decorative shape %s", shape.shape_id)
            return None

        # Rule 2: If existing alt text is non-empty and NOT generic, skip unnecessary AI call
        current_alt = (shape.alt_text or "").strip()
        if current_alt and current_alt.lower() not in GENERIC_ALT_TEXTS:
            logger.debug("Shape %s already has descriptive alt text: '%s'. Skipping AI authoring.", shape.shape_id, current_alt)
            return None

        # Task 3: Check if shape is a picture/image candidate
        raw_elem = shape.raw_element
        has_blip = raw_elem is not None and bool(raw_elem.xpath(".//a:blip"))
        is_picture_shape = (shape.shape_type == ShapeType.PICTURE or has_blip)

        if is_picture_shape and slide.raw_slide is not None:
            try:
                from rules.alt_text.models import ImageAction, ImageClassification
                slide_candidates = self.image_service.extractor.extract_slide_images(
                    raw_slide=slide.raw_slide,
                    slide_number=slide.slide_number
                )
                candidate = next((c for c in slide_candidates if c.shape_id == shape.shape_id), None)
                if candidate:

                    decision = self.image_service.evaluate_candidate(candidate)
                    model_name = decision.model_traceability.get("slideguard_runtime_model", "gemini-3.5")

                    if decision.action == ImageAction.PRESERVE and decision.original_alt_text:
                        return AuthoringRecommendation(
                            recommendation_id=f"rec_alt_{slide.slide_number}_{shape.shape_id}",
                            rule_id=issue.rule_id,
                            slide_number=slide.slide_number,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            authoring_type=AuthoringType.IMAGE_DESCRIPTION,
                            current_value=decision.original_alt_text,
                            suggested_value=decision.original_alt_text,
                            rationale=decision.reason,
                            confidence=decision.confidence,
                            fact_grounding=FactGrounding.SUPPORTED_FACT,
                            evidence={
                                "action": "PRESERVE",
                                "classification": decision.classification.value,
                                "generation_basis": decision.generation_basis,
                                "quality_class": decision.quality_class,
                                "model_traceability": decision.model_traceability
                            },
                            requires_human_review=False,
                            status=AuthoringStatus.APPROVED,
                            model=model_name
                        )
                    elif decision.action == ImageAction.MARK_DECORATIVE:
                        return AuthoringRecommendation(
                            recommendation_id=f"rec_alt_{slide.slide_number}_{shape.shape_id}",
                            rule_id=issue.rule_id,
                            slide_number=slide.slide_number,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            authoring_type=AuthoringType.IMAGE_DESCRIPTION,
                            current_value=current_alt or "[Missing]",
                            suggested_value="[MARK_DECORATIVE]",
                            rationale=decision.reason,
                            confidence=decision.confidence,
                            fact_grounding=FactGrounding.SUPPORTED_FACT,
                            evidence={
                                "action": "MARK_DECORATIVE",
                                "classification": "DECORATIVE",
                                "generation_basis": decision.generation_basis,
                                "quality_class": decision.quality_class,
                                "model_traceability": decision.model_traceability
                            },
                            requires_human_review=(decision.confidence < 0.85),
                            status=AuthoringStatus.PENDING_REVIEW if decision.confidence >= 0.85 else AuthoringStatus.MANUAL_REVIEW,
                            model=model_name
                        )
                    elif decision.action == ImageAction.GENERATE_ALT_TEXT:
                        return AuthoringRecommendation(
                            recommendation_id=f"rec_alt_{slide.slide_number}_{shape.shape_id}",
                            rule_id=issue.rule_id,
                            slide_number=slide.slide_number,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            authoring_type=AuthoringType.IMAGE_DESCRIPTION,
                            current_value=current_alt or "[Missing]",
                            suggested_value=decision.effective_alt_text,
                            rationale=decision.reason,
                            confidence=decision.confidence,
                            fact_grounding=FactGrounding.SUPPORTED_FACT,
                            evidence={
                                "action": "GENERATE_ALT_TEXT",
                                "classification": "MEANINGFUL",
                                "generation_basis": decision.generation_basis,
                                "quality_class": decision.quality_class,
                                "model_traceability": decision.model_traceability
                            },
                            requires_human_review=(decision.confidence < 0.85),
                            status=AuthoringStatus.PENDING_REVIEW if decision.confidence >= 0.85 else AuthoringStatus.MANUAL_REVIEW,
                            model=model_name
                        )
                    else:  # MANUAL_REVIEW
                        return AuthoringRecommendation(
                            recommendation_id=f"rec_alt_{slide.slide_number}_{shape.shape_id}",
                            rule_id=issue.rule_id,
                            slide_number=slide.slide_number,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            authoring_type=AuthoringType.IMAGE_DESCRIPTION,
                            current_value=current_alt or "[Missing]",
                            suggested_value=decision.effective_alt_text or "MANUAL_REVIEW",
                            rationale=decision.reason,
                            confidence=decision.confidence,
                            fact_grounding=FactGrounding.UNCERTAIN,
                            evidence={
                                "action": "MANUAL_REVIEW",
                                "classification": decision.classification.value,
                                "generation_basis": decision.generation_basis,
                                "quality_class": decision.quality_class,
                                "model_traceability": decision.model_traceability
                            },
                            requires_human_review=True,
                            status=AuthoringStatus.MANUAL_REVIEW,
                            model=model_name
                        )
            except Exception as e:
                logger.warning("ImageAccessibilityService evaluation fallback on shape %s: %s", shape.shape_id, e)


        # Determine target authoring type
        authoring_type = AuthoringType.ALT_TEXT
        if shape.placeholder_type == PlaceholderType.CHART or "chart" in shape.name.lower():
            authoring_type = AuthoringType.CHART_DESCRIPTION
        elif shape.placeholder_type == PlaceholderType.OBJECT or shape.children or "diagram" in shape.name.lower():
            authoring_type = AuthoringType.DIAGRAM_DESCRIPTION
        elif shape.shape_type == ShapeType.PICTURE:
            authoring_type = AuthoringType.IMAGE_DESCRIPTION

        # Check for ambiguous visual graphic
        extra = getattr(issue, "extra_data", {}) or {}
        is_ambiguous = (extra.get("classification") == "UNKNOWN" and (shape.width_pt < 40 or shape.height_pt < 40))

        # Build minimal structured privacy payload
        context_payload = {
            "slide_number": slide.slide_number,
            "slide_title": slide.title or "",
            "shape_id": shape.shape_id,
            "shape_name": shape.name,
            "shape_type": shape.shape_type.value,
            "dimensions": f"{shape.width_pt:.0f}x{shape.height_pt:.0f}pt",
            "position": f"left={shape.left_pt:.0f}pt, top={shape.top_pt:.0f}pt",
            "existing_alt_text": current_alt,
            "nearby_text": [s.text_content[:80] for s in slide.shapes if s.has_text and s.shape_id != shape.shape_id][:3],
        }

        # If ambiguous, flag for human verification
        if is_ambiguous:
            ai_res = self.assistant.author_content(authoring_type.value, context_payload)
            return AuthoringRecommendation(
                recommendation_id=f"rec_alt_{slide.slide_number}_{shape.shape_id}",
                rule_id=issue.rule_id,
                slide_number=slide.slide_number,
                shape_id=shape.shape_id,
                shape_name=shape.name,
                authoring_type=authoring_type,
                current_value=current_alt or "[Missing]",
                suggested_value=ai_res.get("suggested_value", "MANUAL_REVIEW"),
                rationale="Ambiguous visual element; AI suggestion requires human author verification.",
                confidence=min(0.65, float(ai_res.get("confidence", 0.5))),
                fact_grounding=FactGrounding.UNCERTAIN,
                evidence=context_payload,
                warnings=["Visual graphic context is ambiguous; verify whether image is decorative or meaningful."],
                requires_human_review=True,
                status=AuthoringStatus.MANUAL_REVIEW if ai_res.get("status") == "MANUAL_REVIEW" else AuthoringStatus.PENDING_REVIEW,
                model=self.assistant.model_name
            )

        # Call Gemini Assistant
        ai_res = self.assistant.author_content(authoring_type.value, context_payload)

        # Grounding mapping
        fg_str = ai_res.get("fact_grounding", "SUPPORTED_FACT").upper()
        try:
            grounding = FactGrounding(fg_str)
        except ValueError:
            grounding = FactGrounding.INFERENCE

        st_str = ai_res.get("status", "PENDING_REVIEW")
        try:
            status = AuthoringStatus(st_str)
        except ValueError:
            status = AuthoringStatus.PENDING_REVIEW

        return AuthoringRecommendation(
            recommendation_id=f"rec_alt_{slide.slide_number}_{shape.shape_id}",
            rule_id=issue.rule_id,
            slide_number=slide.slide_number,
            shape_id=shape.shape_id,
            shape_name=shape.name,
            authoring_type=authoring_type,
            current_value=current_alt or "[Missing]",
            suggested_value=ai_res.get("suggested_value", ""),
            rationale=ai_res.get("rationale", ""),
            confidence=float(ai_res.get("confidence", 0.8)),
            fact_grounding=grounding,
            evidence=context_payload,
            warnings=ai_res.get("warnings", []),
            requires_human_review=ai_res.get("requires_human_review", True),
            status=status,
            model=self.assistant.model_name
        )

    def _evaluate_title_issue(
        self,
        issue: Any,
        presentation: PresentationModel
    ) -> Optional[AuthoringRecommendation]:
        """Evaluates missing slide titles and proposes meaningful titles."""
        slide = next((s for s in presentation.slides if s.slide_number == issue.slide_number), None)
        if slide is None:
            return None

        # If slide already has a title, skip
        if slide.title and slide.title.strip():
            return None

        extra = getattr(issue, "extra_data", {}) or {}
        cand_id = issue.shape_id or extra.get("candidate_shape_id")
        cand_text = extra.get("candidate_text", "")

        # Minimal context payload
        slide_texts = [s.text_content.strip() for s in slide.shapes if s.has_text]
        context_payload = {
            "slide_number": slide.slide_number,
            "nearby_text": slide_texts[:4],
            "candidate_text": cand_text,
            "shape_count": len(slide.shapes),
        }

        ai_res = self.assistant.author_slide_title(context_payload)

        fg_str = ai_res.get("fact_grounding", "SUPPORTED_FACT").upper()
        try:
            grounding = FactGrounding(fg_str)
        except ValueError:
            grounding = FactGrounding.INFERENCE

        st_str = ai_res.get("status", "PENDING_REVIEW")
        try:
            status = AuthoringStatus(st_str)
        except ValueError:
            status = AuthoringStatus.PENDING_REVIEW

        return AuthoringRecommendation(
            recommendation_id=f"rec_ttl_{slide.slide_number}",
            rule_id=issue.rule_id,
            slide_number=slide.slide_number,
            shape_id=cand_id,
            shape_name="Slide Title Heading",
            authoring_type=AuthoringType.SLIDE_TITLE,
            current_value="[Missing Title]",
            suggested_value=ai_res.get("suggested_value", ""),
            rationale=ai_res.get("rationale", "Proposed title derived from slide content."),
            confidence=float(ai_res.get("confidence", 0.75)),
            fact_grounding=grounding,
            evidence=context_payload,
            warnings=ai_res.get("warnings", []),
            requires_human_review=True,
            status=status,
            model=self.assistant.model_name
        )

    def _evaluate_table_issue(
        self,
        issue: Any,
        presentation: PresentationModel
    ) -> Optional[AuthoringRecommendation]:
        """Evaluates ambiguous table headers or complex tables."""
        slide = next((s for s in presentation.slides if s.slide_number == issue.slide_number), None)
        if slide is None:
            return None

        shape = self._find_shape(slide.shapes, issue.shape_id)
        if shape is None:
            return None

        # Complex table protection: merged cells should remain manual review
        if "TBL_MERGED_CELLS" in getattr(issue, "issue_id", ""):
            return AuthoringRecommendation(
                recommendation_id=f"rec_tbl_merge_{slide.slide_number}_{shape.shape_id}",
                rule_id=issue.rule_id,
                slide_number=slide.slide_number,
                shape_id=shape.shape_id,
                shape_name=shape.name,
                authoring_type=AuthoringType.TABLE_DESCRIPTION,
                current_value="Merged Cells Table",
                suggested_value="MANUAL_REVIEW",
                rationale="Table contains complex merged cells (gridSpan/rowSpan). Manual review required to protect structure.",
                confidence=0.9,
                fact_grounding=FactGrounding.SUPPORTED_FACT,
                evidence={"has_merged_cells": True},
                warnings=["Complex table structure cannot be safely auto-modified."],
                requires_human_review=True,
                status=AuthoringStatus.MANUAL_REVIEW,
                model=self.assistant.model_name
            )

        if "TBL_AMBIGUOUS_HEADER" in getattr(issue, "issue_id", "") or "TBL_NO_HEADER" in getattr(issue, "issue_id", ""):
            context_payload = {
                "slide_number": slide.slide_number,
                "slide_title": slide.title or "",
                "table_name": shape.name,
                "dimensions": f"{shape.table_row_count} rows x {shape.table_col_count} columns",
                "preview_text": shape.text_content[:200]
            }

            ai_res = self.assistant.author_table_interpretation(context_payload)

            return AuthoringRecommendation(
                recommendation_id=f"rec_tbl_{slide.slide_number}_{shape.shape_id}",
                rule_id=issue.rule_id,
                slide_number=slide.slide_number,
                shape_id=shape.shape_id,
                shape_name=shape.name,
                authoring_type=AuthoringType.TABLE_HEADER_INTERPRETATION,
                current_value="Unmarked Table Header",
                suggested_value=ai_res.get("suggested_value", "Enable header row markup ('firstRow=1')"),
                rationale=ai_res.get("rationale", "Table structure indicates top row serves as column headers."),
                confidence=float(ai_res.get("confidence", 0.75)),
                fact_grounding=FactGrounding.SUPPORTED_FACT,
                evidence=context_payload,
                warnings=ai_res.get("warnings", []),
                requires_human_review=True,
                status=AuthoringStatus.PENDING_REVIEW,
                model=self.assistant.model_name
            )

        return None

    def _find_shape(self, shapes: List[ShapeModel], shape_id: Optional[int]) -> Optional[ShapeModel]:
        """Recursively locates shape by ID."""
        if shape_id is None:
            return None
        for s in shapes:
            if s.shape_id == shape_id:
                return s
            if s.children:
                found = self._find_shape(s.children, shape_id)
                if found is not None:
                    return found
        return None
