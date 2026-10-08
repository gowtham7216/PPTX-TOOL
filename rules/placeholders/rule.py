"""
SlideGuard Placeholder Conversion Rule.
Evaluates presentation slides for placeholder alignment and converts safe placeholders:
1. Detects standalone title text boxes that lack native placeholder semantics (SG-PH-01).
2. Identifies body/object placeholders for safe structural conversion and geometry materialization.
3. Preserves native title, footer, slide-number, date-time, empty, and specialized placeholders.
"""

from typing import List, Optional, Dict, Any
import logging

from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction
from rules.placeholders.classifier import (
    PlaceholderClassifier,
    PlaceholderAnalysis,
    PlaceholderCategory,
    PlaceholderStatus,
    _safe_xpath,
)
from rules.placeholders.transformer import PlaceholderTransformer, ConversionResult

logger = logging.getLogger(__name__)


class PlaceholderRule(BaseAccessibilityRule):
    """
    Enforces native placeholder semantics and executes safe in-place placeholder conversions.
    WCAG 2.1 - 1.3.1 Info and Relationships (Level A).
    """

    def __init__(
        self,
        classifier: Optional[PlaceholderClassifier] = None,
        transformer: Optional[PlaceholderTransformer] = None,
        assistant: Optional[Any] = None
    ):
        from engine.intelligence.gemini_assistant import GeminiAssistant, resolve_runtime_model
        resolved_assistant = assistant or GeminiAssistant(model_name=resolve_runtime_model("gemini-3.5"))
        self.classifier = classifier or PlaceholderClassifier(assistant=resolved_assistant)
        self.transformer = transformer or PlaceholderTransformer(classifier=self.classifier)

    @property
    def rule_id(self) -> str:
        return "SG-PH-01"

    @property
    def name(self) -> str:
        return "Native Placeholder Usage"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.1 Info and Relationships (Level A)"

    @property
    def description(self) -> str:
        return (
            "Ensures slide headers and content use native slide placeholders instead of loose text boxes, "
            "and safely materializes placeholder structures without altering visual layout."
        )

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []
        raw_slide = slide.raw_slide
        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return issues

        anim_targets = self._get_slide_anim_targets(slide)
        slide_h_pt = presentation.slide_height_pt if presentation else 540.0
        slide_w_pt = presentation.slide_width_pt if presentation else 720.0

        analyses = self.classifier.analyze_slide_placeholders(
            raw_slide=raw_slide,
            slide_number=slide.slide_number,
            animation_target_ids=anim_targets,
            slide_height_pt=slide_h_pt,
            slide_width_pt=slide_w_pt
        )

        for analysis in analyses:
            extra_data = analysis.to_dict()

            # 1. Standalone text box acting as slide title
            if analysis.category == PlaceholderCategory.STANDALONE_TITLE_BOX:
                issues.append(
                    Issue(
                        issue_id=f"PH_MISSING_TITLE_{slide.slide_number}_{analysis.shape_id}",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=slide.slide_number,
                        severity=IssueSeverity.WARNING,
                        shape_id=analysis.shape_id,
                        shape_name=analysis.shape_name,
                        wcag_mapping=self.wcag_mapping,
                        description=(
                            f"Slide {slide.slide_number}: Standalone text box '{analysis.text_preview}' "
                            f"appears to be a slide title but is not a native slide title placeholder."
                        ),
                        recommendation="Convert standalone title text box into a native slide title placeholder.",
                        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
                        extra_data=extra_data
                    )
                )

            # 2. Body / Object placeholder requiring conversion / geometry materialization
            elif analysis.category in (PlaceholderCategory.BODY, PlaceholderCategory.OBJECT) and analysis.has_text:
                if analysis.is_safe_to_convert:
                    issues.append(
                        Issue(
                            issue_id=f"PH_CONVERT_{slide.slide_number}_{analysis.shape_id}",
                            rule_id=self.rule_id,
                            rule_name=self.name,
                            slide_number=slide.slide_number,
                            severity=IssueSeverity.INFO,
                            shape_id=analysis.shape_id,
                            shape_name=analysis.shape_name,
                            wcag_mapping=self.wcag_mapping,
                            description=(
                                f"Slide {slide.slide_number}: Placeholder '{analysis.shape_name}' (ID {analysis.shape_id}) "
                                f"contains accessible content. Materializing geometry guarantees design stability."
                            ),
                            recommendation=analysis.recommendation,
                            behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
                            extra_data=extra_data
                        )
                    )
                else:
                    # Manual review
                    issues.append(
                        Issue(
                            issue_id=f"PH_MANUAL_{slide.slide_number}_{analysis.shape_id}",
                            rule_id=self.rule_id,
                            rule_name=self.name,
                            slide_number=slide.slide_number,
                            severity=IssueSeverity.WARNING,
                            shape_id=analysis.shape_id,
                            shape_name=analysis.shape_name,
                            wcag_mapping=self.wcag_mapping,
                            description=f"Slide {slide.slide_number}: Placeholder '{analysis.shape_name}' requires manual review: {analysis.reason}",
                            recommendation=analysis.recommendation,
                            behavior=RuleBehavior.MANUAL_REVIEW,
                            extra_data=extra_data
                        )
                    )

            # 3. Preserved placeholders with special manual review note (e.g. animated)
            elif analysis.status == PlaceholderStatus.MANUAL_REVIEW:
                issues.append(
                    Issue(
                        issue_id=f"PH_MANUAL_{slide.slide_number}_{analysis.shape_id}",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=slide.slide_number,
                        severity=IssueSeverity.WARNING,
                        shape_id=analysis.shape_id,
                        shape_name=analysis.shape_name,
                        wcag_mapping=self.wcag_mapping,
                        description=f"Slide {slide.slide_number}: Placeholder '{analysis.shape_name}' requires manual review: {analysis.reason}",
                        recommendation=analysis.recommendation,
                        behavior=RuleBehavior.MANUAL_REVIEW,
                        extra_data=extra_data
                    )
                )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """
        Executes physical in-place placeholder conversion for an approved issue.
        """
        shape_id = issue.shape_id or issue.extra_data.get("shape_id")
        action_id = f"act_ph_{slide.slide_number}_{shape_id}"

        raw_slide = slide.raw_slide
        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Cannot convert Shape {shape_id}: Missing raw slide element.",
                status=RemediationStatus.FAILED,
                error_message="Raw slide element is None."
            )

        status_str = issue.extra_data.get("status", "")
        if status_str == PlaceholderStatus.MANUAL_REVIEW.value:
            reason = issue.extra_data.get("reason", "Placeholder flagged for manual review.")
            logger.info("Skipping Shape %d on Slide %d: %s", shape_id, slide.slide_number, reason)
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Preserved Shape {shape_id}: {reason}",
                status=RemediationStatus.SKIPPED,
                error_message=reason
            )

        slide_elem = raw_slide._element
        anim_targets = self._get_slide_anim_targets(slide)

        result: ConversionResult = self.transformer.convert_slide(
            slide_elem=slide_elem,
            slide_number=slide.slide_number,
            raw_slide=raw_slide,
            animation_target_ids=anim_targets,
            approved_shape_ids=[shape_id] if shape_id else None
        )

        if not result.success:
            err = "; ".join(result.errors)
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Placeholder conversion failed on Shape {shape_id}.",
                status=RemediationStatus.FAILED,
                error_message=err
            )

        if result.placeholders_converted > 0:
            status = RemediationStatus.APPLIED
            desc = (
                f"Successfully converted Shape {shape_id} on Slide {slide.slide_number}. "
                f"Preserved exact coordinates, IDs, formatting, and relationships."
            )
        else:
            status = RemediationStatus.SKIPPED
            desc = f"Shape {shape_id} was preserved without modifications."

        return RemediationAction(
            action_id=action_id,
            rule_id=self.rule_id,
            slide_number=slide.slide_number,
            description=desc,
            status=status,
            is_user_approved=True,
            confidence=0.95,
            original_state={"shape_id": shape_id},
            remediated_state={
                "converted_shape_ids": result.converted_shape_ids,
                "preserved_shape_ids": result.preserved_shape_ids,
                "model_traceability": {
                    "antigravity_coding_model": "Gemini 3.8 Flash High",
                    "slideguard_runtime_model": "Gemini 3.5",
                    "runtime_model_identifier": getattr(self.classifier.assistant, "model_name", "gemini-3.5")
                }
            }
        )

    def _get_slide_anim_targets(self, slide: SlideModel) -> List[int]:
        """Extracts animation target shape IDs for the slide."""
        targets: List[int] = []
        raw_slide = slide.raw_slide
        if raw_slide is not None and hasattr(raw_slide, "_element"):
            elem = raw_slide._element
            for tgt in _safe_xpath(elem, ".//p:timing//p:spTgt"):
                spid = tgt.get("spid")
                if spid:
                    try:
                        targets.append(int(spid))
                    except ValueError:
                        pass
        return targets


# Backward-compatible alias
PlaceholderConversionRule = PlaceholderRule

__all__ = ["PlaceholderRule", "PlaceholderConversionRule"]
