"""
SlideGuard List Structure Rule.
Detects and remediates pseudo-lists typed manually with dashes, numbers, or letters,
while preserving native and inherited list styles (bullets, numbers, alphabetical, multi-level).
WCAG 2.1 - 1.3.1 Info and Relationships (Level A).
"""

from typing import List, Optional, Dict, Any, Set
import logging

from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction
from rules.lists.classifier import (
    ListClassifier,
    ParagraphListAnalysis,
    ListCategory,
    ListStatus,
    _safe_xpath,
)
from rules.lists.transformer import ListTransformer, ListRemediationResult

logger = logging.getLogger(__name__)


class ListRule(BaseAccessibilityRule):
    """
    Validates and remediates semantic list markup in PowerPoint text frames.
    Flags manually simulated lists that screen readers treat as plain sentences,
    and executes safe in-place OpenXML promotion without visual distortion.
    """

    def __init__(
        self,
        classifier: Optional[ListClassifier] = None,
        transformer: Optional[ListTransformer] = None
    ):
        self.classifier = classifier or ListClassifier()
        self.transformer = transformer or ListTransformer(classifier=self.classifier)

    @property
    def rule_id(self) -> str:
        return "SG-LST-01"

    @property
    def name(self) -> str:
        return "Semantic List Formatting"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.1 Info and Relationships (Level A)"

    @property
    def description(self) -> str:
        return (
            "Ensures lists use native PowerPoint bullet and numbering semantics "
            "rather than typed prefixes, preserving accessibility and screen reader navigation."
        )

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []
        raw_slide = slide.raw_slide
        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return issues

        analyses = self.classifier.analyze_slide_lists(
            raw_slide=raw_slide,
            slide_number=slide.slide_number,
            slide_model=slide
        )

        for a in analyses:
            extra = a.to_dict()

            if a.status == ListStatus.SAFE_TO_REMEDIATE:
                issues.append(
                    Issue(
                        issue_id=f"LST_PSEUDO_{slide.slide_number}_{a.shape_id}_{a.paragraph_index}",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=slide.slide_number,
                        severity=IssueSeverity.WARNING,
                        shape_id=a.shape_id,
                        shape_name=a.shape_name,
                        wcag_mapping=self.wcag_mapping,
                        description=(
                            f"Slide {slide.slide_number} Shape {a.shape_id} ('{a.shape_name}'): "
                            f"Paragraph {a.paragraph_index + 1} contains simulated list / pseudo-list prefix '{a.pseudo_prefix}' "
                            f"without native PowerPoint list tags."
                        ),
                        recommendation=a.proposed_action,
                        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
                        extra_data=extra
                    )
                )

            elif a.status == ListStatus.MANUAL_REVIEW:
                issues.append(
                    Issue(
                        issue_id=f"LST_MANUAL_{slide.slide_number}_{a.shape_id}_{a.paragraph_index}",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=slide.slide_number,
                        severity=IssueSeverity.INFO,
                        shape_id=a.shape_id,
                        shape_name=a.shape_name,
                        wcag_mapping=self.wcag_mapping,
                        description=(
                            f"Slide {slide.slide_number} Shape {a.shape_id} P{a.paragraph_index + 1}: {a.reason}"
                        ),
                        recommendation=a.proposed_action,
                        behavior=RuleBehavior.MANUAL_REVIEW,
                        extra_data=extra
                    )
                )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """Executes in-place semantic list conversion for an approved issue."""
        extra = issue.extra_data or {}
        shape_id = issue.shape_id or extra.get("shape_id")
        p_idx = extra.get("paragraph_index", 0)
        action_id = f"act_lst_{slide.slide_number}_{shape_id}_{p_idx}"

        raw_slide = slide.raw_slide
        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Cannot remediate Shape {shape_id}: Missing raw slide element.",
                status=RemediationStatus.FAILED,
                error_message="Raw slide element is None."
            )

        # Check if manual review item was skipped
        if issue.behavior == RuleBehavior.MANUAL_REVIEW:
            reason = extra.get("reason", "Flagged for manual review.")
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Preserved Shape {shape_id} P{p_idx + 1}: {reason}",
                status=RemediationStatus.SKIPPED,
                error_message=reason
            )

        slide_elem = raw_slide._element
        target_pair = (shape_id, p_idx)

        res: ListRemediationResult = self.transformer.remediate_slide(
            slide_elem=slide_elem,
            slide_number=slide.slide_number,
            raw_slide=raw_slide,
            approved_targets={target_pair}
        )

        if not res.success or res.errors:
            err = "; ".join(res.errors)
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"List remediation failed on Shape {shape_id} P{p_idx + 1}.",
                status=RemediationStatus.FAILED,
                error_message=err
            )

        if res.paragraphs_remediated > 0:
            status = RemediationStatus.APPLIED
            desc = (
                f"Successfully converted visual list on Slide {slide.slide_number} Shape {shape_id} "
                f"Paragraph {p_idx + 1} to native PowerPoint list semantics."
            )
        else:
            status = RemediationStatus.SKIPPED
            desc = f"Shape {shape_id} Paragraph {p_idx + 1} was preserved without modifications."

        return RemediationAction(
            action_id=action_id,
            rule_id=self.rule_id,
            slide_number=slide.slide_number,
            description=desc,
            status=status,
            is_user_approved=True,
            confidence=0.95,
            original_state={"shape_id": shape_id, "paragraph_index": p_idx, "extra_data": extra},
            remediated_state={"remediated_targets": res.remediated_targets}
        )
