"""
SlideGuard Selective Ungrouping Rule.
Inspects presentation slides for grouped shapes, classifying them into
eligible text/shape elements vs. protected picture containers.
Provides safe, deterministic in-place OpenXML selective ungrouping.
"""

from typing import List, Optional, Dict, Any
import logging

from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction
from rules.ungrouping.classifier import _safe_xpath, GroupClassifier, GroupAnalysis, GroupCategory
from rules.ungrouping.transformer import SelectiveUngroupingTransformer, UngroupResult

logger = logging.getLogger(__name__)


class SelectiveUngroupingRule(BaseAccessibilityRule):
    """
    Validates and remediates grouped PowerPoint objects.
    Selectively ungroups text boxes, autoshapes, and connectors to expose them
    to screen readers while keeping picture elements strictly grouped and protected.
    """

    def __init__(
        self,
        classifier: Optional[GroupClassifier] = None,
        transformer: Optional[SelectiveUngroupingTransformer] = None
    ):
        self.classifier = classifier or GroupClassifier()
        self.transformer = transformer or SelectiveUngroupingTransformer(classifier=self.classifier)

    @property
    def rule_id(self) -> str:
        return "SG-GRP-01"

    @property
    def name(self) -> str:
        return "Selective Ungrouping for Accessible Text"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.1 Info and Relationships (Level A)"

    @property
    def description(self) -> str:
        return (
            "Detects grouped objects hiding text content from assistive technology. "
            "Selectively ungroups text and auto-shapes while preserving image groupings."
        )

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []
        raw_slide = slide.raw_slide
        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return issues

        slide_elem = raw_slide._element
        anim_targets = self._get_slide_anim_targets(slide, presentation)

        # Inspect top-level groups
        top_groups = _safe_xpath(slide_elem, ".//p:spTree/p:grpSp")

        for grp_el in top_groups:
            analysis: GroupAnalysis = self.classifier.analyze_group(
                group_elem=grp_el,
                slide_number=slide.slide_number,
                animation_target_ids=anim_targets
            )

            # Skip picture-only or empty groups without creating remediation issues
            if analysis.category in (GroupCategory.PICTURE_ONLY, GroupCategory.EMPTY):
                continue

            # If the group has eligible text/shape content:
            if analysis.eligible_count > 0:
                if analysis.category == GroupCategory.SHAPE_TEXT_ONLY:
                    desc = (
                        f"Slide {slide.slide_number}: Group '{analysis.group_name}' (ID {analysis.group_id}) "
                        f"contains {analysis.total_children} accessible text/shape element(s) that should be ungrouped."
                    )
                else:
                    desc = (
                        f"Slide {slide.slide_number}: Mixed Group '{analysis.group_name}' (ID {analysis.group_id}) "
                        f"contains {analysis.eligible_count} text/shape element(s) mixed with {analysis.picture_count} "
                        f"protected picture(s). Selective ungrouping will extract text while protecting pictures."
                    )

                if not analysis.is_safe_to_remediate:
                    desc += f" [MANUAL REVIEW REQUIRED: {analysis.skip_reason}]"

                rule_behavior = (
                    RuleBehavior.SAFE_AUTO_REMEDIATE
                    if analysis.is_safe_to_remediate
                    else RuleBehavior.MANUAL_REVIEW
                )

                issue = Issue(
                    issue_id=f"GRP_UNGROUP_{slide.slide_number}_{analysis.group_id}",
                    rule_id=self.rule_id,
                    rule_name=self.name,
                    slide_number=slide.slide_number,
                    severity=IssueSeverity.WARNING,
                    shape_id=analysis.group_id,
                    shape_name=analysis.group_name,
                    wcag_mapping=self.wcag_mapping,
                    description=desc,
                    recommendation="Selectively ungroup text and shape elements while maintaining picture groupings.",
                    behavior=rule_behavior,
                    extra_data=analysis.to_dict()
                )
                issues.append(issue)

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """
        Executes physical in-place selective ungrouping for a specific approved group issue.
        """
        group_id = issue.shape_id or issue.extra_data.get("group_id")
        action_id = f"act_ungroup_{slide.slide_number}_{group_id}"

        raw_slide = slide.raw_slide
        if raw_slide is None or not hasattr(raw_slide, "_element"):
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Cannot ungroup Group {group_id}: Missing raw slide element.",
                status=RemediationStatus.FAILED,
                error_message="Raw slide element is None."
            )

        # Check safety before proceeding
        if not issue.extra_data.get("is_safe_to_remediate", True):
            skip_msg = issue.extra_data.get("skip_reason", "Group flagged as unsafe for automatic remediation.")
            logger.warning("Skipping ungrouping of Group %d on Slide %d: %s", group_id, slide.slide_number, skip_msg)
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Skipped Group {group_id}: {skip_msg}",
                status=RemediationStatus.SKIPPED,
                error_message=skip_msg
            )

        slide_elem = raw_slide._element
        anim_targets = self._get_slide_anim_targets(slide, presentation)

        # Execute transformation on targeted group
        ungroup_result: UngroupResult = self.transformer.ungroup_slide(
            slide_elem=slide_elem,
            slide_number=slide.slide_number,
            animation_target_ids=anim_targets,
            approved_group_ids=[group_id] if group_id else None
        )

        if not ungroup_result.success:
            err = "; ".join(ungroup_result.errors)
            return RemediationAction(
                action_id=action_id,
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                description=f"Selective ungrouping failed on Group {group_id}.",
                status=RemediationStatus.FAILED,
                error_message=err
            )

        if ungroup_result.groups_processed > 0:
            status = RemediationStatus.APPLIED
            desc = (
                f"Successfully selectively ungrouped Group {group_id}. "
                f"Extracted {ungroup_result.shapes_extracted} shape(s) to slide; "
                f"protected {ungroup_result.pictures_protected} picture(s)."
            )
        else:
            status = RemediationStatus.SKIPPED
            desc = f"Group {group_id} was skipped during remediation."

        return RemediationAction(
            action_id=action_id,
            rule_id=self.rule_id,
            slide_number=slide.slide_number,
            description=desc,
            status=status,
            is_user_approved=True,
            confidence=0.95,
            original_state={"group_id": group_id},
            remediated_state={
                "extracted_shape_ids": ungroup_result.extracted_shape_ids,
                "retained_group_ids": ungroup_result.retained_group_ids,
                "removed_group_ids": ungroup_result.removed_group_ids,
            }
        )

    def _get_slide_anim_targets(self, slide: SlideModel, presentation: PresentationModel) -> List[int]:
        """Extracts animation target shape IDs for the given slide."""
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
