"""
SlideGuard Decorative Element Rule (Phase 7).
Identifies visual shapes (dividers, background accent boxes) that should be tagged as decorative.
Provides safe deterministic remediation for unambiguous dividers.
"""

from typing import List
from lxml import etree

from rules.base import BaseAccessibilityRule
from rules.accessibility.base import get_cnvpr_element
from models.enums import RuleBehavior, IssueSeverity, ShapeType, AccessibilityStatus, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.remediation import RemediationAction


class DecorativeRule(BaseAccessibilityRule):
    """
    Identifies purely decorative non-text visual shapes (such as thin separator lines
    or solid background color rectangles) that should be explicitly marked decorative.
    """

    @property
    def rule_id(self) -> str:
        return "SG-DEC-01"

    @property
    def name(self) -> str:
        return "Decorative Element Tagging"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.SAFE_AUTO_REMEDIATE

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.1.1 Non-text Content (Level A)"

    @property
    def description(self) -> str:
        return "Ensures visual embellishments and dividers are properly designated as decorative."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        for shape in slide.shapes:
            if shape.shape_type == ShapeType.AUTO_SHAPE and not shape.has_text:
                # Protect pictures, diagrams, and complex graphic shapes
                if shape.is_decorative:
                    continue

                # Inspect if shape is a divider line or accent bar
                is_divider = (shape.height_pt <= 4.0 or shape.width_pt <= 4.0)

                if is_divider:
                    issues.append(
                        Issue(
                            issue_id=f"DEC_DIVIDER_{slide.slide_number}_{shape.shape_id}",
                            rule_id=self.rule_id,
                            rule_name=self.name,
                            slide_number=slide.slide_number,
                            severity=IssueSeverity.INFO,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            wcag_mapping=self.wcag_mapping,
                            description=f"Shape '{shape.name}' appears to be a divider/accent line and should be marked as decorative.",
                            recommendation="Tag shape as decorative so assistive technologies safely ignore it.",
                            behavior=self.behavior,
                            status=AccessibilityStatus.SAFE_TO_REMEDIATE,
                            confidence="HIGH",
                            explanation=f"AutoShape has zero text and narrow dimension (h={shape.height_pt}pt, w={shape.width_pt}pt).",
                            extra_data={"is_divider": True, "target_shape_id": shape.shape_id}
                        )
                    )

        return issues

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """Deterministically marks divider shape as decorative by setting empty descr in OOXML."""
        target_id = issue.shape_id
        if target_id is None or slide.raw_slide is None:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=target_id,
                description="Unable to locate target shape for decorative tagging.",
                status=RemediationStatus.FAILED,
                confidence=0.0
            )

        try:
            # Locate shape element in slide tree
            sp_tree = slide.raw_slide.shapes._spTree
            target_elem = None
            for child in sp_tree:
                cnvpr = get_cnvpr_element(child)
                if cnvpr is not None and cnvpr.get("id") == str(target_id):
                    target_elem = child
                    cnvpr.set("descr", "")
                    break

            # Update in-memory shape model if present
            for s in slide.shapes:
                if s.shape_id == target_id:
                    s.is_decorative = True
                    s.alt_text = ""

            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=target_id,
                description=f"Marked shape {target_id} as decorative in OOXML (descr='').",
                status=RemediationStatus.APPLIED,
                confidence=0.98,
                is_user_approved=True
            )
        except Exception as e:
            return RemediationAction(
                action_id=f"rem_{issue.issue_id}",
                rule_id=self.rule_id,
                slide_number=slide.slide_number,
                target_shape_id=target_id,
                description=f"Failed to tag shape as decorative: {str(e)}",
                status=RemediationStatus.FAILED,
                confidence=0.0,
                error_message=str(e)
            )
