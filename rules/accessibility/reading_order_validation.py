"""
SlideGuard Final Reading-Order Validation Rule (Phase 7).
Executes a final verification of slide reading sequences following accessibility remediation.
Verifies title first, column flow, and ensures 'CHAPTER 6' is never misclassified as a footer.
"""

from typing import List
from rules.base import BaseAccessibilityRule
from rules.reading_order.rule import ReadingOrderRule
from models.enums import RuleBehavior, IssueSeverity, AccessibilityStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue


class ReadingOrderValidationRule(BaseAccessibilityRule):
    """
    Validates final reading sequence integrity post-remediation.
    Ensures accessibility modifications do not disturb logical narrative order.
    """

    def __init__(self, ro_rule: ReadingOrderRule = None):
        self.ro_rule = ro_rule or ReadingOrderRule()

    @property
    def rule_id(self) -> str:
        return "SG-RO-VAL-01"

    @property
    def name(self) -> str:
        return "Final Reading-Order Validation"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.DETECT_ONLY

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.2 Meaningful Sequence (Level A)"

    @property
    def description(self) -> str:
        return "Verifies final reading order sequence, title precedence, and footer exclusion."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        # Delegate sequence evaluation to ReadingOrderRule
        underlying_issues = self.ro_rule.evaluate(slide, presentation)

        for ui in underlying_issues:
            issues.append(
                Issue(
                    issue_id=f"RO_VAL_{slide.slide_number}_{ui.issue_id}",
                    rule_id=self.rule_id,
                    rule_name=self.name,
                    slide_number=slide.slide_number,
                    severity=IssueSeverity.WARNING,
                    shape_id=ui.shape_id,
                    shape_name=ui.shape_name,
                    wcag_mapping=self.wcag_mapping,
                    description=f"Slide {slide.slide_number} final reading order validation note: {ui.description}",
                    recommendation=ui.recommendation,
                    behavior=RuleBehavior.DETECT_ONLY,
                    status=AccessibilityStatus.MANUAL_REVIEW,
                    confidence=ui.confidence,
                    explanation="Final post-remediation reading sequence check.",
                    extra_data=ui.extra_data
                )
            )

        return issues
