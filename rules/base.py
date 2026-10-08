"""
SlideGuard Base Accessibility Rule.
Establishes the foundation for all modular accessibility rules.
"""

import time
import logging
from abc import abstractmethod
from typing import List, Optional

from engine.common.interfaces import IAccessibilityRule
from models.enums import RuleBehavior, RuleStatus, RemediationStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.rule_result import RuleResult
from models.remediation import RemediationAction

logger = logging.getLogger(__name__)


class BaseAccessibilityRule(IAccessibilityRule):
    """
    Abstract base class for all SlideGuard accessibility rules.
    Provides timing, error isolation, and standardized result reporting.
    """

    @property
    @abstractmethod
    def rule_id(self) -> str:
        pass

    @property
    @abstractmethod
    def name(self) -> str:
        pass

    @property
    @abstractmethod
    def behavior(self) -> RuleBehavior:
        pass

    @property
    def wcag_mapping(self) -> Optional[str]:
        return None

    @property
    def description(self) -> str:
        return ""

    @abstractmethod
    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        """Subclasses implement this method to inspect the slide and return detected issues."""
        pass

    def analyze_slide(self, slide: SlideModel, presentation: PresentationModel) -> RuleResult:
        """
        Executes analysis on a slide with error handling and timing metrics.
        Guarantees that a failure in one rule does not crash the entire process.
        """
        start_time = time.perf_counter()
        try:
            issues = self.evaluate(slide, presentation)
            elapsed_ms = (time.perf_counter() - start_time) * 1000.0

            status = RuleStatus.PASSED if len(issues) == 0 else RuleStatus.FAILED
            message = "Compliant" if status == RuleStatus.PASSED else f"Found {len(issues)} issue(s)"

            return RuleResult(
                rule_id=self.rule_id,
                rule_name=self.name,
                slide_number=slide.slide_number,
                status=status,
                issues=issues,
                execution_time_ms=round(elapsed_ms, 2),
                message=message,
                wcag_mapping=self.wcag_mapping
            )
        except Exception as e:
            elapsed_ms = (time.perf_counter() - start_time) * 1000.0
            logger.error("Error executing rule %s on slide %d: %s", self.rule_id, slide.slide_number, e, exc_info=True)
            return RuleResult(
                rule_id=self.rule_id,
                rule_name=self.name,
                slide_number=slide.slide_number,
                status=RuleStatus.ERROR,
                issues=[],
                execution_time_ms=round(elapsed_ms, 2),
                message=f"Rule execution error: {str(e)}",
                wcag_mapping=self.wcag_mapping
            )

    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """
        Default remediation implementation.
        Returns NOT_IMPLEMENTED until dedicated remediation phase is reached.
        """
        return RemediationAction(
            action_id=f"rem_{issue.issue_id}",
            rule_id=self.rule_id,
            slide_number=slide.slide_number,
            target_shape_id=issue.shape_id,
            description=f"Remediation for {self.rule_id} is not yet implemented in Phase 1.",
            status=RemediationStatus.NOT_IMPLEMENTED,
            confidence=0.0,
            error_message="Remediation not implemented yet"
        )
