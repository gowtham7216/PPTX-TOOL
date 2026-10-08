"""
SlideGuard Presentation Analyzer.
Coordinates rule execution across all presentation slides with strict slide-level error isolation.
"""

from typing import List, Optional, Callable, TYPE_CHECKING
import logging

from models.presentation import PresentationModel
from models.rule_result import RuleResult
from models.enums import RuleStatus
from models.job import QualityMetrics
from engine.accessibility.scoring import AccessibilityScorer

if TYPE_CHECKING:
    from rules.base import BaseAccessibilityRule

logger = logging.getLogger(__name__)


class PresentationAnalyzer:
    """
    Executes configured accessibility rules across slides in a presentation.
    Guarantees slide-level failure isolation: an exception on slide N will not halt slides N+1.
    """

    def __init__(
        self,
        rules: Optional[List["BaseAccessibilityRule"]] = None,
        scorer: Optional[AccessibilityScorer] = None
    ):
        if rules is not None:
            self.rules = rules
        else:
            from rules import get_default_rules
            self.rules = get_default_rules()
        self.scorer = scorer or AccessibilityScorer()

    def analyze(
        self,
        presentation: PresentationModel,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> List[RuleResult]:
        total_slides = presentation.total_slides
        all_results: List[RuleResult] = []

        logger.info("Starting analysis of presentation %s (%d slides, %d rules)",
                    presentation.file_path.name, total_slides, len(self.rules))

        for idx, slide in enumerate(presentation.slides, start=1):
            if progress_callback:
                progress_callback(idx, total_slides, f"Analyzing Slide {idx} of {total_slides}...")

            for rule in self.rules:
                try:
                    result = rule.analyze_slide(slide, presentation)
                    all_results.append(result)
                    if result.has_issues:
                        logger.debug("Slide %d Rule %s found %d issues",
                                     slide.slide_number, rule.rule_id, len(result.issues))
                except Exception as e:
                    logger.error("Unexpected error on slide %d executing rule %s: %s",
                                 slide.slide_number, rule.rule_id, e, exc_info=True)
                    # Even if an unhandled error happens inside analyze_slide wrapper, isolate it:
                    error_result = RuleResult(
                        rule_id=rule.rule_id,
                        rule_name=rule.name,
                        slide_number=slide.slide_number,
                        status=RuleStatus.ERROR,
                        message=f"Slide analysis failed: {str(e)}",
                        wcag_mapping=rule.wcag_mapping
                    )
                    all_results.append(error_result)

        if progress_callback:
            progress_callback(total_slides, total_slides, "Analysis complete.")

        logger.info("Analysis completed. Total rule results: %d", len(all_results))
        return all_results

    def compute_metrics(self, rule_results: List[RuleResult], total_slides: int) -> QualityMetrics:
        """Calculates aggregated quality metrics using the scoring engine."""
        return self.scorer.calculate_metrics(rule_results, total_slides)
