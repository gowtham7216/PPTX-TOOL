"""
SlideGuard Font Readability Rule (Phase 7).
Analyzes font sizes and identifies excessively small text or empty text frames.
Configured strictly as DETECT_ONLY to preserve designer typography.
"""

from typing import List
from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, ShapeType, AccessibilityStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue


class ReadabilityRule(BaseAccessibilityRule):
    """
    Checks for minimum legible font sizes and identifies empty text boxes.
    Classified strictly as DETECT_ONLY to preserve original slide typography.
    """

    @property
    def rule_id(self) -> str:
        return "SG-RDB-01"

    @property
    def name(self) -> str:
        return "Typography & Legibility Check"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.DETECT_ONLY

    @property
    def wcag_mapping(self) -> str:
        return "Accessibility Best Practice (Readability)"

    @property
    def description(self) -> str:
        return "Identifies excessively small fonts (<10-12pt) and detects empty text boxes."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        for shape in slide.shapes:
            # Check for empty text box (screen reader nuisance)
            if shape.shape_type == ShapeType.TEXT_BOX and not shape.has_text:
                issues.append(
                    Issue(
                        issue_id=f"RDB_EMPTY_{slide.slide_number}_{shape.shape_id}",
                        rule_id=self.rule_id,
                        rule_name=self.name,
                        slide_number=slide.slide_number,
                        severity=IssueSeverity.INFO,
                        shape_id=shape.shape_id,
                        shape_name=shape.name,
                        wcag_mapping=self.wcag_mapping,
                        description=f"Empty text box '{shape.name}' detected on slide {slide.slide_number}.",
                        recommendation="Remove empty shape to prevent confusing screen reader announcement.",
                        behavior=self.behavior,
                        status=AccessibilityStatus.MANUAL_REVIEW,
                        confidence="HIGH",
                        explanation="Text box contains no text runs or content.",
                        extra_data={"target_shape_id": shape.shape_id}
                    )
                )

            # Check for very small font sizes (<10pt)
            if shape.paragraphs:
                for p in shape.paragraphs:
                    for r in p.runs:
                        if r.font_size_pt and r.font_size_pt < 10.0 and r.text.strip():
                            issues.append(
                                Issue(
                                    issue_id=f"RDB_SMALL_FONT_{slide.slide_number}_{shape.shape_id}",
                                    rule_id=self.rule_id,
                                    rule_name=self.name,
                                    slide_number=slide.slide_number,
                                    severity=IssueSeverity.WARNING,
                                    shape_id=shape.shape_id,
                                    shape_name=shape.name,
                                    wcag_mapping=self.wcag_mapping,
                                    description=f"Text run '{r.text[:20]}...' uses small font size ({r.font_size_pt:.1f}pt).",
                                    recommendation="Increase font size to at least 14pt (or 12pt minimum) for legibility.",
                                    behavior=self.behavior,
                                    status=AccessibilityStatus.MANUAL_REVIEW,
                                    confidence="HIGH",
                                    explanation=f"Text size of {r.font_size_pt:.1f}pt falls below accessible presentation threshold.",
                                    extra_data={"font_size_pt": r.font_size_pt}
                                )
                            )
                            break

        return issues
