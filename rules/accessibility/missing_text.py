"""
SlideGuard Missing / Empty Text Content Rule (Phase 7).
Detects empty text boxes, whitespace-only paragraphs, and potential missing textual content.
Strictly configured as DETECT_ONLY to avoid accidental deletion of intentional layout boxes.
"""

from typing import List
from rules.base import BaseAccessibilityRule
from models.enums import RuleBehavior, IssueSeverity, ShapeType, AccessibilityStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue


class MissingTextRule(BaseAccessibilityRule):
    """
    Identifies empty or whitespace-only text containers.
    Classifies shapes into INTENTIONAL, POSSIBLE_MISSING_CONTENT, or UNKNOWN.
    """

    @property
    def rule_id(self) -> str:
        return "SG-TXT-01"

    @property
    def name(self) -> str:
        return "Missing or Empty Text Content"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.DETECT_ONLY

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.3.1 Info and Relationships (Level A)"

    @property
    def description(self) -> str:
        return "Flags text containers containing only whitespace or zero characters."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        for shape in slide.shapes:
            if shape.shape_type in (ShapeType.TEXT_BOX, ShapeType.PLACEHOLDER):
                # Calculate total visible characters across all paragraphs
                text_content = ""
                for p in shape.paragraphs:
                    for r in p.runs:
                        text_content += r.text

                is_empty = (not shape.has_text) or (not text_content.strip())

                if is_empty:
                    # Classify if likely intentional or missing content
                    classification = "POSSIBLE_MISSING_CONTENT"
                    if shape.is_decorative or shape.shape_type == ShapeType.AUTO_SHAPE:
                        classification = "INTENTIONAL_OR_DECORATIVE"

                    issues.append(
                        Issue(
                            issue_id=f"TXT_EMPTY_{slide.slide_number}_{shape.shape_id}",
                            rule_id=self.rule_id,
                            rule_name=self.name,
                            slide_number=slide.slide_number,
                            severity=IssueSeverity.INFO,
                            shape_id=shape.shape_id,
                            shape_name=shape.name,
                            wcag_mapping=self.wcag_mapping,
                            description=f"Shape '{shape.name}' contains empty or whitespace-only content.",
                            recommendation="Populate with meaningful content or remove shape if redundant.",
                            behavior=self.behavior,
                            status=AccessibilityStatus.MANUAL_REVIEW,
                            confidence="HIGH",
                            explanation=f"Text container evaluated with classification '{classification}'.",
                            extra_data={
                                "target_shape_id": shape.shape_id,
                                "classification": classification
                            }
                        )
                    )

        return issues
