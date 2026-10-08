"""
SlideGuard Comparison Data Model.
Captures semantic comparison between original and candidate/remediated PPTX files.
Distinguishes between intentional accessibility adjustments and unintended visual deviations.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any


@dataclass
class ComparisonDifference:
    """A single detected change between original and remediated presentation."""
    slide_number: int
    shape_id: int
    attribute: str
    original_value: Any
    remediated_value: Any
    is_intentional_accessibility_change: bool
    description: str


@dataclass
class ComparisonResult:
    """Outcome of structural/semantic comparison."""
    original_slide_count: int = 0
    remediated_slide_count: int = 0
    slide_count_matches: bool = True
    differences: List[ComparisonDifference] = field(default_factory=list)
    unintended_changes_count: int = 0
    intentional_changes_count: int = 0
    integrity_preserved: bool = True

    def add_difference(self, diff: ComparisonDifference):
        self.differences.append(diff)
        if diff.is_intentional_accessibility_change:
            self.intentional_changes_count += 1
        else:
            self.unintended_changes_count += 1
            self.integrity_preserved = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "original_slide_count": self.original_slide_count,
            "remediated_slide_count": self.remediated_slide_count,
            "slide_count_matches": self.slide_count_matches,
            "integrity_preserved": self.integrity_preserved,
            "intentional_changes_count": self.intentional_changes_count,
            "unintended_changes_count": self.unintended_changes_count,
            "differences": [
                {
                    "slide_number": d.slide_number,
                    "shape_id": d.shape_id,
                    "attribute": d.attribute,
                    "original_value": str(d.original_value),
                    "remediated_value": str(d.remediated_value),
                    "is_intentional": d.is_intentional_accessibility_change,
                    "description": d.description
                }
                for d in self.differences
            ]
        }
