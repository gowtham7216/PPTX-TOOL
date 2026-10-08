"""
SlideGuard Accessibility Rules Package (Phase 7).
Exports all Phase 7 accessibility validation and safe remediation rules.
"""

from rules.accessibility.base import BaseAccessibilityRule
from rules.accessibility.alt_text import AltTextRule
from rules.accessibility.decorative import DecorativeRule
from rules.accessibility.titles import SlideTitleRule
from rules.accessibility.tables import TableAccessibilityRule
from rules.accessibility.contrast import TextContrastRule
from rules.accessibility.readability import ReadabilityRule
from rules.accessibility.missing_text import MissingTextRule
from rules.accessibility.language import LanguageRule
from rules.accessibility.metadata import DocumentMetadataRule
from rules.accessibility.reading_order_validation import ReadingOrderValidationRule

__all__ = [
    "BaseAccessibilityRule",
    "AltTextRule",
    "DecorativeRule",
    "SlideTitleRule",
    "TableAccessibilityRule",
    "TextContrastRule",
    "ReadabilityRule",
    "MissingTextRule",
    "LanguageRule",
    "DocumentMetadataRule",
    "ReadingOrderValidationRule",
]
