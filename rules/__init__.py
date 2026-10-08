"""
SlideGuard Rules Package.
Provides registry of default rules and rule discovery.
"""

from typing import List
from rules.base import BaseAccessibilityRule
from rules.reading_order import ReadingOrderRule
from rules.ungrouping import UngroupingRule, SelectiveUngroupingRule
from rules.placeholders import PlaceholderRule
from rules.lists import ListRule
from rules.alt_text import AltTextRule
from rules.decorative import DecorativeRule
from rules.titles import SlideTitleRule
from rules.tables import TableAccessibilityRule
from rules.contrast import TextContrastRule
from rules.readability import ReadabilityRule
from rules.accessibility.missing_text import MissingTextRule
from rules.language import LanguageRule
from rules.metadata import DocumentMetadataRule
from rules.accessibility.reading_order_validation import ReadingOrderValidationRule


def get_default_rules() -> List[BaseAccessibilityRule]:
    """Returns initialized list of all default SlideGuard accessibility rules."""
    return [
        ReadingOrderRule(),
        UngroupingRule(),
        PlaceholderRule(),
        ListRule(),
        AltTextRule(),
        DecorativeRule(),
        SlideTitleRule(),
        TableAccessibilityRule(),
        TextContrastRule(),
        ReadabilityRule(),
        MissingTextRule(),
        LanguageRule(),
        DocumentMetadataRule(),
        ReadingOrderValidationRule(),
    ]


__all__ = [
    "BaseAccessibilityRule",
    "get_default_rules",
    "ReadingOrderRule",
    "UngroupingRule",
    "SelectiveUngroupingRule",
    "PlaceholderRule",
    "ListRule",
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
