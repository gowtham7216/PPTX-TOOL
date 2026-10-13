"""
SlideGuard Placeholders Rule Package.
"""

from rules.placeholders.rule import PlaceholderRule, PlaceholderConversionRule
from rules.placeholders.classifier import (
    PlaceholderClassifier,
    PlaceholderCategory,
    PlaceholderStatus,
    PlaceholderAnalysis,
)
from rules.placeholders.inheritance import (
    PlaceholderInheritanceResolver,
    InheritedProperties,
)
from rules.placeholders.transformer import (
    PlaceholderTransformer,
    ConversionResult,
)
from rules.placeholders.placeholder_engine_v2 import (
    PlaceholderEngineV2,
    PlaceholderInsertionResult,
)

__all__ = [
    "PlaceholderRule",
    "PlaceholderConversionRule",
    "PlaceholderClassifier",
    "PlaceholderCategory",
    "PlaceholderStatus",
    "PlaceholderAnalysis",
    "PlaceholderInheritanceResolver",
    "InheritedProperties",
    "PlaceholderTransformer",
    "ConversionResult",
    "PlaceholderEngineV2",
    "PlaceholderInsertionResult",
]
