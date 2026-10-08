"""
Alt Text Rule Package.
"""

from rules.alt_text.rule import AltTextRule
from rules.alt_text.models import (
    ImageCandidate,
    ImageDecision,
    ImageClassification,
    ImageAction,
    GenerationBasis,
    QualityClass,
)
from rules.alt_text.extractor import ImageExtractor
from rules.alt_text.service import ImageAccessibilityService
from rules.alt_text.transformer import ImageAltTextTransformer
from rules.alt_text.reporter import generate_task_3_report

__all__ = [
    "AltTextRule",
    "ImageCandidate",
    "ImageDecision",
    "ImageClassification",
    "ImageAction",
    "GenerationBasis",
    "QualityClass",
    "ImageExtractor",
    "ImageAccessibilityService",
    "ImageAltTextTransformer",
    "generate_task_3_report",
]
