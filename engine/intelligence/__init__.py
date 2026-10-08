"""
SlideGuard Accessibility Intelligence Package (Phase 8).
Exports context classifiers, confidence engines, issue correlators,
Gemini assistant, and before/after reconciliation services.
"""

from engine.intelligence.context_classifier import (
    ContextualClassifier,
    TitleContext,
    ImageContext,
    TextContext,
    TableContext,
)
from engine.intelligence.confidence_engine import ConfidenceEngine
from engine.intelligence.correlator import IssueCorrelator
from engine.intelligence.gemini_assistant import GeminiAssistant, AIRecommendation
from engine.intelligence.reconciliation import (
    IssueReconciliationService,
    ReconciliationSummary,
    ReconciledItem,
)

__all__ = [
    "ContextualClassifier",
    "TitleContext",
    "ImageContext",
    "TextContext",
    "TableContext",
    "ConfidenceEngine",
    "IssueCorrelator",
    "GeminiAssistant",
    "AIRecommendation",
    "IssueReconciliationService",
    "ReconciliationSummary",
    "ReconciledItem",
]
