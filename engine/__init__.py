"""
SlideGuard Engine Package.
"""

from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.remediation.remediation_service import RemediationService
from engine.verification.verification_service import VerificationService
from engine.comparison.comparison_service import ComparisonService
from engine.accessibility.scoring import AccessibilityScorer

__all__ = [
    "PresentationAnalyzer",
    "RemediationService",
    "VerificationService",
    "ComparisonService",
    "AccessibilityScorer",
]
