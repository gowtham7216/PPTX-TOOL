"""
SlideGuard Core Engine Interfaces.
Defines explicit contracts for Rules, Readers, Writers, Workspace Managers, and AI Providers.
"""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Optional, List, Dict, Any

from models.enums import RuleBehavior
from models.shape import ShapeModel
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.rule_result import RuleResult
from models.remediation import RemediationAction
from models.job import JobContext, QualityMetrics


class IAccessibilityRule(ABC):
    """Contract for any SlideGuard accessibility rule."""

    @property
    @abstractmethod
    def rule_id(self) -> str:
        """Unique rule identifier (e.g., 'RULE-READING-ORDER-01')."""
        pass

    @property
    @abstractmethod
    def name(self) -> str:
        """Human-readable rule name."""
        pass

    @property
    @abstractmethod
    def behavior(self) -> RuleBehavior:
        """Execution mode: DETECT_ONLY, SAFE_AUTO_REMEDIATE, or AI_ASSISTED_REVIEW."""
        pass

    @property
    @abstractmethod
    def wcag_mapping(self) -> Optional[str]:
        """Configurable WCAG success criterion mapping (e.g., '1.3.2 Meaningful Sequence')."""
        pass

    @abstractmethod
    def analyze_slide(self, slide: SlideModel, presentation: PresentationModel) -> RuleResult:
        """Evaluates accessibility rule on a single slide."""
        pass

    @abstractmethod
    def remediate_issue(
        self,
        slide: SlideModel,
        presentation: PresentationModel,
        issue: Issue
    ) -> RemediationAction:
        """
        Attempts deterministic or approved remediation for an identified issue.
        Should raise NotImplementedError or return RemediationStatus.NOT_IMPLEMENTED if not yet ready.
        """
        pass


class IPresentationReader(ABC):
    """Contract for parsing PPTX files into internal models."""

    @abstractmethod
    def read(self, file_path: Path) -> PresentationModel:
        """Parses a PPTX file and returns a structured PresentationModel."""
        pass


class IPresentationWriter(ABC):
    """Contract for writing a presentation back to PPTX file."""

    @abstractmethod
    def write(self, presentation: PresentationModel, output_path: Path) -> None:
        """Saves presentation safely to the target output path."""
        pass


class IWorkspaceManager(ABC):
    """Contract for managing isolated execution workspaces per job."""

    @abstractmethod
    def create_job(self, source_file: Path) -> JobContext:
        """Creates isolated workspace directories and working copy."""
        pass

    @abstractmethod
    def cleanup_job(self, job_context: JobContext, preserve_on_debug: bool = True) -> None:
        """Safely cleans up temporary job files."""
        pass


class IAIProvider(ABC):
    """
    Contract for optional AI assistance.
    AI is strictly advisory and NEVER modifies PPTX directly.
    """

    @property
    @abstractmethod
    def is_available(self) -> bool:
        """Returns True if the provider is configured and reachable."""
        pass

    @abstractmethod
    def suggest_alt_text(self, image_bytes: bytes, context_text: Optional[str] = None) -> Optional[str]:
        """Advisory alt-text suggestion requiring user confirmation."""
        pass

    @abstractmethod
    def classify_decorative(self, shape_attributes: Dict[str, Any]) -> Optional[bool]:
        """Advisory decorative classification requiring user confirmation."""
        pass


class IScoringEngine(ABC):
    """Contract for calculating auditable accessibility quality scores."""

    @abstractmethod
    def calculate_metrics(self, rule_results: List[RuleResult]) -> QualityMetrics:
        """Computes transparent, documented quality score from rule results."""
        pass
