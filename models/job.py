"""
SlideGuard Job & Metrics Data Models.
Manages isolated job context and aggregated quality metrics.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from pathlib import Path
from datetime import datetime
from models.enums import JobStage
from models.rule_result import RuleResult
from models.remediation import RemediationAction
from models.verification import VerificationResult
from models.comparison import ComparisonResult


@dataclass
class QualityMetrics:
    """Documented, auditable accessibility quality score metrics."""
    total_rules_evaluated: int = 0
    passed_rules: int = 0
    failed_rules: int = 0
    warning_rules: int = 0
    total_issues: int = 0
    critical_issues: int = 0
    error_issues: int = 0
    warning_issues: int = 0
    info_issues: int = 0
    accessibility_score: float = 100.0  # 0.0 to 100.0 score based on weighted deductions
    # Quality Score V2 fields
    quality_score_v2: float = 100.0
    correlated_issues_count: int = 0
    high_confidence_issues: int = 0
    medium_confidence_issues: int = 0
    low_confidence_issues: int = 0
    manual_review_issues: int = 0
    remediated_issues: int = 0
    unresolved_issues: int = 0
    unsupported_issues: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "total_rules_evaluated": self.total_rules_evaluated,
            "passed_rules": self.passed_rules,
            "failed_rules": self.failed_rules,
            "warning_rules": self.warning_rules,
            "total_issues": self.total_issues,
            "critical_issues": self.critical_issues,
            "error_issues": self.error_issues,
            "warning_issues": self.warning_issues,
            "info_issues": self.info_issues,
            "accessibility_score": round(self.accessibility_score, 1),
            "quality_score_v2": round(self.quality_score_v2, 1),
            "correlated_issues_count": self.correlated_issues_count,
            "high_confidence_issues": self.high_confidence_issues,
            "medium_confidence_issues": self.medium_confidence_issues,
            "low_confidence_issues": self.low_confidence_issues,
            "manual_review_issues": self.manual_review_issues,
            "remediated_issues": self.remediated_issues,
            "unresolved_issues": self.unresolved_issues,
            "unsupported_issues": self.unsupported_issues,
        }


@dataclass
class JobContext:
    """Runtime context for an isolated execution job."""
    job_id: str
    workspace_dir: Path
    original_file: Path
    working_file: Path
    candidate_file: Optional[Path] = None
    output_file: Optional[Path] = None
    source_file: Optional[Path] = None
    final_output_file: Optional[Path] = None
    final_report_file: Optional[Path] = None
    reports_dir: Optional[Path] = None
    logs_dir: Optional[Path] = None
    stage: JobStage = JobStage.INITIALIZED
    total_slides: int = 0
    slides_analyzed: int = 0
    slides_with_errors: int = 0
    rule_results: List[RuleResult] = field(default_factory=list)
    remediation_actions: List[RemediationAction] = field(default_factory=list)
    verification_result: Optional[VerificationResult] = None
    comparison_result: Optional[ComparisonResult] = None
    metrics: QualityMetrics = field(default_factory=QualityMetrics)
    user_approved_remediation: bool = False
    started_at: datetime = field(default_factory=datetime.utcnow)
    completed_at: Optional[datetime] = None
    error_message: Optional[str] = None
    # Phase 8 Accessibility Intelligence fields
    unified_issues: List[Any] = field(default_factory=list)
    intelligence_summary: Dict[str, Any] = field(default_factory=dict)
    pre_remediation_score: Optional[float] = None
    post_remediation_score: Optional[float] = None
    # Phase 9 AI-Assisted Authoring fields
    ai_recommendations: List[Any] = field(default_factory=list)
    ai_audit_trail: List[Dict[str, Any]] = field(default_factory=list)
    ai_authoring_summary: Dict[str, Any] = field(default_factory=dict)
    # Phase 10 QA & Validation result
    qa_result: Optional[Any] = None
    # Task 4 Comprehensive Placeholder Audit Report
    task_4_placeholder_report: Optional[Dict[str, Any]] = None

    @property
    def duration_seconds(self) -> float:
        if self.completed_at:
            return (self.completed_at - self.started_at).total_seconds()
        return (datetime.utcnow() - self.started_at).total_seconds()
