"""
SlideGuard QA & Validation Data Models (Phase 10).
Defines structured quality assurance status, test results, and compliance metrics.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict, Any, Optional
from datetime import datetime


class QAStatus(str, Enum):
    """Status enumeration for QA validation suites and specific criteria."""
    PASS = "PASS"
    PASS_WITH_WARNINGS = "PASS_WITH_WARNINGS"
    FAIL = "FAIL"


@dataclass
class QACheckItem:
    """Individual QA check result item."""
    name: str
    status: QAStatus
    details: str
    component: str = "core"
    timestamp: datetime = field(default_factory=datetime.utcnow)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "status": self.status.value,
            "details": self.details,
            "component": self.component,
            "timestamp": self.timestamp.isoformat(),
        }


@dataclass
class QAResult:
    """
    Structured Quality Assurance result aggregating validation checks across all dimensions:
    OOXML package, design preservation, original file protection, relationship integrity,
    accessibility consistency, idempotency, AI fallback, workspace cleanup, report accuracy,
    and animation preservation.
    """
    job_id: str
    input_file: str
    output_file: Optional[str] = None
    test_profile: str = "NORMAL"
    tests_run: int = 0
    tests_passed: int = 0
    tests_failed: int = 0
    warnings: List[str] = field(default_factory=list)

    # Dimensional Statuses
    integrity_status: QAStatus = QAStatus.PASS
    design_preservation_status: QAStatus = QAStatus.PASS
    original_protection_status: QAStatus = QAStatus.PASS
    relationship_status: QAStatus = QAStatus.PASS
    accessibility_consistency_status: QAStatus = QAStatus.PASS
    idempotency_status: QAStatus = QAStatus.PASS
    ai_fallback_status: QAStatus = QAStatus.PASS
    workspace_cleanup_status: QAStatus = QAStatus.PASS
    report_accuracy_status: QAStatus = QAStatus.PASS
    animation_preservation_status: QAStatus = QAStatus.PASS
    overall_status: QAStatus = QAStatus.PASS

    # Execution Details and Check Items
    checks: List[QACheckItem] = field(default_factory=list)
    details: Dict[str, Any] = field(default_factory=dict)
    execution_time_seconds: float = 0.0

    def add_check(self, name: str, passed: bool, details: str = "", component: str = "core", is_warning: bool = False) -> None:
        """Helper to append an individual QA check."""
        self.tests_run += 1
        if passed:
            self.tests_passed += 1
            status = QAStatus.PASS
        elif is_warning:
            status = QAStatus.PASS_WITH_WARNINGS
            self.warnings.append(f"[{component}] {name}: {details}")
        else:
            self.tests_failed += 1
            status = QAStatus.FAIL

        self.checks.append(QACheckItem(name=name, status=status, details=details, component=component))

    def evaluate_overall(self) -> QAStatus:
        """Evaluates overall QA status from dimensional statuses and failed tests."""
        dimensional_statuses = [
            self.integrity_status,
            self.design_preservation_status,
            self.original_protection_status,
            self.relationship_status,
            self.accessibility_consistency_status,
            self.idempotency_status,
            self.ai_fallback_status,
            self.workspace_cleanup_status,
            self.report_accuracy_status,
            self.animation_preservation_status,
        ]

        if self.tests_failed > 0 or any(s == QAStatus.FAIL for s in dimensional_statuses):
            self.overall_status = QAStatus.FAIL
        elif len(self.warnings) > 0 or any(s == QAStatus.PASS_WITH_WARNINGS for s in dimensional_statuses):
            self.overall_status = QAStatus.PASS_WITH_WARNINGS
        else:
            self.overall_status = QAStatus.PASS
        return self.overall_status

    def to_dict(self) -> Dict[str, Any]:
        """Converts QA result into a serializable dictionary."""
        return {
            "job_id": self.job_id,
            "input_file": self.input_file,
            "output_file": self.output_file,
            "test_profile": self.test_profile,
            "tests_run": self.tests_run,
            "tests_passed": self.tests_passed,
            "tests_failed": self.tests_failed,
            "warnings_count": len(self.warnings),
            "warnings": self.warnings,
            "dimensions": {
                "integrity_status": self.integrity_status.value,
                "design_preservation_status": self.design_preservation_status.value,
                "original_protection_status": self.original_protection_status.value,
                "relationship_status": self.relationship_status.value,
                "accessibility_consistency_status": self.accessibility_consistency_status.value,
                "idempotency_status": self.idempotency_status.value,
                "ai_fallback_status": self.ai_fallback_status.value,
                "workspace_cleanup_status": self.workspace_cleanup_status.value,
                "report_accuracy_status": self.report_accuracy_status.value,
                "animation_preservation_status": self.animation_preservation_status.value,
            },
            "overall_status": self.overall_status.value,
            "checks": [c.to_dict() for c in self.checks],
            "details": self.details,
            "execution_time_seconds": round(self.execution_time_seconds, 3),
        }
