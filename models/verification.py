"""
SlideGuard Verification Data Model.
Captures structural and OOXML validity checks performed on the remediated PPTX.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any


@dataclass
class VerificationCheck:
    """An individual validation check result."""
    name: str
    passed: bool
    details: str
    slide_number: int = 0  # 0 indicates global check


@dataclass
class VerificationResult:
    """Overall result of post-remediation presentation validation."""
    passed: bool = True
    total_checks: int = 0
    passed_checks: int = 0
    failed_checks: int = 0
    checks: List[VerificationCheck] = field(default_factory=list)
    error_summary: str = ""

    def add_check(self, name: str, passed: bool, details: str, slide_number: int = 0):
        self.checks.append(VerificationCheck(name=name, passed=passed, details=details, slide_number=slide_number))
        self.total_checks += 1
        if passed:
            self.passed_checks += 1
        else:
            self.failed_checks += 1
            self.passed = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "passed": self.passed,
            "total_checks": self.total_checks,
            "passed_checks": self.passed_checks,
            "failed_checks": self.failed_checks,
            "checks": [
                {
                    "name": c.name,
                    "passed": c.passed,
                    "details": c.details,
                    "slide_number": c.slide_number
                }
                for c in self.checks
            ],
            "error_summary": self.error_summary
        }
