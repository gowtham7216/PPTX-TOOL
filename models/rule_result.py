"""
SlideGuard Rule Result Data Model.
Captures the output of executing an individual rule on a slide or presentation.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
from models.enums import RuleStatus
from models.issue import Issue


@dataclass
class RuleResult:
    """The outcome of evaluating a specific accessibility rule against a target."""
    rule_id: str
    rule_name: str
    slide_number: int  # 0 indicates presentation-level rule
    status: RuleStatus
    issues: List[Issue] = field(default_factory=list)
    execution_time_ms: float = 0.0
    message: str = ""
    wcag_mapping: Optional[str] = None

    @property
    def has_issues(self) -> bool:
        return len(self.issues) > 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "rule_id": self.rule_id,
            "rule_name": self.rule_name,
            "slide_number": self.slide_number,
            "status": self.status.value,
            "wcag_mapping": self.wcag_mapping,
            "issues": [issue.to_dict() for issue in self.issues],
            "execution_time_ms": self.execution_time_ms,
            "message": self.message,
        }
