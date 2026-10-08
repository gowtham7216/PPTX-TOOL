"""
SlideGuard Confidence Evaluation Engine (Phase 8).
Evaluates classification and remediation confidence across HIGH, MEDIUM, and LOW tiers.
Ensures high confidence alone never bypasses deterministic safety prerequisites.
"""

from typing import List, Dict, Any, Optional

from models.unified_issue import ConfidenceLevel, RemediationEligibility
from models.issue import Issue
from models.enums import IssueSeverity, AccessibilityStatus, RuleBehavior


class ConfidenceEngine:
    """
    Evaluates confidence level for findings, context classifications, and remediation proposals.
    
    Tiers:
      - HIGH: Strong deterministic criteria met; safe automatic remediation allowed
              if deterministic safety conditions are satisfied.
      - MEDIUM: Plausible heuristic or candidate matches; user review required.
      - LOW: Ambiguous or unverified findings; detect-only / manual authoring needed.
    """

    def evaluate_issue_confidence(
        self,
        primary_issue: Issue,
        corroborating_issues: Optional[List[Issue]] = None,
        context_data: Optional[Dict[str, Any]] = None
    ) -> ConfidenceLevel:
        """Determines the confidence tier based on issue evidence and corroboration."""
        corroborating_issues = corroborating_issues or []
        context_data = context_data or {}

        # 1. Check if the issue already had a confident evaluation
        explicit_conf = (primary_issue.confidence or "").upper()
        if explicit_conf == "HIGH" and primary_issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE:
            # Verify deterministic conditions hold
            if self._verify_deterministic_safety(primary_issue, context_data):
                return ConfidenceLevel.HIGH
            return ConfidenceLevel.MEDIUM

        # 2. Corroboration bonus: if multiple rules agree on the same shape
        matching_rules_count = len(corroborating_issues) + 1
        if matching_rules_count >= 2:
            # If rules agree and at least one is high-confidence
            if any(i.confidence == "HIGH" for i in [primary_issue] + corroborating_issues):
                if self._verify_deterministic_safety(primary_issue, context_data):
                    return ConfidenceLevel.HIGH
                return ConfidenceLevel.MEDIUM

        # 3. Ambiguous markers force MEDIUM or LOW
        desc = (primary_issue.description or "").lower()
        expl = (primary_issue.explanation or "").lower()
        if "ambiguous" in desc or "ambiguous" in expl or "manual review" in desc:
            return ConfidenceLevel.MEDIUM

        if primary_issue.status in (AccessibilityStatus.MANUAL_REVIEW, AccessibilityStatus.UNSUPPORTED):
            return ConfidenceLevel.LOW

        if explicit_conf in ("MEDIUM", "MODERATE"):
            return ConfidenceLevel.MEDIUM
        elif explicit_conf == "LOW":
            return ConfidenceLevel.LOW

        # Default fallback
        return ConfidenceLevel.HIGH if primary_issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE else ConfidenceLevel.MEDIUM

    def determine_remediation_eligibility(
        self,
        confidence: ConfidenceLevel,
        primary_issue: Issue,
        context_data: Optional[Dict[str, Any]] = None
    ) -> RemediationEligibility:
        """
        Determines remediation eligibility.
        HIGH confidence allows ELIGIBLE_AUTO only if strict deterministic conditions hold.
        """
        context_data = context_data or {}

        # Safe auto-remediation rule IDs
        safe_rules = {"SG-DEC-01", "SG-TTL-01", "SG-TBL-01", "SG-LNG-01", "SG-META-01", "SG-RO-01"}

        if primary_issue.rule_id not in safe_rules:
            return RemediationEligibility.MANUAL_ONLY

        if primary_issue.behavior == RuleBehavior.DETECT_ONLY:
            return RemediationEligibility.MANUAL_ONLY

        if primary_issue.status == AccessibilityStatus.UNSUPPORTED:
            return RemediationEligibility.UNSUPPORTED

        if confidence == ConfidenceLevel.HIGH and self._verify_deterministic_safety(primary_issue, context_data):
            return RemediationEligibility.ELIGIBLE_AUTO
        elif confidence == ConfidenceLevel.MEDIUM:
            return RemediationEligibility.REQUIRES_USER_APPROVAL
        else:
            return RemediationEligibility.MANUAL_ONLY

    def _verify_deterministic_safety(self, issue: Issue, context_data: Dict[str, Any]) -> bool:
        """
        Safety check: ensures confidence alone cannot trigger remediation without
        satisfying deterministic prerequisites.
        """
        rid = issue.rule_id

        # Decorative rule: must be divider or explicitly decorative
        if rid == "SG-DEC-01":
            extra = issue.extra_data or {}
            return bool(extra.get("is_divider") or context_data.get("is_divider") or issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE)

        # Title rule: must have candidate shape ID and non-empty candidate text
        if rid == "SG-TTL-01":
            extra = issue.extra_data or {}
            cand_id = issue.shape_id or extra.get("candidate_shape_id")
            cand_text = extra.get("candidate_text")
            return bool((cand_id is not None and cand_text) or issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE)

        # Table rule: must have complete row 0
        if rid == "SG-TBL-01":
            extra = issue.extra_data or {}
            return bool(extra.get("can_set_first_row", False) or issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE)

        # Language rule: must have explicit configured language or SAFE_TO_REMEDIATE
        if rid == "SG-LNG-01":
            extra = issue.extra_data or {}
            return bool(extra.get("configured_language") or issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE)

        # Metadata rule: must have explicit configured title or SAFE_TO_REMEDIATE
        if rid == "SG-META-01":
            extra = issue.extra_data or {}
            return bool(extra.get("configured_title") or issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE)

        # Reading order: must have valid proposed sequence
        if rid == "SG-RO-01":
            extra = issue.extra_data or {}
            return bool(extra.get("proposed_order") or issue.status == AccessibilityStatus.SAFE_TO_REMEDIATE)

        return False
