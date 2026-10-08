"""
SlideGuard Cross-Rule Correlator & Issue Deduplicator (Phase 8).
Correlates related issues across distinct rules into unified semantic decisions
and eliminates duplicate/overlapping findings using contextual identity keys.
"""

from typing import List, Dict, Tuple, Optional, Any, Set
from collections import defaultdict

from models.issue import Issue
from models.rule_result import RuleResult
from models.unified_issue import (
    UnifiedIssue,
    ConfidenceLevel,
    RemediationEligibility,
    ContextCategory,
)
from models.enums import IssueSeverity, AccessibilityStatus, RuleBehavior
from engine.intelligence.confidence_engine import ConfidenceEngine
from engine.intelligence.context_classifier import ContextualClassifier


class IssueCorrelator:
    """
    Correlates findings from multiple independent rules into unified semantic issues.
    Preserves all original rule evidence without loss.
    """

    def __init__(
        self,
        confidence_engine: Optional[ConfidenceEngine] = None,
        context_classifier: Optional[ContextualClassifier] = None,
    ):
        self.confidence_engine = confidence_engine or ConfidenceEngine()
        self.context_classifier = context_classifier or ContextualClassifier()

    def correlate(
        self,
        rule_results: List[RuleResult],
        slide_models: Optional[Dict[int, Any]] = None,
    ) -> List[UnifiedIssue]:
        """
        Processes all rule results across the presentation, performing:
          1. Contextual deduplication
          2. Cross-rule correlation by shape, slide, and semantic role
          3. Confidence evaluation and remediation eligibility determination
        """
        # Step 1: Collect all raw issues
        raw_issues: List[Issue] = []
        for res in rule_results:
            raw_issues.extend(res.issues)

        if not raw_issues:
            return []

        # Step 2: Contextual deduplication
        deduplicated_issues = self._deduplicate_issues(raw_issues)

        # Step 3: Group issues by slide and shape
        grouped_by_target = defaultdict(list)
        for issue in deduplicated_issues:
            target_key = (issue.slide_number, issue.shape_id)
            grouped_by_target[target_key].append(issue)

        raw_by_target = defaultdict(list)
        for issue in raw_issues:
            target_key = (issue.slide_number, issue.shape_id)
            raw_by_target[target_key].append(issue)

        # Step 4: Synthesize unified issues
        unified_issues: List[UnifiedIssue] = []
        for (slide_num, shape_id), issues in grouped_by_target.items():
            all_raw = raw_by_target[(slide_num, shape_id)]
            unified = self._create_unified_issue(slide_num, shape_id, issues, all_raw_issues=all_raw)
            unified_issues.append(unified)

        # Sort cleanly by slide number then shape_id
        unified_issues.sort(key=lambda u: (u.slide_number, u.shape_id if u.shape_id is not None else -1))
        return unified_issues

    def _deduplicate_issues(self, issues: List[Issue]) -> List[Issue]:
        """
        Eliminates duplicate or overlapping issues using contextual identity:
        (slide_number, shape_id, rule_id, issue_type_normalized)
        """
        seen: Dict[Tuple[int, Optional[int], str, str], Issue] = {}
        for issue in issues:
            # Derive normalized issue type key
            issue_type = self._derive_issue_type(issue)
            key = (issue.slide_number, issue.shape_id, issue.rule_id, issue_type)

            if key not in seen:
                seen[key] = issue
            else:
                # Merge extra data into existing issue to preserve evidence
                existing = seen[key]
                if issue.extra_data and existing.extra_data is not None:
                    existing.extra_data.update(issue.extra_data)

        return list(seen.values())

    def _derive_issue_type(self, issue: Issue) -> str:
        """Derives a semantic issue type key from issue ID and description."""
        iid = (issue.issue_id or "").upper()
        if "TTL" in iid or "TITLE" in iid:
            return "TITLE_SEMANTICS"
        elif "DEC" in iid or "DIVIDER" in iid:
            return "DECORATIVE_TAGGING"
        elif "ALT" in iid:
            return "ALT_TEXT_DEFICIENCY"
        elif "TBL" in iid or "TABLE" in iid:
            return "TABLE_STRUCTURE"
        elif "CNT" in iid or "CONTRAST" in iid:
            return "COLOR_CONTRAST"
        elif "RDB" in iid or "READABILITY" in iid:
            return "READABILITY_FONT"
        elif "TXT" in iid or "EMPTY" in iid:
            return "TEXT_EMPTINESS"
        elif "LNG" in iid or "LANG" in iid:
            return "DOCUMENT_LANGUAGE"
        elif "META" in iid:
            return "DOCUMENT_METADATA"
        elif "RO" in iid or "READING_ORDER" in iid:
            return "READING_ORDER_SEQUENCE"
        elif "GRP" in iid:
            return "GROUP_STRUCTURE"
        elif "LST" in iid:
            return "LIST_FORMATTING"
        elif "PH" in iid:
            return "PLACEHOLDER_STRUCTURE"
        return issue.rule_id

    def _create_unified_issue(
        self,
        slide_number: int,
        shape_id: Optional[int],
        issues: List[Issue],
        all_raw_issues: Optional[List[Issue]] = None,
    ) -> UnifiedIssue:
        """Combines multiple correlated issues on the same target into a single UnifiedIssue."""
        # 1. Determine primary issue (highest severity and most specific rule)
        severity_order = {
            IssueSeverity.CRITICAL: 4,
            IssueSeverity.ERROR: 3,
            IssueSeverity.WARNING: 2,
            IssueSeverity.INFO: 1,
        }
        sorted_issues = sorted(
            issues,
            key=lambda x: (
                severity_order.get(x.severity, 0),
                1 if x.status == AccessibilityStatus.SAFE_TO_REMEDIATE else 0,
            ),
            reverse=True,
        )
        primary = sorted_issues[0]
        corroborating = sorted_issues[1:]

        # 2. Collect all rule IDs
        all_rule_ids = set()
        for i in issues:
            all_rule_ids.add(i.rule_id)
        related_rule_ids = [r for r in all_rule_ids if r != primary.rule_id]

        # 3. Categorize context
        category = self._categorize_context(primary, issues)

        # 4. Synthesize unified explanation and action
        explanation = primary.explanation or primary.description
        if corroborating:
            corr_names = ", ".join(sorted(set(i.rule_name for i in corroborating)))
            explanation = f"{explanation} Corroborated by related checks: {corr_names}."

        recommended_action = primary.recommendation

        # 5. Evaluate confidence & remediation eligibility
        merged_extra: Dict[str, Any] = {}
        for i in issues:
            if i.extra_data:
                merged_extra.update(i.extra_data)

        confidence = self.confidence_engine.evaluate_issue_confidence(
            primary_issue=primary,
            corroborating_issues=corroborating,
            context_data=merged_extra,
        )

        eligibility = self.confidence_engine.determine_remediation_eligibility(
            confidence=confidence,
            primary_issue=primary,
            context_data=merged_extra,
        )

        # Determine unified status
        status = primary.status
        if any(i.status == AccessibilityStatus.SAFE_TO_REMEDIATE for i in issues) and eligibility == RemediationEligibility.ELIGIBLE_AUTO:
            status = AccessibilityStatus.SAFE_TO_REMEDIATE

        shape_name = primary.shape_name or (corroborating[0].shape_name if corroborating else None)

        unified_id = f"UI_{slide_number}_{shape_id if shape_id is not None else 'DOC'}_{primary.rule_id}"

        return UnifiedIssue(
            unified_id=unified_id,
            slide_number=slide_number,
            shape_id=shape_id,
            shape_name=shape_name,
            primary_rule_id=primary.rule_id,
            primary_rule_name=primary.rule_name,
            related_rule_ids=related_rule_ids,
            severity=primary.severity,
            confidence=confidence,
            explanation=explanation,
            recommended_action=recommended_action,
            remediation_eligibility=eligibility,
            semantic_role=merged_extra.get("semantic_role"),
            context_category=category.value,
            raw_issues=all_raw_issues if all_raw_issues is not None else issues,  # All original evidence preserved intact
            status=status,
            verification_status=primary.verification_status,
            extra_data=merged_extra,
        )

    def _categorize_context(self, primary: Issue, issues: List[Issue]) -> ContextCategory:
        """Determines the overarching ContextCategory for the correlated issue."""
        rids = {i.rule_id for i in issues}

        if any(r.startswith("SG-TTL") or "TTL" in r for r in rids):
            return ContextCategory.TITLE
        elif any(r.startswith("SG-ALT") or r.startswith("SG-DEC") for r in rids):
            return ContextCategory.IMAGE
        elif any(r.startswith("SG-TBL") for r in rids):
            return ContextCategory.TABLE
        elif any(r.startswith("SG-CNT") or r.startswith("SG-RDB") or r.startswith("SG-TXT") for r in rids):
            return ContextCategory.TEXT
        elif any(r.startswith("SG-RO") for r in rids):
            return ContextCategory.READING_ORDER
        elif any(r.startswith("SG-LNG") or r.startswith("SG-META") for r in rids):
            return ContextCategory.DOCUMENT
        elif any(r.startswith("SG-GRP") for r in rids):
            return ContextCategory.GROUP
        elif any(r.startswith("SG-LST") for r in rids):
            return ContextCategory.LIST
        return ContextCategory.OTHER
