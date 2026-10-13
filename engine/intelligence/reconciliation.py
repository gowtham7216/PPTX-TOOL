"""
SlideGuard Issue Reconciliation Service (Phase 8).
Compares pre-remediation and post-remediation findings to classify issues into:
  - RESOLVED: Issue successfully eliminated
  - REMEDIATED: Issue deterministically modified by an applied action
  - UNCHANGED: Issue remains identical
  - NEW: Issue introduced after remediation
  - CHANGED: Issue altered in severity or description
  - MANUAL_REVIEW: Flagged for human review
  - UNSUPPORTED: Beyond automated remediation scope
Preserves slide, shape, and rule traceability.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Any, Set

from models.unified_issue import UnifiedIssue, ReconciliationStatus
from models.issue import Issue
from models.remediation import RemediationAction
from models.enums import RemediationStatus, AccessibilityStatus


@dataclass
class ReconciledItem:
    """Individual reconciled issue tracking status across remediation."""
    unified_id: str
    slide_number: int
    shape_id: Optional[int]
    rule_id: str
    status: ReconciliationStatus
    description: str
    pre_severity: Optional[str] = None
    post_severity: Optional[str] = None
    action_taken: Optional[str] = None
    details: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "unified_id": self.unified_id,
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "rule_id": self.rule_id,
            "status": self.status.value,
            "description": self.description,
            "pre_severity": self.pre_severity,
            "post_severity": self.post_severity,
            "action_taken": self.action_taken,
            "details": self.details,
        }


@dataclass
class ReconciliationSummary:
    """Aggregated metrics and items for before/after issue reconciliation."""
    total_pre: int = 0
    total_post: int = 0
    resolved_count: int = 0
    remediated_count: int = 0
    unchanged_count: int = 0
    new_count: int = 0
    changed_count: int = 0
    manual_review_count: int = 0
    baseline_manual_review_count: int = 0
    decomposed_remnants_count: int = 0
    unsupported_count: int = 0
    raw_issues_pre: int = 0
    raw_issues_post: int = 0
    raw_issues_remediated: int = 0
    items: List[ReconciledItem] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "total_pre": self.total_pre,
            "total_post": self.total_post,
            "resolved_count": self.resolved_count,
            "remediated_count": self.remediated_count,
            "unchanged_count": self.unchanged_count,
            "new_count": self.new_count,
            "changed_count": self.changed_count,
            "manual_review_count": self.manual_review_count,
            "baseline_manual_review_count": self.baseline_manual_review_count,
            "decomposed_remnants_count": self.decomposed_remnants_count,
            "unsupported_count": self.unsupported_count,
            "raw_issues_pre": self.raw_issues_pre,
            "raw_issues_post": self.raw_issues_post,
            "raw_issues_remediated": self.raw_issues_remediated,
            "items": [it.to_dict() for it in self.items],
        }


class IssueReconciliationService:
    """Reconciles pre-remediation audit issues against post-remediation analysis."""

    RULE_FAMILY_MAP: Dict[str, Set[str]] = {
        "SG-RO-01": {"SG-RO-01", "SG-RO-VAL-01", "PIPELINE_STEP_1_READING_ORDER"},
        "SG-RO-VAL-01": {"SG-RO-01", "SG-RO-VAL-01", "PIPELINE_STEP_1_READING_ORDER"},
        "PIPELINE_STEP_1_READING_ORDER": {"SG-RO-01", "SG-RO-VAL-01", "PIPELINE_STEP_1_READING_ORDER"},
        "SG-GRP-01": {"SG-GRP-01", "PIPELINE_STEP_2_UNGROUPING"},
        "PIPELINE_STEP_2_UNGROUPING": {"SG-GRP-01", "PIPELINE_STEP_2_UNGROUPING"},
        "SG-PH-01": {"SG-PH-01", "SG-TTL-01", "PIPELINE_STEP_3_PLACEHOLDERS"},
        "PIPELINE_STEP_3_PLACEHOLDERS": {"SG-PH-01", "SG-TTL-01", "PIPELINE_STEP_3_PLACEHOLDERS"},
        "SG-LST-01": {"SG-LST-01", "PIPELINE_STEP_4_LISTS"},
        "PIPELINE_STEP_4_LISTS": {"SG-LST-01", "PIPELINE_STEP_4_LISTS"},
        "SG-TTL-01": {"SG-TTL-01", "SG-PH-01", "PIPELINE_STEP_5_ACCESSIBILITY"},
        "SG-DEC-01": {"SG-DEC-01", "PIPELINE_STEP_5_ACCESSIBILITY"},
        "SG-TBL-01": {"SG-TBL-01", "PIPELINE_STEP_5_ACCESSIBILITY"},
        "SG-LNG-01": {"SG-LNG-01", "PIPELINE_STEP_5_ACCESSIBILITY"},
        "SG-META-01": {"SG-META-01", "PIPELINE_STEP_5_ACCESSIBILITY"},
        "SG-ALT-01": {"SG-ALT-01", "PIPELINE_STEP_5_ACCESSIBILITY"},
        "PIPELINE_STEP_5_ACCESSIBILITY": {
            "SG-TTL-01", "SG-DEC-01", "SG-TBL-01", "SG-LNG-01", "SG-META-01", "SG-ALT-01"
        },
    }

    def _find_correlated_action(
        self,
        issue: UnifiedIssue,
        actions: List[RemediationAction],
    ) -> Optional[RemediationAction]:
        """Correlates an applied remediation action to a UnifiedIssue using stable slide,
        shape identifiers, and documented cross-rule relationships."""
        if not actions:
            return None

        # Collect issue target shapes
        issue_shapes: Set[int] = set()
        if issue.shape_id is not None:
            issue_shapes.add(issue.shape_id)
        for raw in issue.raw_issues:
            if getattr(raw, "shape_id", None) is not None:
                issue_shapes.add(raw.shape_id)
            if getattr(raw, "extra_data", None):
                if raw.extra_data.get("candidate_shape_id") is not None:
                    issue_shapes.add(raw.extra_data["candidate_shape_id"])
                if raw.extra_data.get("group_id") is not None:
                    issue_shapes.add(raw.extra_data["group_id"])

        # Collect issue rule IDs and family
        issue_rules: Set[str] = {issue.primary_rule_id}
        issue_rules.update(issue.related_rule_ids)
        issue_rules.update(r.rule_id for r in issue.raw_issues if getattr(r, "rule_id", None))

        rule_family: Set[str] = set()
        for r in issue_rules:
            rule_family.update(self.RULE_FAMILY_MAP.get(r, {r}))

        # Best match candidates
        exact_shape_matches = []
        slide_level_matches = []

        for action in actions:
            if action.status != RemediationStatus.APPLIED:
                continue

            # Slide match check
            slide_matches = (action.slide_number == issue.slide_number) or (action.slide_number == 0)
            if not slide_matches:
                continue

            # Action rule family
            act_rules = {action.rule_id} if action.rule_id else set()
            act_family = set().union(*(self.RULE_FAMILY_MAP.get(r, {r}) for r in act_rules))
            rule_matches = bool(act_family & issue_rules or act_rules & rule_family)

            # Action target shapes
            act_shapes: Set[int] = set()
            if action.target_shape_id is not None:
                act_shapes.add(action.target_shape_id)
            orig_st = getattr(action, "original_state", {}) or {}
            if isinstance(orig_st, dict):
                for k in ("shape_id", "group_id", "candidate_shape_id"):
                    val = orig_st.get(k)
                    if val is not None:
                        act_shapes.add(val)
            rem_st = getattr(action, "remediated_state", {}) or {}
            if isinstance(rem_st, dict):
                for sid in rem_st.get("converted_shape_ids", []):
                    act_shapes.add(sid)
                for sid in rem_st.get("extracted_shape_ids", []):
                    act_shapes.add(sid)
                for target in rem_st.get("remediated_targets", []):
                    if isinstance(target, (list, tuple)) and len(target) >= 1:
                        act_shapes.add(target[0])

            # Evaluate match quality
            if issue.shape_id is not None:
                # Shape-level issue
                if act_shapes and (act_shapes & issue_shapes):
                    if rule_matches:
                        exact_shape_matches.append(action)
                elif not act_shapes and rule_matches and action.slide_number == issue.slide_number:
                    slide_level_matches.append(action)
            else:
                # Slide-level issue (e.g. reading order, presentation language)
                if rule_matches:
                    if not act_shapes or action.slide_number == issue.slide_number:
                        slide_level_matches.append(action)

        if exact_shape_matches:
            return exact_shape_matches[0]
        if slide_level_matches:
            return slide_level_matches[0]
        return None

    def reconcile(
        self,
        pre_issues: List[UnifiedIssue],
        post_issues: List[UnifiedIssue],
        remediation_actions: Optional[List[RemediationAction]] = None,
        raw_pre_issues: Optional[List[Issue]] = None,
        raw_post_issues: Optional[List[Issue]] = None,
    ) -> ReconciliationSummary:
        """Compares pre and post unified issues and assigns reconciliation statuses."""
        remediation_actions = remediation_actions or []
        summary = ReconciliationSummary(
            total_pre=len(pre_issues),
            total_post=len(post_issues),
        )

        # Track distinct raw rule violations counts
        if raw_pre_issues is not None:
            summary.raw_issues_pre = len(raw_pre_issues)
            summary.raw_issues_post = len(raw_post_issues or [])
            from collections import Counter
            post_counts = Counter(
                (issue.slide_number, issue.shape_id, getattr(issue, "rule_id", ""))
                for issue in (raw_post_issues or [])
            )
            raw_rem = 0
            for issue in raw_pre_issues:
                k = (issue.slide_number, issue.shape_id, getattr(issue, "rule_id", ""))
                if post_counts[k] > 0:
                    post_counts[k] -= 1
                else:
                    raw_rem += 1
            summary.raw_issues_remediated = raw_rem
        elif any(getattr(u, "raw_issues", None) for u in pre_issues):
            summary.raw_issues_pre = sum(len(u.raw_issues) for u in pre_issues)
            summary.raw_issues_post = sum(len(u.raw_issues) for u in post_issues)
            summary.raw_issues_remediated = max(0, summary.raw_issues_pre - summary.raw_issues_post)

        # Index post issues by match key: (slide_number, shape_id, primary_rule_id)
        post_by_key = {}
        for post in post_issues:
            key = (post.slide_number, post.shape_id, post.primary_rule_id)
            post_by_key[key] = post

        matched_post_keys = set()

        # Step 1: Reconcile each pre-issue
        for pre in pre_issues:
            key = (pre.slide_number, pre.shape_id, pre.primary_rule_id)
            matching_post = post_by_key.get(key)

            # Check if an applied remediation targeted this issue using stable identifiers
            action = self._find_correlated_action(pre, remediation_actions)

            # Check if post has a remnant issue from pre's cluster on same shape/slide
            remnant_post = None
            if matching_post is None:
                for rel_rule in pre.related_rule_ids:
                    rel_key = (pre.slide_number, pre.shape_id, rel_rule)
                    if rel_key in post_by_key:
                        remnant_post = post_by_key[rel_key]
                        matched_post_keys.add(rel_key)
                        break

            if matching_post is None:
                # Issue no longer present post-remediation!
                rec_status = ReconciliationStatus.REMEDIATED
                summary.remediated_count += 1
                summary.resolved_count += 1

                if action:
                    if remnant_post:
                        summary.decomposed_remnants_count += 1
                        details = (
                            f"Safely remediated primary issue by action: {action.description}; "
                            f"related {remnant_post.primary_rule_id} preserved for manual review."
                        )
                        remnant_post.reconciliation_status = ReconciliationStatus.MANUAL_REVIEW
                        remnant_post.verification_status = "VERIFIED"
                    else:
                        details = f"Safely remediated by action: {action.description}"
                else:
                    if remnant_post:
                        summary.decomposed_remnants_count += 1
                        details = (
                            f"Safely resolved primary issue post-remediation; "
                            f"related {remnant_post.primary_rule_id} preserved for manual review."
                        )
                        remnant_post.reconciliation_status = ReconciliationStatus.MANUAL_REVIEW
                        remnant_post.verification_status = "VERIFIED"
                    else:
                        details = "Issue verified resolved post-remediation."

                item = ReconciledItem(
                    unified_id=pre.unified_id,
                    slide_number=pre.slide_number,
                    shape_id=pre.shape_id,
                    rule_id=pre.primary_rule_id,
                    status=rec_status,
                    description=pre.explanation or pre.recommended_action,
                    pre_severity=pre.severity.value,
                    post_severity=None,
                    action_taken=action.description if action else None,
                    details=details,
                )
                pre.reconciliation_status = rec_status
                pre.verification_status = "VERIFIED"
                summary.items.append(item)
            else:
                # Issue still exists post-remediation
                matched_post_keys.add(key)
                if pre.status == AccessibilityStatus.UNSUPPORTED or matching_post.status == AccessibilityStatus.UNSUPPORTED:
                    rec_status = ReconciliationStatus.UNSUPPORTED
                    summary.unsupported_count += 1
                    details = "Unsupported structural issue; preserved without alteration."
                elif pre.status == AccessibilityStatus.MANUAL_REVIEW or matching_post.status == AccessibilityStatus.MANUAL_REVIEW:
                    rec_status = ReconciliationStatus.MANUAL_REVIEW
                    summary.baseline_manual_review_count += 1
                    details = "Requires human manual authoring or verification."
                elif pre.severity != matching_post.severity or pre.explanation != matching_post.explanation:
                    rec_status = ReconciliationStatus.CHANGED
                    summary.changed_count += 1
                    details = f"Severity changed from {pre.severity.value} to {matching_post.severity.value}."
                else:
                    rec_status = ReconciliationStatus.UNCHANGED
                    summary.unchanged_count += 1
                    details = "Issue unchanged after remediation."

                item = ReconciledItem(
                    unified_id=pre.unified_id,
                    slide_number=pre.slide_number,
                    shape_id=pre.shape_id,
                    rule_id=pre.primary_rule_id,
                    status=rec_status,
                    description=pre.explanation or pre.recommended_action,
                    pre_severity=pre.severity.value,
                    post_severity=matching_post.severity.value,
                    action_taken=action.description if action else None,
                    details=details,
                )
                pre.reconciliation_status = rec_status
                pre.verification_status = "VERIFIED"
                matching_post.reconciliation_status = rec_status
                matching_post.verification_status = "VERIFIED"
                summary.items.append(item)

        # Step 2: Check for truly NEW issues in post_issues
        for post in post_issues:
            key = (post.slide_number, post.shape_id, post.primary_rule_id)
            if key not in matched_post_keys:
                rec_status = ReconciliationStatus.NEW
                summary.new_count += 1
                item = ReconciledItem(
                    unified_id=post.unified_id,
                    slide_number=post.slide_number,
                    shape_id=post.shape_id,
                    rule_id=post.primary_rule_id,
                    status=rec_status,
                    description=post.explanation or post.recommended_action,
                    pre_severity=None,
                    post_severity=post.severity.value,
                    details="New finding detected post-remediation.",
                )
                post.reconciliation_status = rec_status
                post.verification_status = "VERIFIED"
                summary.items.append(item)

        # Aggregate candidate manual review count (baseline manual review + decomposed remnants)
        summary.manual_review_count = summary.baseline_manual_review_count + summary.decomposed_remnants_count

        return summary
