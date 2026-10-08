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
from typing import List, Dict, Tuple, Optional, Any

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
    unsupported_count: int = 0
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
            "unsupported_count": self.unsupported_count,
            "items": [it.to_dict() for it in self.items],
        }


class IssueReconciliationService:
    """Reconciles pre-remediation audit issues against post-remediation analysis."""

    def reconcile(
        self,
        pre_issues: List[UnifiedIssue],
        post_issues: List[UnifiedIssue],
        remediation_actions: Optional[List[RemediationAction]] = None,
    ) -> ReconciliationSummary:
        """Compares pre and post unified issues and assigns reconciliation statuses."""
        remediation_actions = remediation_actions or []
        summary = ReconciliationSummary(
            total_pre=len(pre_issues),
            total_post=len(post_issues),
        )

        # Index applied actions by (slide_number, target_shape_id) and by rule_id
        applied_by_target = {}
        for action in remediation_actions:
            if action.status == RemediationStatus.APPLIED:
                key = (action.slide_number, action.target_shape_id)
                applied_by_target[key] = action
                if action.rule_id:
                    applied_by_target[(action.slide_number, action.rule_id)] = action

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

            # Check if an applied remediation targeted this issue
            action = applied_by_target.get((pre.slide_number, pre.shape_id)) or applied_by_target.get((pre.slide_number, pre.primary_rule_id))

            if matching_post is None:
                # Issue no longer present post-remediation!
                if action:
                    rec_status = ReconciliationStatus.REMEDIATED
                    summary.remediated_count += 1
                    summary.resolved_count += 1
                    details = f"Safely remediated by action: {action.description}"
                else:
                    rec_status = ReconciliationStatus.RESOLVED
                    summary.resolved_count += 1
                    details = "Issue resolved post-remediation."

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
                    summary.manual_review_count += 1
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
                matching_post.reconciliation_status = rec_status
                summary.items.append(item)

        # Step 2: Check for NEW issues in post_issues
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
                summary.items.append(item)

        return summary
