"""
Regression and unit tests for Chapter 7 audit findings and remediation fixes:
1. Reading Order per-slide status tracking and candidate XML verification.
2. Placeholder conversion candidate detection and post-remediation audit metrics.
3. Accessibility issue reconciliation without false 'NEW' issues.
4. Scoring consistency between Section 1 and Section 7.
"""

import pytest
from pathlib import Path
from unittest.mock import MagicMock

from models.job import JobContext, JobStage, QualityMetrics
from models.presentation import PresentationModel, SlideModel
from models.shape import ShapeModel
from models.issue import Issue
from models.rule_result import RuleResult
from models.unified_issue import UnifiedIssue, ReconciliationStatus
from models.remediation import RemediationAction
from models.enums import IssueSeverity, RuleBehavior, AccessibilityStatus, RemediationStatus, RuleStatus
from engine.intelligence.reconciliation import IssueReconciliationService
from engine.remediation.remediation_service import RemediationService
from reports.generator import ReportGenerator


def test_reconciliation_eliminates_false_new_issues_from_decomposed_clusters():
    """Verifies that pre-existing contrast/typography issues in composite clusters
    are not falsely marked as 'NEW' issues post-remediation."""
    reconciler = IssueReconciliationService()

    # Pre-remediation: composite cluster with primary SG-TTL-01 and related SG-CNT-01
    pre_issue = UnifiedIssue(
        unified_id="UI_slide1_SG-TTL-01",
        slide_number=1,
        shape_id=2,
        primary_rule_id="SG-TTL-01",
        related_rule_ids=["SG-CNT-01", "SG-PH-01"],
        severity=IssueSeverity.WARNING,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        remediation_eligibility=AccessibilityStatus.SAFE_TO_REMEDIATE,
        explanation="Missing title placeholder",
        recommended_action="Promote shape to title placeholder",
    )

    # Remediation action applied to slide 1 shape 2
    action = RemediationAction(
        action_id="act_title_1",
        rule_id="SG-TTL-01",
        slide_number=1,
        target_shape_id=2,
        description="Promoted shape 2 to title",
        status=RemediationStatus.APPLIED,
        confidence=1.0,
    )

    # Post-remediation: Title and placeholder resolved, but contrast remains on shape 2
    post_issue = UnifiedIssue(
        unified_id="UI_slide1_SG-CNT-01",
        slide_number=1,
        shape_id=2,
        primary_rule_id="SG-CNT-01",
        related_rule_ids=[],
        severity=IssueSeverity.WARNING,
        status=AccessibilityStatus.MANUAL_REVIEW,
        remediation_eligibility=AccessibilityStatus.MANUAL_REVIEW,
        explanation="Insufficient color contrast",
        recommended_action="Adjust text or fill color",
    )

    summary = reconciler.reconcile(
        pre_issues=[pre_issue],
        post_issues=[post_issue],
        remediation_actions=[action]
    )

    # Must NOT be marked as NEW issue
    assert summary.new_count == 0
    assert summary.remediated_count == 1
    assert summary.manual_review_count == 1
    assert pre_issue.reconciliation_status == ReconciliationStatus.REMEDIATED
    assert pre_issue.verification_status == "VERIFIED"
    assert post_issue.reconciliation_status == ReconciliationStatus.MANUAL_REVIEW
    assert post_issue.verification_status == "VERIFIED"


def test_reading_order_report_includes_status_and_verification(tmp_path):
    """Verifies that Section 2 of the report displays per-slide status and XML verification."""
    context = JobContext(
        job_id="test_ro_job",
        original_file=Path("test.pptx"),
        working_file=tmp_path / "test.pptx",
        workspace_dir=tmp_path,
        total_slides=2,
        slides_analyzed=2,
        stage=JobStage.COMPLETED,
    )
    context.metrics = QualityMetrics(accessibility_score=70.0, quality_score_v2=72.0)

    # Pre-remediation SG-RO-01 issues
    issue1 = Issue(
        issue_id="ro_1",
        slide_number=1,
        shape_id=None,
        rule_id="SG-RO-01",
        rule_name="Reading Order",
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.REMEDIATED,
        verification_status="VERIFIED",
        description="Reading order non-sequential",
        recommendation="Reorder shapes according to visual reading flow",
        extra_data={"original_order": [1, 2], "proposed_order": [2, 1], "confidence": "HIGH"}
    )
    issue2 = Issue(
        issue_id="ro_2",
        slide_number=2,
        shape_id=None,
        rule_id="SG-RO-01",
        rule_name="Reading Order",
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        verification_status="PENDING",
        description="Reading order optimal",
        recommendation="Preserve reading order",
        extra_data={"original_order": [1, 2], "proposed_order": [1, 2], "confidence": "HIGH"}
    )
    context.rule_results = [
        RuleResult(
            rule_id="SG-RO-01",
            rule_name="Reading Order",
            slide_number=0,
            status=RuleStatus.FAILED,
            issues=[issue1, issue2]
        )
    ]

    # Remediation actions
    action1 = RemediationAction(
        action_id="act_ro_1",
        rule_id="SG-RO-01",
        slide_number=1,
        description="Reordered shapes [1, 2] -> [2, 1]",
        status=RemediationStatus.APPLIED,
        confidence=0.95,
    )
    action2 = RemediationAction(
        action_id="act_ro_2",
        rule_id="SG-RO-01",
        slide_number=2,
        description="Sequence already optimal",
        status=RemediationStatus.NO_CHANGES,
        confidence=1.0,
    )
    context.remediation_actions = [action1, action2]

    generator = ReportGenerator()
    out_file = tmp_path / "report.md"
    generator.generate_markdown_report(context, out_file)

    content = out_file.read_text(encoding="utf-8")
    assert "## 2. Reading Order Analysis & Proposal" in content
    assert "Reading Order Remediation Summary" in content
    assert "| **Applied in Candidate XML** | 1 |" in content
    assert "| **No Changes Required** | 1 |" in content
    assert "| 1 | **HIGH** | `[1, 2]` | `[2, 1]` | None | `APPLIED` | ✅ Verified in Candidate XML |" in content
    assert "| 2 | **HIGH** | `[1, 2]` | `[1, 2]` | None | `NO_CHANGES` | ✅ Optimal Sequence |" in content


def test_placeholder_report_metrics_accuracy(tmp_path):
    """Verifies that Section 4 displays non-zero proposed and applied conversions when audit is populated."""
    context = JobContext(
        job_id="test_ph_job",
        original_file=Path("test.pptx"),
        working_file=tmp_path / "test.pptx",
        workspace_dir=tmp_path,
        total_slides=35,
        slides_analyzed=35,
        stage=JobStage.COMPLETED,
    )
    context.metrics = QualityMetrics(accessibility_score=75.0, quality_score_v2=78.0)
    context.task_4_placeholder_report = {
        "total_slides_evaluated": 35,
        "slides_changed": 35,
        "slides_unchanged": 0,
        "candidates_detected": 625,
        "conversions_proposed": 35,
        "conversions_approved": 35,
        "conversions_applied": 35,
        "conversions_verified": 35,
        "existing_objects_converted": 35,
        "native_placeholders_preserved": 0,
        "visual_integrity_verified": True,
        "objects": [
            {
                "slide_number": 1,
                "shape_id": 7,
                "shape_name": "Title Box",
                "semantic_role": "TITLE",
                "action": "CONVERT_EXISTING",
                "proposed_placeholder_type": "title",
                "confidence": "HIGH",
                "reason": "Promote visual title to native placeholder",
            }
        ]
    }

    generator = ReportGenerator()
    out_file = tmp_path / "report.md"
    generator.generate_markdown_report(context, out_file)

    content = out_file.read_text(encoding="utf-8")
    assert "## 4. Placeholder Conversion & Analysis" in content
    assert "| **Candidates Detected** | 625 |" in content
    assert "| **Conversions Proposed** | 35 |" in content
    assert "| **Conversions Applied** | 35 |" in content
    assert "| **Conversions Verified** | 35 |" in content
    assert "| **Existing Objects Converted** | 35 |" in content
    assert "| 1 | `7` | Title Box | `TITLE` | **CONVERT_EXISTING** | `title` |" in content


def test_score_consistency_between_section1_and_section7(tmp_path):
    """Verifies that authoritative quality scores, unique remediated issues, and raw check counts
    are consistent and documented across Section 1 and Section 7."""
    context = JobContext(
        job_id="test_score_job",
        original_file=Path("Chapter 7.pptx"),
        working_file=tmp_path / "Chapter 7.pptx",
        workspace_dir=tmp_path,
        total_slides=35,
        slides_analyzed=35,
        stage=JobStage.COMPLETED,
    )
    context.metrics = QualityMetrics(
        accessibility_score=70.5,
        quality_score_v2=71.8,
        total_issues=486,
        error_issues=10,
        warning_issues=476,
        remediated_issues=70,
        unresolved_issues=244,
        raw_remediated_issues=141,
        correlated_issues_count=278,
    )
    context.pre_remediation_score = 71.8
    context.post_remediation_score = 80.8
    context.pre_accessibility_score = 70.5
    context.post_accessibility_score = 78.7
    context.post_metrics = QualityMetrics(
        accessibility_score=78.7,
        quality_score_v2=80.8,
        total_issues=345,
        error_issues=9,
        warning_issues=336,
    )
    context.intelligence_summary = {
        "total_pre": 278,
        "total_post": 244,
        "remediated_count": 70,
        "unchanged_count": 8,
        "changed_count": 1,
        "manual_review_count": 235,
        "baseline_manual_review_count": 199,
        "decomposed_remnants_count": 36,
        "unsupported_count": 0,
        "new_count": 0,
        "raw_issues_pre": 486,
        "raw_issues_post": 345,
        "raw_issues_remediated": 141,
    }

    generator = ReportGenerator()
    out_file = tmp_path / "report.md"
    generator.generate_markdown_report(context, out_file)

    content = out_file.read_text(encoding="utf-8")

    # Section 1 verification
    assert "## 1. Executive Summary & Quality Metrics" in content
    assert "| **Authoritative Quality Score (Multi-Signal V2)** | **71.8 / 100** | **80.8 / 100** |" in content
    assert "| **Raw Severity Accessibility Score (V1)** | 70.5 / 100 | 78.7 / 100 |" in content
    assert "| **Total Unique Issues** | 278 | 244 |" in content
    assert "| **Remediated & Verified Issues** | - | **70** |" in content
    assert "| **Preserved / Manual Review Issues** | - | 244 |" in content
    assert "| **Raw Rule Check Violations** | 486 | 345 |" in content
    assert "| - Remediated Raw Rule Checks | - | **141** |" in content

    # Section 7 verification
    assert "## 7. Accessibility Intelligence Summary" in content
    assert "| **Remediated Issues** | 70 |" in content
    assert "| **Unresolved Issues** | 244 |" in content
    assert "| **Raw Rule Checks Remediated** | 141 |" in content
    assert "| **Authoritative Quality Score (Before)** | 71.8 / 100 |" in content
    assert "| **Authoritative Quality Score (After)** | 80.8 / 100 |" in content
    assert "| **Raw Severity Accessibility Score (Before)** | 70.5 / 100 |" in content
    assert "| **Raw Severity Accessibility Score (After)** | 78.7 / 100 |" in content


def test_unique_issue_deduplication_not_raw_rule_counts():
    """Verifies that multiple rules referring to the same shape/slide target
    correlate into a single unique issue rather than inflating the unique remediated count."""
    from engine.intelligence.correlator import IssueCorrelator

    # Slide 1 has reading order flagged by both SG-RO-01 and SG-RO-VAL-01
    ro_issue_1 = Issue(
        issue_id="ro_1",
        slide_number=1,
        shape_id=None,
        rule_id="SG-RO-01",
        rule_name="Reading Order Sequence",
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Reading order non-sequential",
        recommendation="Reorder shapes according to visual reading flow",
    )
    ro_issue_2 = Issue(
        issue_id="ro_val_1",
        slide_number=1,
        shape_id=None,
        rule_id="SG-RO-VAL-01",
        rule_name="Reading Order Validation",
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Reading order validation failed",
        recommendation="Verify reading sequence in candidate",
    )

    # Slide 1 Shape 7 has title issue flagged by both SG-TTL-01 and SG-PH-01
    ttl_issue = Issue(
        issue_id="ttl_1",
        slide_number=1,
        shape_id=7,
        rule_id="SG-TTL-01",
        rule_name="Missing Title Placeholder",
        severity=IssueSeverity.ERROR,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Missing title",
        recommendation="Promote text box to title placeholder",
    )
    ph_issue = Issue(
        issue_id="ph_1",
        slide_number=1,
        shape_id=7,
        rule_id="SG-PH-01",
        rule_name="Native Placeholder Usage",
        severity=IssueSeverity.WARNING,
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Not native placeholder",
        recommendation="Convert to native placeholder",
    )

    rule_results = [
        RuleResult(rule_id="SG-RO-01", rule_name="RO", slide_number=0, status=RuleStatus.FAILED, issues=[ro_issue_1]),
        RuleResult(rule_id="SG-RO-VAL-01", rule_name="RO-VAL", slide_number=0, status=RuleStatus.FAILED, issues=[ro_issue_2]),
        RuleResult(rule_id="SG-TTL-01", rule_name="TTL", slide_number=0, status=RuleStatus.FAILED, issues=[ttl_issue]),
        RuleResult(rule_id="SG-PH-01", rule_name="PH", slide_number=0, status=RuleStatus.FAILED, issues=[ph_issue]),
    ]

    correlator = IssueCorrelator()
    unified_issues = correlator.correlate(rule_results)

    # 4 raw rule check issues must correlate into exactly 2 unique issues
    assert len(rule_results[0].issues) + len(rule_results[1].issues) + len(rule_results[2].issues) + len(rule_results[3].issues) == 4
    assert len(unified_issues) == 2, f"Expected 2 unique issues, got {len(unified_issues)}"

    # Reconcile when both are remediated
    actions = [
        RemediationAction(
            action_id="act_ro_1",
            rule_id="SG-RO-01",
            slide_number=1,
            description="Reordered slide 1",
            status=RemediationStatus.APPLIED,
            confidence=1.0,
        ),
        RemediationAction(
            action_id="act_ttl_1",
            rule_id="SG-TTL-01",
            slide_number=1,
            target_shape_id=7,
            description="Promoted shape 7 to title placeholder",
            status=RemediationStatus.APPLIED,
            confidence=1.0,
        ),
    ]

    reconciler = IssueReconciliationService()
    summary = reconciler.reconcile(pre_issues=unified_issues, post_issues=[], remediation_actions=actions)

    assert summary.remediated_count == 2
    assert summary.total_pre == 2
    assert summary.total_post == 0
    assert len(summary.items) == 2


def test_reconciliation_category_mutual_consistency():
    """Verifies that Remediated, Unchanged, Changed, Manual Review, and New form
    mutually consistent partitions without duplicate items or hardcoded totals."""
    reconciler = IssueReconciliationService()

    # Pre issues: 1 remediated, 1 unchanged, 1 changed, 1 baseline manual review
    pre_rem = UnifiedIssue(
        unified_id="u_rem", slide_number=1, shape_id=10, primary_rule_id="SG-TTL-01",
        related_rule_ids=["SG-CNT-01"], severity=IssueSeverity.WARNING,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE, explanation="Title issue"
    )
    pre_unc = UnifiedIssue(
        unified_id="u_unc", slide_number=2, shape_id=20, primary_rule_id="SG-ALT-01",
        severity=IssueSeverity.ERROR, status=AccessibilityStatus.SAFE_TO_REMEDIATE, explanation="Missing alt"
    )
    pre_chg = UnifiedIssue(
        unified_id="u_chg", slide_number=3, shape_id=30, primary_rule_id="SG-RDB-01",
        severity=IssueSeverity.WARNING, status=AccessibilityStatus.SAFE_TO_REMEDIATE, explanation="Readability borderline"
    )
    pre_man = UnifiedIssue(
        unified_id="u_man", slide_number=4, shape_id=40, primary_rule_id="SG-CNT-01",
        severity=IssueSeverity.WARNING, status=AccessibilityStatus.MANUAL_REVIEW, explanation="Low contrast"
    )
    pre_issues = [pre_rem, pre_unc, pre_chg, pre_man]

    # Post issues:
    # 1. remnant from u_rem (contrast remains on shape 10)
    post_remnant = UnifiedIssue(
        unified_id="u_remnant", slide_number=1, shape_id=10, primary_rule_id="SG-CNT-01",
        severity=IssueSeverity.WARNING, status=AccessibilityStatus.MANUAL_REVIEW, explanation="Contrast still low"
    )
    # 2. post for u_unc (identical)
    post_unc = UnifiedIssue(
        unified_id="u_unc_post", slide_number=2, shape_id=20, primary_rule_id="SG-ALT-01",
        severity=IssueSeverity.ERROR, status=AccessibilityStatus.SAFE_TO_REMEDIATE, explanation="Missing alt"
    )
    # 3. post for u_chg (changed severity)
    post_chg = UnifiedIssue(
        unified_id="u_chg_post", slide_number=3, shape_id=30, primary_rule_id="SG-RDB-01",
        severity=IssueSeverity.INFO, status=AccessibilityStatus.SAFE_TO_REMEDIATE, explanation="Readability adjusted"
    )
    # 4. post for u_man (identical manual review)
    post_man = UnifiedIssue(
        unified_id="u_man_post", slide_number=4, shape_id=40, primary_rule_id="SG-CNT-01",
        severity=IssueSeverity.WARNING, status=AccessibilityStatus.MANUAL_REVIEW, explanation="Low contrast"
    )
    post_issues = [post_remnant, post_unc, post_chg, post_man]

    action = RemediationAction(
        action_id="act_title", rule_id="SG-TTL-01", slide_number=1, target_shape_id=10,
        description="Promoted shape 10", status=RemediationStatus.APPLIED, confidence=1.0
    )

    summary = reconciler.reconcile(pre_issues=pre_issues, post_issues=post_issues, remediation_actions=[action])

    # 1. Baseline Pre-Issue Partition:
    # total_pre == remediated + unchanged + changed + baseline_manual_review + unsupported
    assert summary.total_pre == 4
    assert summary.remediated_count == 1
    assert summary.unchanged_count == 1
    assert summary.changed_count == 1
    assert summary.baseline_manual_review_count == 1
    assert summary.unsupported_count == 0
    assert summary.total_pre == (
        summary.remediated_count
        + summary.unchanged_count
        + summary.changed_count
        + summary.baseline_manual_review_count
        + summary.unsupported_count
    )

    # 2. Candidate Post-Issue Partition:
    # total_post == manual_review (candidate) + unchanged + changed + new + unsupported
    assert summary.total_post == 4
    assert summary.new_count == 0
    assert summary.decomposed_remnants_count == 1
    assert summary.manual_review_count == summary.baseline_manual_review_count + summary.decomposed_remnants_count
    assert summary.manual_review_count == 2
    assert summary.total_post == (
        summary.manual_review_count
        + summary.unchanged_count
        + summary.changed_count
        + summary.new_count
        + summary.unsupported_count
    )

    # 3. Items list length matches total_pre + new_count (no duplicates)
    assert len(summary.items) == summary.total_pre + summary.new_count == 4


def test_reconciliation_no_remediated_zero_when_actions_applied():
    """Verifies that remediated_count is never 0 when issues are resolved post-remediation,
    even when actions are indexed by slide number or when action is None."""
    reconciler = IssueReconciliationService()

    pre_issue = UnifiedIssue(
        unified_id="u_ro",
        slide_number=5,
        shape_id=None,
        primary_rule_id="SG-RO-01",
        severity=IssueSeverity.WARNING,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        explanation="Reading order non-sequential",
    )

    # Even if action list is empty, if the issue was resolved post-remediation,
    # remediated_count must be >= 1 (never falsely 0)
    summary_empty_actions = reconciler.reconcile(pre_issues=[pre_issue], post_issues=[], remediation_actions=[])
    assert summary_empty_actions.remediated_count == 1
    assert summary_empty_actions.resolved_count == 1
    assert summary_empty_actions.new_count == 0

    # With slide-level action
    action = RemediationAction(
        action_id="act_ro",
        rule_id="SG-RO-01",
        slide_number=5,
        description="Reordered slide 5 shapes",
        status=RemediationStatus.APPLIED,
        confidence=1.0,
    )
    summary_with_action = reconciler.reconcile(pre_issues=[pre_issue], post_issues=[], remediation_actions=[action])
    assert summary_with_action.remediated_count == 1
    assert summary_with_action.items[0].status == ReconciliationStatus.REMEDIATED


def test_reconciliation_successful_remediation_distinct_counts():
    """Verifies successful remediation where raw rule violations resolved (4)
    and unified issues remediated (2) are tracked as distinct metrics without forcing equality."""
    # Raw issues: 2 for reading order (SG-RO-01 + SG-RO-VAL-01), 2 for title (SG-TTL-01 + SG-PH-01)
    ro_raw_1 = Issue(
        issue_id="ro_1", slide_number=1, shape_id=None, rule_id="SG-RO-01", rule_name="RO",
        severity=IssueSeverity.WARNING, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Reading order non-sequential", recommendation="Reorder",
    )
    ro_raw_2 = Issue(
        issue_id="ro_val_1", slide_number=1, shape_id=None, rule_id="SG-RO-VAL-01", rule_name="RO-VAL",
        severity=IssueSeverity.WARNING, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Reading order validation failed", recommendation="Verify sequence",
    )
    ttl_raw_1 = Issue(
        issue_id="ttl_1", slide_number=1, shape_id=7, rule_id="SG-TTL-01", rule_name="TTL",
        severity=IssueSeverity.ERROR, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Missing title", recommendation="Promote text box",
    )
    ttl_raw_2 = Issue(
        issue_id="ph_1", slide_number=1, shape_id=7, rule_id="SG-PH-01", rule_name="PH",
        severity=IssueSeverity.WARNING, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Not native placeholder", recommendation="Convert to placeholder",
    )

    # 2 Unified issues representing the 4 raw findings
    u_ro = UnifiedIssue(
        unified_id="UI_1_DOC_SG-RO-01", slide_number=1, shape_id=None, primary_rule_id="SG-RO-01",
        related_rule_ids=["SG-RO-VAL-01"], severity=IssueSeverity.WARNING,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE, raw_issues=[ro_raw_1, ro_raw_2],
        explanation="Reading sequence invalid", recommended_action="Reorder shapes",
    )
    u_ttl = UnifiedIssue(
        unified_id="UI_1_7_SG-TTL-01", slide_number=1, shape_id=7, primary_rule_id="SG-TTL-01",
        related_rule_ids=["SG-PH-01"], severity=IssueSeverity.ERROR,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE, raw_issues=[ttl_raw_1, ttl_raw_2],
        explanation="Missing title placeholder", recommended_action="Promote shape 7",
    )

    actions = [
        RemediationAction(
            action_id="act_ro_1", rule_id="SG-RO-01", slide_number=1,
            description="Reordered slide 1 shapes in-place", status=RemediationStatus.APPLIED, confidence=0.95
        ),
        RemediationAction(
            action_id="act_ttl_1", rule_id="SG-TTL-01", slide_number=1, target_shape_id=7,
            description="Promoted shape 7 to native title placeholder", status=RemediationStatus.APPLIED, confidence=0.96
        ),
    ]

    reconciler = IssueReconciliationService()
    raw_pre = [ro_raw_1, ro_raw_2, ttl_raw_1, ttl_raw_2]
    summary = reconciler.reconcile(
        pre_issues=[u_ro, u_ttl],
        post_issues=[],
        remediation_actions=actions,
        raw_pre_issues=raw_pre,
        raw_post_issues=[]
    )

    # Distinct metrics verified:
    assert summary.total_pre == 2
    assert summary.total_post == 0
    assert summary.remediated_count == 2
    assert summary.resolved_count == 2
    assert summary.raw_issues_pre == 4
    assert summary.raw_issues_post == 0
    assert summary.raw_issues_remediated == 4
    # Raw resolved (4) != Unified remediated (2)
    assert summary.raw_issues_remediated != summary.remediated_count
    assert len(summary.items) == 2
    assert summary.items[0].action_taken == "Reordered slide 1 shapes in-place"
    assert summary.items[1].action_taken == "Promoted shape 7 to native title placeholder"


def test_reconciliation_multiple_raw_findings_grouped_into_one_unified_issue():
    """Verifies that when 3 raw findings on one shape are grouped into 1 UnifiedIssue,
    remediating the primary rule leaves the secondary remnant for manual review,
    recording 1 unified remediated and 2 raw violations resolved."""
    raw_ttl = Issue(
        issue_id="ttl_1", slide_number=1, shape_id=5, rule_id="SG-TTL-01", rule_name="Title Presence",
        severity=IssueSeverity.ERROR, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Missing title", recommendation="Promote shape 5",
    )
    raw_ph = Issue(
        issue_id="ph_1", slide_number=1, shape_id=5, rule_id="SG-PH-01", rule_name="Placeholder",
        severity=IssueSeverity.WARNING, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        description="Not native placeholder", recommendation="Convert to placeholder",
    )
    raw_cnt = Issue(
        issue_id="cnt_1", slide_number=1, shape_id=5, rule_id="SG-CNT-01", rule_name="Color Contrast",
        severity=IssueSeverity.WARNING, behavior=RuleBehavior.MANUAL_REVIEW,
        description="Low contrast color", recommendation="Adjust contrast",
    )

    u_pre = UnifiedIssue(
        unified_id="UI_1_5_SG-TTL-01", slide_number=1, shape_id=5, primary_rule_id="SG-TTL-01",
        related_rule_ids=["SG-PH-01", "SG-CNT-01"], severity=IssueSeverity.ERROR,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE, raw_issues=[raw_ttl, raw_ph, raw_cnt],
        explanation="Title box has low contrast and lacks placeholder tags",
    )

    action = RemediationAction(
        action_id="act_title_5", rule_id="SG-TTL-01", slide_number=1, target_shape_id=5,
        description="Promoted shape 5 to title placeholder", status=RemediationStatus.APPLIED, confidence=0.96
    )

    # Post-remediation: Title and placeholder are resolved, but contrast remains on shape 5
    u_post_remnant = UnifiedIssue(
        unified_id="UI_1_5_SG-CNT-01", slide_number=1, shape_id=5, primary_rule_id="SG-CNT-01",
        related_rule_ids=[], severity=IssueSeverity.WARNING, status=AccessibilityStatus.MANUAL_REVIEW,
        explanation="Low contrast color on shape 5", raw_issues=[raw_cnt],
    )

    reconciler = IssueReconciliationService()
    summary = reconciler.reconcile(
        pre_issues=[u_pre],
        post_issues=[u_post_remnant],
        remediation_actions=[action],
        raw_pre_issues=[raw_ttl, raw_ph, raw_cnt],
        raw_post_issues=[raw_cnt],
    )

    assert summary.total_pre == 1
    assert summary.total_post == 1
    assert summary.remediated_count == 1
    assert summary.decomposed_remnants_count == 1
    assert summary.manual_review_count == 1
    assert summary.raw_issues_pre == 3
    assert summary.raw_issues_post == 1
    assert summary.raw_issues_remediated == 2
    # Verify distinction
    assert summary.raw_issues_remediated == 2
    assert summary.remediated_count == 1


def test_reconciliation_unrelated_actions_not_falsely_correlated():
    """Verifies that remediation actions applied to completely different slides or shapes
    are not correlated to unrelated issues."""
    u_cnt = UnifiedIssue(
        unified_id="UI_3_12_SG-CNT-01", slide_number=3, shape_id=12, primary_rule_id="SG-CNT-01",
        severity=IssueSeverity.WARNING, status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        explanation="Low contrast on slide 3 shape 12",
    )
    # Action on slide 1 shape 2 (completely different slide and rule)
    unrelated_action = RemediationAction(
        action_id="act_title_slide1", rule_id="SG-TTL-01", slide_number=1, target_shape_id=2,
        description="Promoted shape 2 on slide 1", status=RemediationStatus.APPLIED, confidence=0.95
    )

    reconciler = IssueReconciliationService()
    # u_cnt remains post-remediation
    summary = reconciler.reconcile(
        pre_issues=[u_cnt],
        post_issues=[u_cnt],
        remediation_actions=[unrelated_action],
    )

    assert summary.total_pre == 1
    assert summary.total_post == 1
    assert summary.remediated_count == 0
    assert summary.unchanged_count == 1
    assert summary.items[0].action_taken is None
    assert summary.items[0].status == ReconciliationStatus.UNCHANGED


def test_reconciliation_zero_action_cases():
    """Verifies that when zero actions are applied and issues remain unchanged,
    remediated_count and raw_issues_remediated remain exactly 0."""
    raw1 = Issue(
        issue_id="alt_1", slide_number=2, shape_id=10, rule_id="SG-ALT-01", rule_name="Alt Text",
        severity=IssueSeverity.ERROR, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        description="Missing alt", recommendation="Provide alt",
    )
    raw2 = Issue(
        issue_id="cnt_1", slide_number=2, shape_id=10, rule_id="SG-CNT-01", rule_name="Contrast",
        severity=IssueSeverity.WARNING, behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE,
        description="Low contrast", recommendation="Adjust color",
    )

    u_issue = UnifiedIssue(
        unified_id="UI_2_10_SG-ALT-01", slide_number=2, shape_id=10, primary_rule_id="SG-ALT-01",
        related_rule_ids=["SG-CNT-01"], severity=IssueSeverity.ERROR,
        status=AccessibilityStatus.SAFE_TO_REMEDIATE, raw_issues=[raw1, raw2],
        explanation="Image lacks alt text and caption has low contrast",
    )

    reconciler = IssueReconciliationService()
    summary = reconciler.reconcile(
        pre_issues=[u_issue],
        post_issues=[u_issue],
        remediation_actions=[],
        raw_pre_issues=[raw1, raw2],
        raw_post_issues=[raw1, raw2],
    )

    assert summary.total_pre == 1
    assert summary.total_post == 1
    assert summary.remediated_count == 0
    assert summary.resolved_count == 0
    assert summary.unchanged_count == 1
    assert summary.raw_issues_pre == 2
    assert summary.raw_issues_post == 2
    assert summary.raw_issues_remediated == 0
    assert summary.items[0].action_taken is None
    assert summary.items[0].status == ReconciliationStatus.UNCHANGED
