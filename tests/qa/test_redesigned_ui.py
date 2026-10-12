"""
Comprehensive tests for SlideGuard Redesigned Enterprise UI.
Validates:
1. 5-Section Navigation: Dashboard, Slides, Reading Order, Remediation, Report.
2. Dashboard real metrics: Score / 100, Errors, Warnings, Review, Passed, RO summary, Placeholder summary.
3. Category filters and shortcuts.
4. Slides view deck navigator, filtering, and selected slide inspector.
5. Reading Order view: Check Reading Order and Fix Reading Order (batch).
6. Remediation view: 4 primary action cards (Reading Order, Placeholders, Validate & Compare, Generate Report).
7. Report view: Pre/post scores, operations summary, per-slide results, QA suite.
"""

import tkinter as tk
from pathlib import Path
from unittest.mock import MagicMock, patch
import pytest

from app.ui.main_window import SlideGuardMainWindow
from models.job import JobContext, QualityMetrics
from models.enums import RemediationStatus, IssueSeverity, RuleStatus, RuleBehavior
from models.remediation import RemediationAction
from models.rule_result import RuleResult
from models.issue import Issue


@pytest.fixture
def tk_app():
    """Initializes and cleans up SlideGuardMainWindow in a Tk root."""
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("Tkinter display not available")

    mock_controller = MagicMock()
    app = SlideGuardMainWindow(root, controller=mock_controller)
    yield app, root, mock_controller
    try:
        root.destroy()
    except Exception:
        pass


def test_five_navigation_sections(tk_app):
    """Verify that the 5 main navigation sections are present and selectable."""
    app, root, _ = tk_app

    tabs = app.notebook.tabs()
    assert len(tabs) == 5, f"Expected 5 navigation tabs, found {len(tabs)}"

    titles = [app.notebook.tab(t, "text").strip() for t in tabs]
    expected = ["Dashboard", "Slides", "Reading Order", "Remediation", "Report"]
    assert titles == expected

    # Verify switching via _select_nav
    for idx in range(5):
        app._select_nav(idx)
        assert app.active_nav_index == idx
        assert app.notebook.index("current") == idx


def test_dashboard_real_data_population(tk_app, tmp_path):
    """Verify that Dashboard metrics reflect real analysis job data."""
    app, root, _ = tk_app

    orig = tmp_path / "presentation.pptx"
    orig.touch()

    ctx = JobContext(
        job_id="job_test_dash",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig,
        total_slides=10
    )
    ctx.metrics = QualityMetrics(
        accessibility_score=85.0,
        total_issues=5,
        error_issues=2,
        warning_issues=2,
        info_issues=1,
        passed_rules=12
    )

    ro_issue = Issue(
        issue_id="ro-1",
        rule_id="SG-RO-01",
        rule_name="Reading Order",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        description="Reading order mismatch",
        recommendation="Reorder",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        extra_data={"original_order": [1, 2], "proposed_order": [2, 1], "confidence": "HIGH"}
    )
    ph_issue = Issue(
        issue_id="ph-1",
        rule_id="SG-PH-01",
        rule_name="Placeholders",
        slide_number=2,
        severity=IssueSeverity.WARNING,
        description="Convert to placeholder",
        recommendation="Convert",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        extra_data={"placeholder_type": "title", "status": "SAFE_TO_CONVERT"}
    )

    ctx.rule_results = [
        RuleResult(rule_id="SG-RO-01", rule_name="Reading Order", slide_number=1, status=RuleStatus.FAILED, issues=[ro_issue]),
        RuleResult(rule_id="SG-PH-01", rule_name="Placeholders", slide_number=2, status=RuleStatus.FAILED, issues=[ph_issue])
    ]

    app.job_context = ctx
    app._on_analysis_finished()

    # Dashboard Score
    assert "85.0" in app.metric_score_var.get()

    # Metric Cards
    assert app.metric_errors_var.get() == "2"
    assert app.metric_warnings_var.get() == "2"
    assert app.metric_review_var.get() == "1"
    assert app.metric_passed_var.get() == "12"

    # Reading Order Summary
    assert app.ro_total_var.get() == "10"
    assert app.ro_incorrect_var.get() == "1"
    assert app.ro_correct_var.get() == "9"

    # Placeholder Summary
    assert app.ph_safe_var.get() == "1"


def test_slides_view_filtering_and_inspection(tk_app, tmp_path):
    """Verify that Slides view lists deck slides and filters correctly."""
    app, root, _ = tk_app

    orig = tmp_path / "slides.pptx"
    orig.touch()

    ctx = JobContext(
        job_id="job_test_slides",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig,
        total_slides=3
    )
    # Slide 1 has error, Slide 2 has warning, Slide 3 is clean
    issue_err = Issue(
        issue_id="i-err", rule_id="SG-TTL-01", rule_name="Title", slide_number=1,
        severity=IssueSeverity.ERROR, description="Missing title", recommendation="Add title", behavior=RuleBehavior.MANUAL_REVIEW
    )
    issue_warn = Issue(
        issue_id="i-warn", rule_id="SG-COL-01", rule_name="Contrast", slide_number=2,
        severity=IssueSeverity.WARNING, description="Low contrast", recommendation="Increase contrast", behavior=RuleBehavior.MANUAL_REVIEW
    )

    ctx.rule_results = [
        RuleResult(rule_id="SG-TTL-01", rule_name="Title", slide_number=1, status=RuleStatus.FAILED, issues=[issue_err]),
        RuleResult(rule_id="SG-COL-01", rule_name="Contrast", slide_number=2, status=RuleStatus.FAILED, issues=[issue_warn]),
    ]
    app.job_context = ctx
    app._on_analysis_finished()

    # Default All filter shows 3 slides
    deck_rows = app.deck_tree.get_children()
    assert len(deck_rows) == 3

    # Filter by Errors
    app.slide_filter_var.set("Errors")
    app._refresh_slide_deck_list()
    deck_rows = app.deck_tree.get_children()
    assert len(deck_rows) == 1
    assert "Slide 1" in app.deck_tree.item(deck_rows[0])["values"][0]

    # Inspect Slide 1
    app._load_slide_inspector(1)
    tree_items = app.tree.get_children()
    assert len(tree_items) == 1
    row_vals = app.tree.item(tree_items[0])["values"]
    assert row_vals[0] == "ERROR"
    assert row_vals[1] == "SG-TTL-01"


def test_reading_order_view_and_batch_fix(tk_app, tmp_path):
    """Verify dedicated reading order view and batch fix action."""
    app, root, mock_controller = tk_app

    orig = tmp_path / "ro.pptx"
    orig.touch()

    ctx = JobContext(
        job_id="job_test_ro",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig,
        total_slides=2
    )
    ro_issue = Issue(
        issue_id="ro-1",
        rule_id="SG-RO-01",
        rule_name="Reading Order",
        slide_number=1,
        severity=IssueSeverity.WARNING,
        description="Reading order mismatch",
        recommendation="Reorder",
        behavior=RuleBehavior.SAFE_AUTO_REMEDIATE,
        extra_data={"original_order": [10, 20], "proposed_order": [20, 10], "confidence": "HIGH", "reasoning": ["Title first"]}
    )
    ctx.rule_results = [
        RuleResult(rule_id="SG-RO-01", rule_name="Reading Order", slide_number=1, status=RuleStatus.FAILED, issues=[ro_issue])
    ]
    app.job_context = ctx
    app._on_analysis_finished()

    ro_rows = app.ro_tree.get_children()
    assert len(ro_rows) == 1
    r_vals = app.ro_tree.item(ro_rows[0])["values"]
    assert r_vals[0] == 1
    assert r_vals[1] == "HIGH"
    assert "[10, 20]" in str(r_vals[2])
    assert "[20, 10]" in str(r_vals[3])

    # Test Fix Reading Order (Batch) button triggers execution
    app.approval_var.set(True)
    with patch("threading.Thread") as mock_thread:
        app.btn_fix_ro.invoke()
        assert mock_thread.called


def test_remediation_workspace_primary_cards(tk_app):
    """Verify the 4 primary actions in Remediation workspace and absence of top-level ungroup/list sections."""
    app, root, _ = tk_app

    # Verify 4 primary action buttons
    assert hasattr(app, "btn_apply_ro")
    assert hasattr(app, "btn_apply_ph")
    assert hasattr(app, "btn_run_val")
    assert hasattr(app, "generate_report_btn")

    # Verify absence of standalone top tabs for ungrouping and list semantics
    tabs = [app.notebook.tab(t, "text").strip() for t in app.notebook.tabs()]
    assert "Selective Ungrouping" not in tabs
    assert "List Semantics" not in tabs
    assert "Alt Text" not in tabs


def test_report_view_summary_and_actions(tk_app, tmp_path):
    """Verify Report page data population and export options."""
    app, root, _ = tk_app

    orig = tmp_path / "rep.pptx"
    orig.touch()

    ctx = JobContext(
        job_id="job_test_rep",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig,
        post_remediation_score=94.5
    )
    ctx.metrics = QualityMetrics(quality_score_v2=78.0)
    ctx.remediation_actions = [
        RemediationAction(action_id="a1", rule_id="SG-RO-01", slide_number=1, status=RemediationStatus.APPLIED, description="Reordered shapes"),
        RemediationAction(action_id="a2", rule_id="SG-PH-01", slide_number=2, status=RemediationStatus.SKIPPED, description="Preserved native")
    ]
    app.job_context = ctx
    app._populate_reports_tab()

    assert "78.0/100" in app.val_scores_var.get()
    assert "94.5/100" in app.val_scores_var.get()
    assert "Applied: 1" in app.ops_summary_var.get()
    assert "Preserved/Skipped: 1" in app.ops_summary_var.get()

    res_rows = app.slide_results_tree.get_children()
    assert len(res_rows) == 2
