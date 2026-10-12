"""
Tests for SlideGuard Refactored Desktop UI.
Verifies the clean 4-section layout, remediation dashboard separation,
absence of ungrouping/lists from top tabs, independent remediation buttons,
approval gates, and per-slide reporting.
"""

from pathlib import Path
import tkinter as tk
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


def test_five_section_layout(tk_app):
    """Verify that the notebook has exactly 5 clean top-level tabs: Dashboard, Slides, Reading Order, Remediation, Report."""
    app, root, _ = tk_app

    tabs = app.notebook.tabs()
    assert len(tabs) == 5, f"Expected 5 main tabs, found {len(tabs)}"

    tab_titles = [app.notebook.tab(t, "text").strip() for t in tabs]
    expected_tabs = [
        "Dashboard",
        "Slides",
        "Reading Order",
        "Remediation",
        "Report"
    ]
    assert tab_titles == expected_tabs

    # Selective Ungrouping and List Semantics must NOT be top-level tabs
    assert "Selective Ungrouping" not in tab_titles
    assert "List Semantics" not in tab_titles
    assert "List Semantics Review" not in tab_titles


def test_issues_and_review_subtabs_exclusion(tk_app):
    """Verify that Selective Ungrouping and List Semantics are removed from Review subtabs."""
    app, root, _ = tk_app

    subtabs = app.review_notebook.tabs()
    subtab_titles = [app.review_notebook.tab(s, "text").strip() for s in subtabs]

    # Verify expected review subtabs
    assert any("Detected Issues" in t for t in subtab_titles)
    assert any("Alternative Text" in t or "Alt Text" in t for t in subtab_titles)
    assert any("Intelligence" in t for t in subtab_titles)
    assert any("Placeholder" in t for t in subtab_titles)

    # Selective Ungrouping and List Semantics must NOT appear as visible subtabs
    for t in subtab_titles:
        assert "Ungroup" not in t, f"Ungrouping tab found in review tabs: {t}"
        assert "List Semantics" not in t, f"List Semantics tab found in review tabs: {t}"


def test_remediation_dashboard_action_buttons(tk_app):
    """Verify that the centralized Remediation tab has independent action buttons."""
    app, root, _ = tk_app

    # Action Card 1: Reading Order
    assert hasattr(app, "btn_apply_ro")
    assert hasattr(app, "btn_inspect_ro")
    assert hasattr(app, "ro_card_status_var")

    # Action Card 2: Placeholder Conversion
    assert hasattr(app, "btn_apply_ph")
    assert hasattr(app, "btn_inspect_ph")
    assert hasattr(app, "ph_card_status_var")

    # Action Card 3: Alt Text & Decorative Images
    assert hasattr(app, "btn_apply_auth")
    assert hasattr(app, "btn_inspect_auth")
    assert hasattr(app, "auth_card_status_var")

    # Action Card 4: Other Existing Remediations
    assert hasattr(app, "btn_apply_ungroup")
    assert hasattr(app, "btn_apply_lists")
    assert hasattr(app, "btn_run_pipeline")
    assert hasattr(app, "other_card_status_var")

    # Execution log
    assert hasattr(app, "rem_log_text")

    # Reports & Validation Header buttons
    assert hasattr(app, "btn_run_val")
    assert hasattr(app, "generate_report_btn")


def test_user_approval_gate_cancelled(tk_app, tmp_path):
    """Verify that remediation refuses execution if user denies approval."""
    app, root, mock_controller = tk_app

    orig = tmp_path / "test.pptx"
    orig.touch()
    ctx = JobContext(
        job_id="test_gate",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig
    )
    ctx.rule_results = [
        RuleResult(
            rule_id="SG-RO-01",
            rule_name="Reading Order",
            slide_number=1,
            status=RuleStatus.FAILED,
            issues=[
                Issue(
                    issue_id="iss-1",
                    rule_id="SG-RO-01",
                    rule_name="Reading Order",
                    slide_number=1,
                    severity=IssueSeverity.ERROR,
                    description="Out of order",
                    recommendation="Align reading order",
                    behavior=RuleBehavior.SAFE_AUTO_REMEDIATE
                )
            ]
        )
    ]
    ctx.user_approved_remediation = False
    app.job_context = ctx
    app.approval_var.set(False)

    # Mock user declining approval prompt
    with patch("tkinter.messagebox.askyesno", return_value=False):
        app._on_apply_reading_order_clicked()

    # Controller should not have been called
    assert not mock_controller.execute_reading_order_remediation.called
    assert not ctx.user_approved_remediation
    assert "cancelled" in app.status_var.get().lower()


def test_user_approval_gate_approved(tk_app, tmp_path):
    """Verify that user approval allows remediation to execute."""
    app, root, mock_controller = tk_app

    orig = tmp_path / "test.pptx"
    orig.touch()
    ctx = JobContext(
        job_id="test_gate",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig
    )
    ctx.rule_results = [
        RuleResult(
            rule_id="SG-RO-01",
            rule_name="Reading Order",
            slide_number=1,
            status=RuleStatus.FAILED,
            issues=[
                Issue(
                    issue_id="iss-2",
                    rule_id="SG-RO-01",
                    rule_name="Reading Order",
                    slide_number=1,
                    severity=IssueSeverity.ERROR,
                    description="Out of order",
                    recommendation="Align reading order",
                    behavior=RuleBehavior.SAFE_AUTO_REMEDIATE
                )
            ]
        )
    ]
    ctx.user_approved_remediation = False
    app.job_context = ctx
    app.approval_var.set(False)

    # Mock user accepting approval prompt
    with patch("tkinter.messagebox.askyesno", return_value=True):
        with patch("threading.Thread") as mock_thread:
            app._on_apply_reading_order_clicked()
            assert mock_thread.called
            assert ctx.user_approved_remediation is True
            assert app.approval_var.get() is True


def test_populate_reports_tab(tk_app, tmp_path):
    """Verify that _populate_reports_tab properly updates summaries and per-slide treeview."""
    app, root, _ = tk_app

    orig = tmp_path / "test.pptx"
    orig.touch()
    ctx = JobContext(
        job_id="test_reports",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig
    )
    ctx.metrics = QualityMetrics(quality_score_v2=82.5)
    ctx.post_remediation_score = 96.0

    action1 = RemediationAction(
        action_id="act-1",
        rule_id="SG-RO-01",
        slide_number=1,
        status=RemediationStatus.APPLIED,
        description="Reordered shapes logically"
    )
    action2 = RemediationAction(
        action_id="act-2",
        rule_id="SG-PH-01",
        slide_number=2,
        status=RemediationStatus.SKIPPED,
        description="Native placeholder preserved"
    )
    ctx.remediation_actions = [action1, action2]

    app.job_context = ctx
    app._populate_reports_tab()

    assert "82.5/100" in app.val_scores_var.get()
    assert "96.0/100" in app.val_scores_var.get()
    assert "Applied: 1" in app.ops_summary_var.get()
    assert "Preserved/Skipped: 1" in app.ops_summary_var.get()

    tree_rows = app.slide_results_tree.get_children()
    assert len(tree_rows) == 2
    row1 = app.slide_results_tree.item(tree_rows[0])["values"]
    assert row1[0] == 1
    assert row1[1] == "SG-RO-01"
    assert row1[2] == "APPLIED"


def test_concurrency_guard_busy(tk_app, tmp_path):
    """Verify that is_busy flag guards against overlapping operations."""
    app, root, mock_controller = tk_app

    app.is_busy = True
    orig = tmp_path / "test.pptx"
    orig.touch()
    ctx = JobContext(
        job_id="test_busy",
        workspace_dir=tmp_path,
        original_file=orig,
        working_file=orig
    )
    ctx.user_approved_remediation = True
    app.job_context = ctx

    with patch("threading.Thread") as mock_thread:
        app._on_apply_reading_order_clicked()
        assert not mock_thread.called

        app._on_apply_placeholders_clicked()
        assert not mock_thread.called

        app._on_apply_alt_text_clicked()
        assert not mock_thread.called

        app._on_apply_ungrouping_clicked()
        assert not mock_thread.called

        app._on_apply_lists_clicked()
        assert not mock_thread.called

        app._on_apply_validation_clicked()
        assert not mock_thread.called

        app._on_generate_report_clicked()
        assert not mock_thread.called
