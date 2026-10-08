"""
Tests for Phase 10 QA & Validation Desktop UI Tab.
Verifies that all 10 QA dimensions, status badges, and check tables render and update properly.
"""

import tkinter as tk
from pathlib import Path
import pytest

from app.ui.main_window import SlideGuardMainWindow
from models.qa import QAResult, QAStatus


def test_qa_tab_rendering_and_population():
    """Verify that the QA tab initializes correctly and reflects QAResult updates."""
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("Tkinter display not available")

    try:
        app = SlideGuardMainWindow(root)

        # Verify QA variables exist
        assert hasattr(app, "qa_overall_badge_var")
        assert hasattr(app, "qa_stats_var")
        assert hasattr(app, "qa_dim_ooxml_var")
        assert hasattr(app, "qa_dim_design_var")
        assert hasattr(app, "qa_dim_orig_var")
        assert hasattr(app, "qa_dim_rel_var")
        assert hasattr(app, "qa_dim_acc_var")
        assert hasattr(app, "qa_dim_idemp_var")
        assert hasattr(app, "qa_dim_ai_var")
        assert hasattr(app, "qa_dim_ws_var")
        assert hasattr(app, "qa_dim_rep_var")
        assert hasattr(app, "qa_dim_anim_var")
        assert hasattr(app, "qa_tree")

        # Create sample QAResult
        qa_res = QAResult(
            job_id="test_qa_job",
            input_file="sample.pptx",
            integrity_status=QAStatus.PASS,
            design_preservation_status=QAStatus.PASS,
            original_protection_status=QAStatus.PASS,
            relationship_status=QAStatus.PASS,
            accessibility_consistency_status=QAStatus.PASS,
            idempotency_status=QAStatus.PASS,
            ai_fallback_status=QAStatus.PASS,
            workspace_cleanup_status=QAStatus.PASS,
            report_accuracy_status=QAStatus.PASS,
            animation_preservation_status=QAStatus.PASS,
        )
        qa_res.add_check("Check 1", True, "All valid", component="ooxml")
        qa_res.add_check("Check 2", True, "Animation intact", component="animation")
        qa_res.evaluate_overall()

        # Populate QA tab
        app._populate_qa_tab(qa_res)

        # Assert UI variables updated
        assert "PASS" in app.qa_overall_badge_var.get()
        assert "Passed: 2" in app.qa_stats_var.get()
        assert "PASS" in app.qa_dim_ooxml_var.get()
        assert "PASS" in app.qa_dim_anim_var.get()

        # Assert rows inserted into treeview
        rows = app.qa_tree.get_children()
        assert len(rows) == 2

    finally:
        root.destroy()
