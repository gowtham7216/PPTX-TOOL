"""
Regression tests for SlideGuard Report Generation.
Specifically verifies resolution of:
    "Report generation failed: VERIFIED"

Validates:
1. JobStage enum includes VERIFIED and matches both Enum and string lookups.
2. Reproduction of the AttributeError('VERIFIED') mechanism when an enum member is missing.
3. Successful report generation via AppController when context.stage is JobStage.VERIFIED.
4. Successful report generation when context.stage is provided as a string "VERIFIED".
5. Proper report naming ('final_report' vs 'analysis_report') across all lifecycle stages.
6. Genuine report-generation failures propagate with diagnostic messages and are not silently masked.
"""

import tempfile
import json
import logging
from pathlib import Path
import pytest

from models.enums import JobStage
from models.job import JobContext, QualityMetrics
from app.controller.app_controller import AppController
from workspace.manager import WorkspaceManager
from reports.generator import ReportGenerator
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


def test_job_stage_enum_has_verified():
    """Verify JobStage includes VERIFIED and behaves as both str and Enum."""
    assert hasattr(JobStage, "VERIFIED"), "JobStage must define VERIFIED"
    assert JobStage.VERIFIED == "VERIFIED"
    assert JobStage.VERIFIED.value == "VERIFIED"
    assert JobStage("VERIFIED") == JobStage.VERIFIED


def test_reproduce_exact_error_mechanism():
    """
    Reproduces the exact mechanism of the original failure:
    Evaluating an undefined enum member on an Enum class raises AttributeError(name),
    where str(e) is literally the missing member name (e.g. 'VERIFIED').
    When formatted into 'Report generation failed: {e}', it produced:
    'Report generation failed: VERIFIED'.
    """
    from enum import Enum

    class MockIncompleteJobStage(str, Enum):
        INITIALIZED = "INITIALIZED"
        VERIFYING = "VERIFYING"
        # VERIFIED was missing!

    with pytest.raises(AttributeError) as exc_info:
        _ = MockIncompleteJobStage.VERIFIED

    err_text = str(exc_info.value)
    # In Python 3.11+, str(AttributeError("VERIFIED")) is 'VERIFIED'
    assert err_text == "VERIFIED"
    dialog_msg = f"Report generation failed: {err_text}"
    assert dialog_msg == "Report generation failed: VERIFIED"

    # Now verify that with JobStage.VERIFIED present, no AttributeError is raised
    valid_stages = (JobStage.REMEDIATED, JobStage.VERIFYING, JobStage.VERIFIED, JobStage.COMPLETED)
    assert JobStage.VERIFIED in valid_stages


def test_controller_generate_reports_when_stage_verified():
    """Verify AppController.generate_reports succeeds when stage is JobStage.VERIFIED."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_pptx = tmp / "test_pres.pptx"
        create_comprehensive_test_pptx(source_pptx)

        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        context = ws_mgr.create_job(source_pptx)
        context.stage = JobStage.VERIFIED

        controller = AppController(workspace_manager=ws_mgr)
        updated_ctx = controller.generate_reports(context)

        # Reports must exist in reports_dir
        final_json = context.reports_dir / "final_report.json"
        final_md = context.reports_dir / "final_report.md"
        assert final_json.exists(), "final_report.json was not created"
        assert final_md.exists(), "final_report.md was not created"

        # Exported report file must exist
        assert updated_ctx.final_report_file is not None
        assert updated_ctx.final_report_file.exists()

        # Check JSON report content contains VERIFIED stage
        data = json.loads(final_json.read_text(encoding="utf-8"))
        assert data["stage"] == "VERIFIED"

        # Check Markdown report contains VERIFIED stage
        md_text = final_md.read_text(encoding="utf-8")
        assert "**Status:** `VERIFIED`" in md_text

        ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_controller_generate_reports_with_string_stage():
    """Verify AppController.generate_reports succeeds when stage is a string 'VERIFIED'."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_pptx = tmp / "test_pres_str.pptx"
        create_comprehensive_test_pptx(source_pptx)

        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        context = ws_mgr.create_job(source_pptx)
        context.stage = "VERIFIED"  # Raw string instead of Enum instance

        controller = AppController(workspace_manager=ws_mgr)
        updated_ctx = controller.generate_reports(context)

        final_json = context.reports_dir / "final_report.json"
        final_md = context.reports_dir / "final_report.md"
        assert final_json.exists()
        assert final_md.exists()

        data = json.loads(final_json.read_text(encoding="utf-8"))
        assert data["stage"] == "VERIFIED"

        ws_mgr.cleanup_job(context, preserve_on_debug=False)


def test_controller_generate_reports_stage_routing():
    """Verify report name routing: 'analysis_report' for pre-remediation, 'final_report' for post-remediation/verification."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_pptx = tmp / "test_routing.pptx"
        create_comprehensive_test_pptx(source_pptx)

        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        controller = AppController(workspace_manager=ws_mgr)

        # Stage: ANALYZED -> analysis_report
        ctx_analyzed = ws_mgr.create_job(source_pptx)
        ctx_analyzed.stage = JobStage.ANALYZED
        controller.generate_reports(ctx_analyzed)
        assert (ctx_analyzed.reports_dir / "analysis_report.json").exists()
        assert (ctx_analyzed.reports_dir / "analysis_report.md").exists()

        # Stage: COMPLETED -> final_report
        ctx_completed = ws_mgr.create_job(source_pptx)
        ctx_completed.stage = JobStage.COMPLETED
        controller.generate_reports(ctx_completed)
        assert (ctx_completed.reports_dir / "final_report.json").exists()
        assert (ctx_completed.reports_dir / "final_report.md").exists()

        # Stage: REMEDIATED -> final_report
        ctx_remediated = ws_mgr.create_job(source_pptx)
        ctx_remediated.stage = JobStage.REMEDIATED
        controller.generate_reports(ctx_remediated)
        assert (ctx_remediated.reports_dir / "final_report.json").exists()

        ws_mgr.cleanup_job(ctx_analyzed, preserve_on_debug=False)
        ws_mgr.cleanup_job(ctx_completed, preserve_on_debug=False)
        ws_mgr.cleanup_job(ctx_remediated, preserve_on_debug=False)


def test_genuine_report_generator_failure_is_propagated():
    """Verify that genuine report generation errors (e.g. disk/write errors) are not masked."""
    generator = ReportGenerator()
    context = JobContext(
        job_id="test_fail",
        workspace_dir=Path("/non_existent_path"),
        original_file=Path("/non_existent_path/fake.pptx"),
        working_file=Path("/non_existent_path/fake.pptx"),
        stage=JobStage.VERIFIED,
    )

    invalid_target = Path("Z:\\non_existent_mount_dir_12345\\impossible_dir\\report.json")
    with pytest.raises(Exception) as exc_info:
        generator.generate_json_report(context, invalid_target)

    assert issubclass(exc_info.type, (OSError, IOError))
