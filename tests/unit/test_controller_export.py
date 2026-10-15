"""
Unit tests for SlideGuard AppController export and collision management.
Verifies:
1. One PPTX export per remediation action (no intermediate double-exports).
2. Multiple remediation actions in one job reuse the same output path.
3. Automatic and manual report generation reuse the same report path.
4. A new job does not overwrite a previous job's outputs.
5. The original PPTX is never overwritten under any circumstances.
6. Existing collision-safe naming behavior remains correct.
"""

import hashlib
import shutil
from pathlib import Path
import pytest

from app.controller.app_controller import AppController
from workspace.manager import WorkspaceManager
from models.enums import JobStage


def calculate_sha256(path: Path) -> str:
    """Calculates SHA256 of a file."""
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


@pytest.fixture
def test_setup(tmp_path, sample_presentation_path):
    """Sets up an isolated test directory with source presentation and workspace."""
    test_dir = tmp_path / "test_run"
    test_dir.mkdir(parents=True, exist_ok=True)
    source_pptx = test_dir / "TestPresentation.pptx"
    shutil.copy2(sample_presentation_path, source_pptx)
    initial_hash = calculate_sha256(source_pptx)

    ws_dir = tmp_path / "workspaces"
    ws_mgr = WorkspaceManager(base_workspace_dir=ws_dir)
    controller = AppController(workspace_manager=ws_mgr)

    return {
        "test_dir": test_dir,
        "source_pptx": source_pptx,
        "initial_hash": initial_hash,
        "controller": controller,
    }


def test_one_pptx_export_per_remediation_action(test_setup):
    """Verify that a single remediation action creates exactly one step PPTX, not numbered duplicates."""
    controller = test_setup["controller"]
    source = test_setup["source_pptx"]
    test_dir = test_setup["test_dir"]

    context = controller.initialize_job(source)
    context = controller.execute_analysis(context)
    context.user_approved_remediation = True

    context = controller.execute_reading_order_remediation(context)
    assert context.stage == JobStage.COMPLETED
    assert context.final_output_file is not None
    assert context.final_output_file.exists()

    # Verify step remediated file was created in test_dir
    remediated_files = list(test_dir.glob("*_reading_order_fixed*.pptx"))
    assert len(remediated_files) == 1, f"Expected 1 remediated file, found: {remediated_files}"
    assert remediated_files[0].name == "TestPresentation_reading_order_fixed.pptx"


def test_multiple_remediation_actions_produce_step_outputs(test_setup):
    """Verify that manual remediation steps export distinct step outputs while preserving cumulative state."""
    controller = test_setup["controller"]
    source = test_setup["source_pptx"]
    test_dir = test_setup["test_dir"]

    context = controller.initialize_job(source)
    context = controller.execute_analysis(context)
    context.user_approved_remediation = True

    # Step 1: Placeholders
    context = controller.execute_placeholder_remediation(context)
    ph_output = context.final_output_file
    assert ph_output.name == "TestPresentation_placeholder_fixed.pptx"
    assert ph_output.exists()

    # Step 2: Reading Order
    context = controller.execute_reading_order_remediation(context)
    ro_output = context.final_output_file
    assert ro_output.name == "TestPresentation_reading_order_fixed.pptx"
    assert ro_output.exists()

    # Step 3: Outline View Rearrange
    context = controller.execute_outline_order_remediation(context)
    outline_output = context.final_output_file
    assert outline_output.name == "TestPresentation_outline_view_fixed.pptx"
    assert outline_output.exists()

    # Re-running same step within same job updates that step's file without creating (1)
    context = controller.execute_reading_order_remediation(context)
    assert context.final_output_file == ro_output
    assert not (test_dir / "TestPresentation_reading_order_fixed (1).pptx").exists()


def test_automatic_and_manual_report_reuse_same_report_path(test_setup):
    """Verify that automatic export during remediation and manual 'Generate Audit Report' update the same report."""
    controller = test_setup["controller"]
    source = test_setup["source_pptx"]
    test_dir = test_setup["test_dir"]

    context = controller.initialize_job(source)
    context = controller.execute_analysis(context)
    context.user_approved_remediation = True

    # Remediation automatically exports report
    context = controller.execute_reading_order_remediation(context)
    first_report = context.final_report_file
    assert first_report is not None
    assert first_report.exists()
    assert first_report.name == "TestPresentation_SlideGuard_Report.md"

    # User manually clicks "Generate Audit Report"
    context = controller.generate_reports(context)
    assert context.final_report_file == first_report

    # Verify still exactly one report markdown file in test_dir
    report_files = list(test_dir.glob("*_SlideGuard_Report*.md"))
    assert len(report_files) == 1, f"Expected 1 report file, found: {report_files}"
    assert report_files[0] == first_report


def test_new_job_does_not_overwrite_previous_job_outputs(test_setup):
    """Verify that a brand new job preserves outputs from prior jobs and uses collision-safe numbering."""
    controller = test_setup["controller"]
    source = test_setup["source_pptx"]
    test_dir = test_setup["test_dir"]

    # Job 1 execution
    ctx1 = controller.initialize_job(source)
    ctx1 = controller.execute_analysis(ctx1)
    ctx1.user_approved_remediation = True
    ctx1 = controller.execute_reading_order_remediation(ctx1)

    job1_pptx = ctx1.final_output_file
    job1_report = ctx1.final_report_file
    assert job1_pptx.name == "TestPresentation_reading_order_fixed.pptx"
    assert job1_report.name == "TestPresentation_SlideGuard_Report.md"

    # Job 2 execution (new job on same source file)
    controller.current_presentation = None
    ctx2 = controller.initialize_job(source)
    ctx2 = controller.execute_analysis(ctx2)
    ctx2.user_approved_remediation = True
    ctx2 = controller.execute_reading_order_remediation(ctx2)

    job2_pptx = ctx2.final_output_file
    job2_report = ctx2.final_report_file

    assert job2_pptx.name == "TestPresentation_reading_order_fixed (1).pptx"
    assert job2_report.name == "TestPresentation_SlideGuard_Report (1).md"

    # Both Job 1 and Job 2 files must exist independently
    assert job1_pptx.exists()
    assert job2_pptx.exists()
    assert job1_report.exists()
    assert job2_report.exists()


def test_original_pptx_is_never_overwritten(test_setup):
    """Verify that the original source PPTX hash remains completely untouched across all operations."""
    controller = test_setup["controller"]
    source = test_setup["source_pptx"]
    initial_hash = test_setup["initial_hash"]

    context = controller.initialize_job(source)
    context = controller.execute_analysis(context)
    context.user_approved_remediation = True
    context = controller.execute_remediation(context)
    context = controller.generate_reports(context)

    # Source hash must remain 100% identical
    final_hash = calculate_sha256(source)
    assert final_hash == initial_hash, "Original PPTX was mutated!"

    # Verify safety guard prevents accidental overwrite if output path is set to source
    context.final_output_file = source
    controller._export_candidate_pptx(context)
    assert calculate_sha256(source) == initial_hash, "Guard failed to prevent source overwrite!"


def test_collision_safe_path_helper(tmp_path):
    """Verify static collision-safe path generation."""
    src = tmp_path / "Deck.pptx"
    src.touch()

    # First safe path
    p0 = AppController.get_collision_safe_path(tmp_path, "Deck", "_Remediated", ".pptx", [src])
    assert p0.name == "Deck_Remediated.pptx"
    p0.touch()

    # Second safe path when p0 exists
    p1 = AppController.get_collision_safe_path(tmp_path, "Deck", "_Remediated", ".pptx", [src])
    assert p1.name == "Deck_Remediated (1).pptx"
    p1.touch()

    # Third safe path when p0 and p1 exist
    p2 = AppController.get_collision_safe_path(tmp_path, "Deck", "_Remediated", ".pptx", [src])
    assert p2.name == "Deck_Remediated (2).pptx"

    # Forbidden path enforcement: if base name is forbidden, must not return it
    pf = AppController.get_collision_safe_path(tmp_path, "Deck", "", ".pptx", [src])
    assert pf.resolve() != src.resolve()
    assert pf.name == "Deck (1).pptx"
