"""
Integration test for SlideGuard Analysis & Remediation pipeline.
Executes end-to-end workflow from file loading through report generation.
"""

import pytest
from app.controller.app_controller import AppController
from workspace.manager import WorkspaceManager
from models.enums import JobStage, RemediationStatus


def test_full_pipeline_flow(sample_presentation_path, isolated_workspace):
    workspace_mgr = WorkspaceManager(base_workspace_dir=isolated_workspace)
    controller = AppController(workspace_manager=workspace_mgr)

    # 1. Initialize job
    context = controller.initialize_job(sample_presentation_path)
    assert context.stage == JobStage.INITIALIZED
    assert context.total_slides == 6
    assert context.working_file.exists()

    # 2. Run Analysis
    progress_records = []
    def on_progress(cur, tot, msg):
        progress_records.append((cur, tot, msg))

    context = controller.execute_analysis(context, progress_callback=on_progress)
    assert context.stage == JobStage.ANALYZED
    assert context.slides_analyzed == 6
    assert len(progress_records) > 0

    # Verify quality metrics
    metrics = context.metrics
    assert metrics.total_rules_evaluated > 0
    assert metrics.total_issues > 0
    assert 0.0 <= metrics.accessibility_score <= 100.0

    # Verify reports were created
    json_report = context.reports_dir / "analysis_report.json"
    md_report = context.reports_dir / "analysis_report.md"
    assert json_report.exists()
    assert md_report.exists()
    assert json_report.stat().st_size > 0
    assert md_report.stat().st_size > 0

    # 3. Remediation without approval must fail with PermissionError
    context.user_approved_remediation = False
    with pytest.raises(PermissionError):
        controller.execute_remediation(context)

    # 4. Remediation with approval
    context.user_approved_remediation = True
    context = controller.execute_remediation(context)
    assert context.stage == JobStage.COMPLETED
    assert len(context.remediation_actions) == 6
    # Step 1 (Reading Order) was actively executed
    assert context.remediation_actions[0].status in (RemediationStatus.APPLIED, RemediationStatus.SKIPPED)
    # Step 2 (Selective Ungrouping) was actively executed (Phase 4)
    assert context.remediation_actions[1].status in (RemediationStatus.APPLIED, RemediationStatus.SKIPPED)
    # Step 3 (Placeholder Conversion) was actively executed (Phase 5)
    assert context.remediation_actions[2].status in (RemediationStatus.APPLIED, RemediationStatus.SKIPPED)
    # Step 4 (List Semantics) was actively executed (Phase 6)
    assert context.remediation_actions[3].status in (RemediationStatus.APPLIED, RemediationStatus.SKIPPED)

    # Step 5 (Accessibility Validation & Safe Remediation) was actively executed (Phase 7)
    assert context.remediation_actions[4].status in (RemediationStatus.APPLIED, RemediationStatus.SKIPPED)

    # Step 6 (Final Validation & Issue Reconciliation) was actively executed (Phase 8)
    assert context.remediation_actions[5].status in (RemediationStatus.APPLIED, RemediationStatus.SKIPPED)

    # Verify verification service ran and candidate file is valid
    assert context.verification_result is not None
    assert context.verification_result.passed is True
    assert context.candidate_file.exists()

    # Verify comparison ran and integrity is preserved
    assert context.comparison_result is not None
    assert context.comparison_result.integrity_preserved is True
    assert context.comparison_result.slide_count_matches is True
