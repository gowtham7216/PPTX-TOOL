"""
SlideGuard QA Tests: Idempotency & Repeated Remediation (Phase 10).
Verifies that running SlideGuard on its own output does NOT:
- re-convert placeholders
- re-transform or strip list formatting
- re-order already aligned reading orders
- degrade or destabilize accessibility scores
- introduce duplicate shape IDs or structural drift
"""

import tempfile
from pathlib import Path
import pytest

from engine.qa.qa_service import QAService
from workspace.manager import WorkspaceManager
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.remediation.remediation_service import RemediationService
from engine.intelligence.correlator import IssueCorrelator
from tests.qa.synthetic_profiles import create_profile_normal, create_profile_list_heavy


def _run_pipeline(input_path: Path, output_path: Path):
    reader = PresentationReader()
    writer = PresentationWriter()
    analyzer = PresentationAnalyzer()
    correlator = IssueCorrelator()
    remediator = RemediationService(writer=writer)

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        context = ws_mgr.create_job(input_path)
        pres = reader.read(context.working_file)

        rule_results = analyzer.analyze(pres)
        context.rule_results = rule_results
        context.unified_issues = correlator.correlate(rule_results)
        context.metrics = analyzer.compute_metrics(rule_results, total_slides=pres.total_slides)

        context.user_approved_remediation = True
        context.candidate_file = output_path
        actions = remediator.remediate(context, pres)
        writer.write(pres, output_path)

        # Release log handlers
        ws_mgr.cleanup_job(context, preserve_on_debug=False)

    return context, pres


def test_idempotency_on_normal_presentation():
    """Verify idempotency on normal multi-slide presentation."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source = tmp / "normal_input.pptx"
        create_profile_normal(source)

        res = qa_service.verify_idempotency(source, _run_pipeline)
        assert res["is_idempotent"] is True, f"Idempotency failed: {res['errors']}"
        assert res["run_2_score"] >= res["run_1_score"] - 0.5


def test_idempotency_on_list_heavy_presentation():
    """Verify list formatting is not duplicated or stripped on secondary run."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source = tmp / "list_input.pptx"
        create_profile_list_heavy(source)

        res = qa_service.verify_idempotency(source, _run_pipeline)
        assert res["is_idempotent"] is True, f"List idempotency failed: {res['errors']}"
        assert res["run_2_actions_applied"] <= 1  # only reconciliation or 0 list actions


def test_score_stability_across_repeated_runs():
    """Verify accessibility score is deterministic and stable across repeated runs."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source = tmp / "score_test.pptx"
        create_profile_normal(source)

        out_a = tmp / "run_a.pptx"
        out_b = tmp / "run_b.pptx"

        ctx_a, pres_a = _run_pipeline(source, out_a)
        ctx_b, pres_b = _run_pipeline(out_a, out_b)

        assert ctx_a.metrics.quality_score_v2 <= 100.0
        assert ctx_b.metrics.quality_score_v2 <= 100.0
        # Score must not decrease on repeated remediation
        assert ctx_b.metrics.quality_score_v2 >= ctx_a.metrics.quality_score_v2 - 0.1
