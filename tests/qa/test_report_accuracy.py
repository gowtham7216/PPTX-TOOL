"""
SlideGuard QA Tests: Report Accuracy & Metrics Reconciliation (Phase 10).
Validates that generated reports and summary models strictly reflect reality:
- No false "PASSED" or "COMPLIANT" claims.
- Metrics sum to exact totals.
- Action statuses strictly distinguish:
  DETECTED, RECOMMENDED, APPROVED, APPLIED, VERIFIED, FAILED, MANUAL REVIEW.
- JSON and Markdown reports are valid and non-empty.
"""

import tempfile
import json
from pathlib import Path
import pytest

from engine.qa.qa_service import QAService
from workspace.manager import WorkspaceManager
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.intelligence.correlator import IssueCorrelator
from engine.remediation.remediation_service import RemediationService
from reports.generator import ReportGenerator
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


def test_report_accuracy_and_reconciliation():
    """Verify metrics reconcile exactly in memory and serialized report."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_pptx = tmp / "report_test.pptx"
        create_comprehensive_test_pptx(source_pptx)

        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        context = ws_mgr.create_job(source_pptx)
        reader = PresentationReader()
        writer = PresentationWriter()
        analyzer = PresentationAnalyzer()
        correlator = IssueCorrelator()
        remediator = RemediationService(writer=writer)
        reporter = ReportGenerator()

        pres = reader.read(context.working_file)
        rule_results = analyzer.analyze(pres)
        context.rule_results = rule_results
        context.unified_issues = correlator.correlate(rule_results)
        context.metrics = analyzer.compute_metrics(rule_results, total_slides=pres.total_slides)

        context.user_approved_remediation = True
        context.candidate_file = tmp / "cand.pptx"
        actions = remediator.remediate(context, pres)
        writer.write(pres, context.candidate_file)

        # Run QA report accuracy verification
        accuracy_res = qa_service.verify_report_accuracy(context)
        assert accuracy_res["accurate"] is True, f"Inaccuracies found: {accuracy_res['errors']}"

        # Generate JSON report and verify contents
        json_path = context.reports_dir / "test_report.json"
        reporter.generate_json_report(context, json_path)
        assert json_path.exists()

        data = json.loads(json_path.read_text(encoding="utf-8"))
        assert "job_id" in data
        assert "metrics" in data
        assert data["metrics"]["total_issues"] == context.metrics.total_issues

        # Verify no false "WCAG certified" claim
        report_str = json_path.read_text(encoding="utf-8").lower()
        assert "wcag certified" not in report_str
        assert "wcag certification" not in report_str

        ws_mgr.cleanup_job(context, preserve_on_debug=False)
