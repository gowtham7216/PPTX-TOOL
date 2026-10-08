"""
SlideGuard QA Tests: End-to-End Pipeline Validation (Phase 10).
Validates the entire SlideGuard lifecycle from end to end:
  INPUT PPTX
  → Phase 2 inspection
  → Phase 3 reading order
  → Phase 4 selective ungrouping
  → Phase 5 placeholder conversion
  → Phase 6 list semantics
  → Phase 7 accessibility validation & safe remediation
  → Phase 8 intelligence & correlation
  → Phase 9 AI-assisted authoring (Gemini 3.8 Flash High)
  → validation
  → reconciliation
  → final report & structured QAResult
"""

import tempfile
from pathlib import Path
import pytest
from pptx import Presentation

from engine.qa.qa_service import QAService, calculate_sha256
from workspace.manager import WorkspaceManager
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.intelligence.correlator import IssueCorrelator
from engine.authoring.authoring_service import AuthoringService
from engine.remediation.remediation_service import RemediationService
from engine.verification.verification_service import VerificationService
from engine.comparison.comparison_service import ComparisonService
from reports.generator import ReportGenerator
from models.ai_authoring import AuthoringStatus
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


def test_full_pipeline_end_to_end():
    """Execute complete SlideGuard pipeline and verify all invariants and models."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_file = tmp / "e2e_source.pptx"
        create_comprehensive_test_pptx(source_file)

        orig_hash = calculate_sha256(source_file)
        orig_bytes = source_file.read_bytes()

        # Workspace initialization
        ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
        context = ws_mgr.create_job(source_file)
        reader = PresentationReader()
        writer = PresentationWriter()
        analyzer = PresentationAnalyzer()
        correlator = IssueCorrelator()
        auth_service = AuthoringService()
        remediator = RemediationService(writer=writer)
        verifier = VerificationService()
        comparer = ComparisonService()
        reporter = ReportGenerator()
        qa_service = QAService()

        # Step 1: Read & inspect
        pres = reader.read(context.working_file)
        assert pres.total_slides == 6

        # Step 2: Analyze
        rule_results = analyzer.analyze(pres)
        context.rule_results = rule_results
        context.unified_issues = correlator.correlate(rule_results)
        context.metrics = analyzer.compute_metrics(rule_results, total_slides=pres.total_slides)
        assert len(rule_results) > 0

        # Step 3: AI Authoring Recommendations
        recs = auth_service.generate_recommendations(context, pres)
        context.ai_recommendations = recs
        assert len(recs) > 0

        # Step 4: User Approval Simulation
        for r in recs:
            if r.status != AuthoringStatus.MANUAL_REVIEW:
                r.approve()
        context.user_approved_remediation = True

        # Step 5: Remediation Pipeline
        cand_path = tmp / "e2e_remediated.pptx"
        context.candidate_file = cand_path
        actions = remediator.remediate(context, pres)
        assert len(actions) == 6  # 6-step pipeline
        writer.write(pres, cand_path)
        assert cand_path.exists()

        try:
            # Step 6: Verification
            verification = verifier.verify_file(cand_path, expected_slide_count=6)
            assert verification.passed is True

            # Step 7: Comparison / Design Preservation
            approved_shape_ids = set()
            for a in actions:
                tid = getattr(a, "target_shape_id", None) or getattr(a, "shape_id", None)
                if tid:
                    approved_shape_ids.add(tid)
            for r_res in context.rule_results:
                for issue in r_res.issues:
                    if issue.shape_id:
                        approved_shape_ids.add(issue.shape_id)
            for rec in context.ai_recommendations:
                if rec.shape_id:
                    approved_shape_ids.add(rec.shape_id)

            cand_pres = reader.read(cand_path)
            comparison = comparer.compare(
                original=reader.read(source_file),
                remediated=cand_pres,
                intentional_shape_ids=list(approved_shape_ids)
            )
            assert comparison.integrity_preserved is True

            # Step 8: Reports
            json_rep = context.reports_dir / "final_report.json"
            md_rep = context.reports_dir / "final_report.md"
            reporter.generate_json_report(context, json_rep)
            reporter.generate_markdown_report(context, md_rep)
            assert json_rep.exists()
            assert md_rep.exists()

            # Step 9: Source immutability
            assert calculate_sha256(source_file) == orig_hash
            assert source_file.read_bytes() == orig_bytes

            # Step 10: Run QAService validation
            qa_res = qa_service.run_comprehensive_qa(source_file, profile_name="END_TO_END")
            assert qa_res.integrity_status.value == "PASS"
            assert qa_res.original_protection_status.value == "PASS"
            assert qa_res.animation_preservation_status.value == "PASS"
            assert qa_res.overall_status.value in ["PASS", "PASS_WITH_WARNINGS"]
        finally:
            # Cleanup
            ws_mgr.cleanup_job(context, preserve_on_debug=False)
