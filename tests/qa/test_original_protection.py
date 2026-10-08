"""
SlideGuard QA Tests: Original File Protection (Phase 10).
Verifies that the original PPTX source file is NEVER modified, corrupted, or overwritten
under any circumstance: normal remediation, pipeline error, or cancellation.
"""

import shutil
import tempfile
from pathlib import Path
import pytest

from engine.qa.qa_service import QAService, calculate_sha256
from workspace.manager import WorkspaceManager
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.remediation.remediation_service import RemediationService
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


def test_original_file_hash_unmodified_after_full_pipeline():
    """Verify SHA-256 hash before and after complete analysis and remediation is 100% identical."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_pptx = tmp / "original_source.pptx"
        create_comprehensive_test_pptx(source_pptx)

        orig_hash = calculate_sha256(source_pptx)
        orig_bytes = source_pptx.read_bytes()

        # Execute full QA run
        res = qa_service.run_comprehensive_qa(source_pptx, profile_name="PROTECTION_TEST")

        assert res.original_protection_status.value == "PASS"
        assert calculate_sha256(source_pptx) == orig_hash
        assert source_pptx.read_bytes() == orig_bytes


def test_original_file_protected_on_pipeline_failure():
    """Verify source file is unharmed even when an unexpected exception occurs inside the pipeline."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        source_pptx = tmp / "original_protected_fail.pptx"
        create_comprehensive_test_pptx(source_pptx)

        def failing_run(src_path: Path):
            ws_mgr = WorkspaceManager(base_workspace_dir=tmp / "workspaces")
            ctx = ws_mgr.create_job(src_path)
            try:
                # Intentionally simulate crash
                raise RuntimeError("Simulated mid-pipeline pipeline crash!")
            finally:
                ws_mgr.cleanup_job(ctx, preserve_on_debug=False)

        res = qa_service.verify_original_protection(source_pptx, failing_run)
        assert res["is_protected"] is True
        assert res["hash_before"] == res["hash_after"]
        assert "Simulated mid-pipeline pipeline crash!" in res["error_encountered"]


def test_original_file_protected_on_real_fixtures():
    """Verify real fixture files remain byte-for-byte identical when processed."""
    qa_service = QAService()
    fixtures = [
        Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx"),
        Path("tests/fixtures/Sec1-5-TrigFunctions (2)original.pptx"),
    ]

    for fix in fixtures:
        assert fix.exists(), f"Fixture {fix} missing"
        hash_before = calculate_sha256(fix)

        def inspection_run(path: Path):
            reader = PresentationReader()
            pres = reader.read(path)
            return len(pres.slides)

        res = qa_service.verify_original_protection(fix, inspection_run)
        assert res["is_protected"] is True
        assert calculate_sha256(fix) == hash_before
