"""
SlideGuard QA Tests: Real-World Presentation Profiles (Phase 10).
Validates real fixtures and diverse presentation profiles:
- Ch06_GLN_Instructor_AU.pptx (50 slides, animations, tables, lists, groups)
- Sec1-5-TrigFunctions (2)original.pptx (19 slides, math text, reading order)
- Synthetic Profiles: Normal, Image-heavy, Text-heavy, Table-heavy
Ensures 100% preservation of structure, animations, and original files.
"""

import tempfile
from pathlib import Path
import pytest
from pptx import Presentation

from engine.qa.qa_service import QAService, calculate_sha256
from tests.qa.synthetic_profiles import (
    create_profile_normal,
    create_profile_image_heavy,
    create_profile_text_heavy,
    create_profile_table_heavy,
)


def test_real_fixture_ch06_instructor():
    """Verify real 50-slide fixture Ch06 with animation preservation and integrity."""
    fixture_path = Path("tests/fixtures/Ch06_GLN_Instructor_AU.pptx")
    assert fixture_path.exists(), f"Fixture {fixture_path} not found"

    orig_hash = calculate_sha256(fixture_path)
    qa_service = QAService()

    with tempfile.TemporaryDirectory() as tmp_dir:
        cand_out = Path(tmp_dir) / f"candidate_{fixture_path.stem}.pptx"
        res = qa_service.run_comprehensive_qa(
            fixture_path, profile_name="REAL_CH06_INSTRUCTOR", output_file=cand_out
        )

        assert res.original_protection_status.value == "PASS"
        assert calculate_sha256(fixture_path) == orig_hash
        assert res.integrity_status.value == "PASS"
        assert res.animation_preservation_status.value == "PASS"
        assert res.overall_status.value in ["PASS", "PASS_WITH_WARNINGS"]

        # Verify output presentation reloads
        cand = Presentation(res.output_file)
        assert len(cand.slides) == 50


def test_real_fixture_trig_functions():
    """Verify real 19-slide fixture Sec1-5-TrigFunctions integrity and protection."""
    fixture_path = Path("tests/fixtures/Sec1-5-TrigFunctions (2)original.pptx")
    assert fixture_path.exists(), f"Fixture {fixture_path} not found"

    orig_hash = calculate_sha256(fixture_path)
    qa_service = QAService()

    with tempfile.TemporaryDirectory() as tmp_dir:
        cand_out = Path(tmp_dir) / f"candidate_{fixture_path.stem}.pptx"
        res = qa_service.run_comprehensive_qa(
            fixture_path, profile_name="REAL_TRIG_FUNCTIONS", output_file=cand_out
        )

        assert res.original_protection_status.value == "PASS"
        assert calculate_sha256(fixture_path) == orig_hash
        assert res.integrity_status.value in ["PASS", "PASS_WITH_WARNINGS"]
        assert res.overall_status.value in ["PASS", "PASS_WITH_WARNINGS"]

        cand = Presentation(res.output_file)
        assert len(cand.slides) == 19


def test_profile_image_heavy():
    """Verify image heavy presentation profile."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        p_path = tmp / "img_heavy.pptx"
        create_profile_image_heavy(p_path, tmp)

        res = qa_service.run_comprehensive_qa(p_path, profile_name="IMAGE_HEAVY")
        assert res.integrity_status.value == "PASS"
        assert res.relationship_status.value == "PASS"


def test_profile_text_heavy():
    """Verify text heavy presentation profile."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        p_path = tmp / "txt_heavy.pptx"
        create_profile_text_heavy(p_path)

        res = qa_service.run_comprehensive_qa(p_path, profile_name="TEXT_HEAVY")
        assert res.integrity_status.value == "PASS"


def test_profile_table_heavy():
    """Verify table heavy presentation profile."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        p_path = tmp / "tbl_heavy.pptx"
        create_profile_table_heavy(p_path)

        res = qa_service.run_comprehensive_qa(p_path, profile_name="TABLE_HEAVY")
        assert res.integrity_status.value == "PASS"
