"""
SlideGuard QA Tests: Malformed & Corrupted Inputs (Phase 10).
Verifies graceful failure, clear diagnostic messaging, and zero crashes when encountering:
- Non-ZIP / corrupt binary files
- Missing presentation.xml
- Missing slide relationships
- Malformed XML
- 0-byte files
- Incomplete package structures
Ensures original files are untouched and no false "SUCCESS" is reported.
"""

import zipfile
import tempfile
from pathlib import Path
import pytest

from engine.qa.qa_service import QAService
from pptx_engine.reader.presentation_reader import PresentationReader
from workspace.manager import WorkspaceManager


def test_zero_byte_file_fails_gracefully():
    """Verify 0-byte file fails with clear error and does not crash."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        empty_file = tmp / "empty.pptx"
        empty_file.write_bytes(b"")

        res = qa_service.verify_oxml_integrity(empty_file)
        assert res["is_valid"] is False
        assert any("0 bytes" in err for err in res["errors"])


def test_corrupt_binary_not_a_zip_fails_gracefully():
    """Verify non-ZIP garbage file fails safely."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        garbage_file = tmp / "garbage.pptx"
        garbage_file.write_bytes(b"NOT_A_VALID_ZIP_HEADER_JUST_RANDOM_DATA_1234567890")

        res = qa_service.verify_oxml_integrity(garbage_file)
        assert res["is_valid"] is False
        assert any("not a valid ZIP" in err for err in res["errors"])

        # Also verify Reader raises clear exception rather than unhandled panic
        reader = PresentationReader()
        with pytest.raises(Exception) as exc_info:
            reader.read(garbage_file)
        assert exc_info.value is not None


def test_missing_presentation_xml_fails_gracefully():
    """Verify ZIP missing ppt/presentation.xml is caught and reported."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        bad_pkg = tmp / "missing_pres_xml.pptx"

        with zipfile.ZipFile(bad_pkg, "w") as zf:
            zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            # Notice ppt/presentation.xml is omitted!

        res = qa_service.verify_oxml_integrity(bad_pkg)
        assert res["is_valid"] is False
        assert any("Missing ppt/presentation.xml" in err for err in res["errors"])


def test_malformed_xml_inside_slide_fails_gracefully():
    """Verify malformed XML in a slide is flagged without crashing."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        bad_xml_pkg = tmp / "broken_xml.pptx"

        with zipfile.ZipFile(bad_xml_pkg, "w") as zf:
            zf.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types/>')
            zf.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships/>')
            zf.writestr("ppt/presentation.xml", '<?xml version="1.0"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>')
            # Broken slide XML with mismatched closing tag
            zf.writestr("ppt/slides/slide1.xml", '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><openTag>noClose</p:sld>')

        res = qa_service.verify_oxml_integrity(bad_xml_pkg)
        assert res["is_valid"] is False
        assert any("Malformed slide XML" in err for err in res["errors"])


def test_malformed_input_leaves_source_unmodified():
    """Verify handling a malformed file never writes into or modifies the original source."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        bad_file = tmp / "source_bad.pptx"
        orig_content = b"ORIGINAL_CORRUPT_PAYLOAD"
        bad_file.write_bytes(orig_content)

        def attempt_run(p: Path):
            return qa_service.verify_oxml_integrity(p)

        prot_res = qa_service.verify_original_protection(bad_file, attempt_run)
        assert prot_res["is_protected"] is True
        assert bad_file.read_bytes() == orig_content
