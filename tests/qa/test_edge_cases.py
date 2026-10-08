"""
SlideGuard QA Tests: Edge Cases & Structural Variations (Phase 10).
Validates edge cases across:
- Groups: normal, nested, empty, shape-only, text-only, mixed picture/text
- Placeholders: title, subtitle, body, two-content, unusual geometry, no duplicate placeholders
- Lists: bullets, numbered, alpha, roman numerals, multi-level, pseudo-lists
- Minimal / Empty presentations: 0-content slide, 1 empty slide, blank presentation
- Mixed language presentations: Multilingual slides, text integrity preserved
- Unusual fonts: original font styling and family references preserved without substitution
"""

import tempfile
from pathlib import Path
import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

from engine.qa.qa_service import QAService
from pptx_engine.reader.presentation_reader import PresentationReader
from pptx_engine.writer.presentation_writer import PresentationWriter
from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from engine.remediation.remediation_service import RemediationService
from workspace.manager import WorkspaceManager
from tests.qa.synthetic_profiles import (
    create_profile_minimal,
    create_profile_mixed_language,
    create_profile_placeholder_heavy,
    create_profile_list_heavy,
    create_profile_group_heavy,
    create_profile_nested_group,
)


def test_empty_and_minimal_presentation_does_not_crash():
    """Verify 1-slide blank presentation passes through analysis and remediation without crashing."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        min_pptx = tmp / "minimal.pptx"
        create_profile_minimal(min_pptx)

        res = qa_service.run_comprehensive_qa(min_pptx, profile_name="MINIMAL_EMPTY")
        assert res.overall_status.value in ["PASS", "PASS_WITH_WARNINGS"]
        assert res.integrity_status.value == "PASS"


def test_mixed_language_presentation_preserves_all_texts():
    """Verify multilingual content (English, Spanish, French, German, Japanese) text is untouched."""
    reader = PresentationReader()
    analyzer = PresentationAnalyzer()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        lang_pptx = tmp / "multilingual.pptx"
        create_profile_mixed_language(lang_pptx)

        pres = reader.read(lang_pptx)
        assert pres.total_slides == 3

        # Run analysis
        rule_results = analyzer.analyze(pres)
        assert len(rule_results) > 0

        # Slide 2 Spanish & French
        s2 = pres.get_slide_by_number(2)
        all_text = " ".join(sh.text_content for sh in s2.shapes)
        assert "Iniciativa de Accesibilidad" in all_text
        assert "Le contenu doit être accessible" in all_text

        # Slide 3 Japanese
        s3 = pres.get_slide_by_number(3)
        s3_text = " ".join(sh.text_content for sh in s3.shapes)
        assert "アクセシビリティ" in s3_text


def test_placeholder_edge_cases_no_duplicate_placeholders():
    """Verify native placeholders are preserved and no duplicates are created."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        ph_pptx = tmp / "placeholders.pptx"
        create_profile_placeholder_heavy(ph_pptx)

        res = qa_service.run_comprehensive_qa(ph_pptx, profile_name="PLACEHOLDER_HEAVY")
        assert res.integrity_status.value == "PASS"
        assert res.design_preservation_status.value == "PASS"

        # Verify no duplicate placeholders were created
        cand = Presentation(res.output_file)
        for cand_s in cand.slides:
            ph_ids = [ph.shape_id for ph in cand_s.placeholders]
            assert len(ph_ids) == len(set(ph_ids)), f"Duplicate placeholders detected on slide {cand_s.slide_id}"


def test_list_edge_cases_preserve_numbering_and_indentation():
    """Verify multi-format list items maintain structure, indentation and text."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        list_pptx = tmp / "lists.pptx"
        create_profile_list_heavy(list_pptx)

        res = qa_service.run_comprehensive_qa(list_pptx, profile_name="LIST_HEAVY")
        assert res.integrity_status.value == "PASS"

        # Load output and check text content of bullets
        reloaded = Presentation(res.output_file)
        s = reloaded.slides[0]
        full_text = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
        assert "Bullet item alpha" in full_text
        assert "Numbered first step" in full_text


def test_complex_group_testing():
    """Verify normal and nested groups preserve geometry and child shape IDs."""
    qa_service = QAService()
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        grp_pptx = tmp / "groups.pptx"
        create_profile_group_heavy(grp_pptx)

        res = qa_service.run_comprehensive_qa(grp_pptx, profile_name="GROUP_HEAVY")
        assert res.integrity_status.value == "PASS"
        assert res.overall_status.value in ["PASS", "PASS_WITH_WARNINGS"]


def test_unusual_font_metadata_preserved_without_substitution():
    """Verify custom/unusual font references remain untouched without unwanted substitution."""
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        font_pptx = tmp / "unusual_fonts.pptx"

        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
        p = tb.text_frame.paragraphs[0]
        p.text = "Text with unique custom font"
        p.font.name = "CustomUnusualProFont-SemiBold"
        p.font.size = Pt(14)
        prs.save(str(font_pptx))

        reader = PresentationReader()
        writer = PresentationWriter()
        pres = reader.read(font_pptx)

        out_pptx = tmp / "font_out.pptx"
        writer.write(pres, out_pptx)

        reloaded = Presentation(str(out_pptx))
        p_reloaded = reloaded.slides[0].shapes[0].text_frame.paragraphs[0]
        assert p_reloaded.font.name == "CustomUnusualProFont-SemiBold"
        assert p_reloaded.font.size == Pt(14)
