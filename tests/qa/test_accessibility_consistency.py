"""
SlideGuard QA Tests: Accessibility Consistency & Rule Determinism (Phase 10).
Validates that:
- All 10 Phase 7 accessibility rules execute deterministically:
  SG-ALT-01, SG-DEC-01, SG-TTL-01, SG-TBL-01, SG-CNT-01,
  SG-RDB-01, SG-TXT-01, SG-LNG-01, SG-META-01, SG-RO-VAL-01
- Repeated runs on the same input yield 100% consistent issue counts and quality scores.
- SlideGuard internal quality score is distinctly distinguished from formal WCAG certification.
"""

import tempfile
from pathlib import Path
import pytest

from engine.analyzer.presentation_analyzer import PresentationAnalyzer
from pptx_engine.reader.presentation_reader import PresentationReader
from tests.fixtures.sample_generator import create_comprehensive_test_pptx


EXPECTED_RULE_IDS = [
    "SG-ALT-01",
    "SG-DEC-01",
    "SG-TTL-01",
    "SG-TBL-01",
    "SG-CNT-01",
    "SG-RDB-01",
    "SG-TXT-01",
    "SG-LNG-01",
    "SG-META-01",
    "SG-RO-VAL-01",
]


def test_all_phase_7_rules_evaluated():
    """Verify all 10 Phase 7 rules are active and evaluated in analyzer."""
    analyzer = PresentationAnalyzer()
    reader = PresentationReader()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        sample_pptx = tmp / "rule_test.pptx"
        create_comprehensive_test_pptx(sample_pptx)

        pres = reader.read(sample_pptx)
        results = analyzer.analyze(pres)

        evaluated_rule_ids = [r.rule_id for r in results]
        for expected_id in EXPECTED_RULE_IDS:
            assert expected_id in evaluated_rule_ids, f"Rule {expected_id} missing from evaluation!"


def test_deterministic_scoring_across_repeated_evaluations():
    """Verify that analyzing the exact same presentation yields identical score and issues."""
    analyzer = PresentationAnalyzer()
    reader = PresentationReader()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        sample_pptx = tmp / "deterministic_test.pptx"
        create_comprehensive_test_pptx(sample_pptx)

        pres = reader.read(sample_pptx)

        # Run 1
        res1 = analyzer.analyze(pres)
        m1 = analyzer.compute_metrics(res1, total_slides=pres.total_slides)

        # Run 2
        res2 = analyzer.analyze(pres)
        m2 = analyzer.compute_metrics(res2, total_slides=pres.total_slides)

        # Run 3
        res3 = analyzer.analyze(pres)
        m3 = analyzer.compute_metrics(res3, total_slides=pres.total_slides)

        assert m1.total_issues == m2.total_issues == m3.total_issues
        assert m1.critical_issues == m2.critical_issues == m3.critical_issues
        assert m1.accessibility_score == m2.accessibility_score == m3.accessibility_score
        assert m1.quality_score_v2 == m2.quality_score_v2 == m3.quality_score_v2


def test_internal_quality_score_not_represented_as_wcag_certification():
    """Verify score is reported as internal quality metric and does not claim WCAG certification."""
    analyzer = PresentationAnalyzer()
    reader = PresentationReader()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        sample_pptx = tmp / "cert_test.pptx"
        create_comprehensive_test_pptx(sample_pptx)

        pres = reader.read(sample_pptx)
        res = analyzer.analyze(pres)
        metrics = analyzer.compute_metrics(res, total_slides=pres.total_slides)

        d = metrics.to_dict()
        assert "wcag_certified" not in d
        assert "wcag_certification" not in d
        assert "accessibility_score" in d
        assert "quality_score_v2" in d
