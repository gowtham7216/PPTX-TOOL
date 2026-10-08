"""
Unit tests for SlideGuard Accessibility Rules.
Validates individual rule detection across slides.
"""

from pptx_engine.reader.presentation_reader import PresentationReader
from rules.titles import SlideTitleRule
from rules.reading_order import ReadingOrderRule
from rules.lists import ListRule
from rules.contrast import TextContrastRule
from rules.readability import ReadabilityRule
from rules.metadata import DocumentMetadataRule
from models.enums import RuleStatus


def test_slide_title_rule(sample_presentation_path):
    reader = PresentationReader()
    presentation = reader.read(sample_presentation_path)
    rule = SlideTitleRule()

    # Slide 1 has a title -> should pass
    s1 = presentation.get_slide_by_number(1)
    res1 = rule.analyze_slide(s1, presentation)
    assert res1.status == RuleStatus.PASSED

    # Slide 5 has NO title -> should fail
    s5 = presentation.get_slide_by_number(5)
    res5 = rule.analyze_slide(s5, presentation)
    assert res5.status == RuleStatus.FAILED
    assert len(res5.issues) == 1
    assert "no defined title" in res5.issues[0].description


def test_list_rule(sample_presentation_path):
    reader = PresentationReader()
    presentation = reader.read(sample_presentation_path)
    rule = ListRule()

    # Slide 2 contains pseudo bullet dashes ("- First objective...")
    s2 = presentation.get_slide_by_number(2)
    res2 = rule.analyze_slide(s2, presentation)
    assert res2.has_issues is True
    assert any("simulated list" in i.description for i in res2.issues)


def test_contrast_and_readability_rules(sample_presentation_path):
    reader = PresentationReader()
    presentation = reader.read(sample_presentation_path)

    # Slide 5 has low contrast and empty text box
    s5 = presentation.get_slide_by_number(5)

    contrast_rule = TextContrastRule()
    c_res = contrast_rule.analyze_slide(s5, presentation)
    assert c_res.has_issues is True
    assert any("low-contrast" in i.description for i in c_res.issues)

    readability_rule = ReadabilityRule()
    r_res = readability_rule.analyze_slide(s5, presentation)
    assert r_res.has_issues is True
    # Should flag empty text box and small font size
    assert any("Empty text box" in i.description for i in r_res.issues)
    assert any("small font size" in i.description for i in r_res.issues)


def test_document_metadata_rule(sample_presentation_path):
    reader = PresentationReader()
    presentation = reader.read(sample_presentation_path)
    rule = DocumentMetadataRule()

    # Document has title set in fixture -> should pass
    s1 = presentation.get_slide_by_number(1)
    res = rule.analyze_slide(s1, presentation)
    assert res.status == RuleStatus.PASSED
