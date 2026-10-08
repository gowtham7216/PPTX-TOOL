"""
SlideGuard Task 3B: Alt-Text Quality Classification and Semantic Audit.
Implements strict checks for:
  - ALT_TEXT_TOO_GENERIC detection (formulaic title repetition)
  - Generation basis identification (VISUAL_ANALYSIS, STRUCTURAL_CONTEXT, EXISTING_AUTHOR_TEXT, DETERMINISTIC)
  - Quality classification (SPECIFIC, CONTEXTUAL, GENERIC, MANUAL_REVIEW)
"""

import re
from typing import Optional, Tuple
from rules.alt_text.models import GenerationBasis, QualityClass


GENERIC_TITLE_PATTERNS = [
    r"^(illustration\s+for\s+)",
    r"^(visual\s+graphic\s+supporting\s+)",
    r"^(image\s+supporting\s+)",
    r"^(picture\s+related\s+to\s+)",
    r"^(graphic\s+illustrating\s+)",
    r"^(diagram\s+for\s+)",
    r"^(image\s+for\s+)",
    r"^(graphic\s+for\s+)",
]


def is_generic_title_repetition(alt_text: str, title: str) -> bool:
    """
    Returns True if alt text merely restates the slide title with a generic formulaic preamble,
    e.g., 'Illustration for [title]', 'Visual graphic supporting [title]'.
    """
    if not alt_text or not title:
        return False

    norm_alt = alt_text.lower().strip().rstrip(".")
    norm_ttl = title.lower().strip().rstrip(".")

    if not norm_alt or not norm_ttl:
        return False

    # Direct equality
    if norm_alt == norm_ttl:
        return True

    for pat in GENERIC_TITLE_PATTERNS:
        m = re.match(pat, norm_alt)
        if m:
            remainder = norm_alt[m.end():].strip().rstrip(".")
            if remainder == norm_ttl:
                return True
            if norm_ttl in remainder and len(remainder) <= len(norm_ttl) + 5:
                return True

    return False


def classify_alt_text_quality(
    alt_text: str,
    title: str,
    generation_basis: str,
    has_visual_analysis: bool = False,
    is_caption_grounded: bool = False
) -> Tuple[str, str]:
    """
    Determines (generation_basis, quality_class) conforming strictly to Task 3B rules:
      - Do not claim SPECIFIC if visual pixels were not inspected.
      - Formulaic title repetition is classified as GENERIC.
      - Grounded contextual descriptions without visual pixels are CONTEXTUAL.
      - Existing author text and deterministic decorative markers are SPECIFIC.
    """
    basis = generation_basis

    if basis == GenerationBasis.EXISTING_AUTHOR_TEXT.value:
        return basis, QualityClass.SPECIFIC.value

    if basis == GenerationBasis.DETERMINISTIC.value:
        return basis, QualityClass.SPECIFIC.value

    if has_visual_analysis:
        basis = GenerationBasis.VISUAL_ANALYSIS.value
        if alt_text and not is_generic_title_repetition(alt_text, title):
            return basis, QualityClass.SPECIFIC.value
        return basis, QualityClass.CONTEXTUAL.value

    # Structural context
    basis = GenerationBasis.STRUCTURAL_CONTEXT.value

    if not alt_text:
        return basis, QualityClass.MANUAL_REVIEW.value

    if is_generic_title_repetition(alt_text, title):
        return basis, QualityClass.GENERIC.value

    if is_caption_grounded or len(alt_text.split()) >= 4:
        return basis, QualityClass.CONTEXTUAL.value

    return basis, QualityClass.GENERIC.value
