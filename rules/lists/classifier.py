"""
SlideGuard List Classifier.
Deterministically classifies PowerPoint paragraphs and shapes into semantic list categories:
BULLET, NUMBERED, LETTERED, MULTI_LEVEL, NONE, INHERITED, AMBIGUOUS, UNSUPPORTED.
Evaluates safe conversion eligibility without modifying visual layout or inventing text.
"""

from enum import Enum
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Any, Set
import re
import logging
from lxml import etree

from rules.lists.inheritance import (
    ListInheritanceResolver,
    InheritedListProperties,
    _safe_xpath,
    NSMAP,
)

logger = logging.getLogger(__name__)

# Standard pseudo-bullet prefixes
PSEUDO_BULLET_SYMBOLS = ["•", "*", "-", "–", "—", "▪", "▫", "◦", "■", "▸", "►", "✓", "✔", "➢", "❖"]
PSEUDO_BULLET_REGEX = re.compile(r"^([\u2022\*\-\u2013\u2014\u25AA\u25AB\u25E6\u25A0\u25B8\u25BA\u2713\u2714\u27A2\u2756]|o)\s+(.*)", re.DOTALL)

# Pseudo-numbering regexes
NUMERIC_PERIOD_REGEX = re.compile(r"^(\d+)\.\s+(.*)", re.DOTALL)
NUMERIC_PAREN_R_REGEX = re.compile(r"^(\d+)\)\s+(.*)", re.DOTALL)
NUMERIC_PAREN_BOTH_REGEX = re.compile(r"^\((\d+)\)\s+(.*)", re.DOTALL)

ALPHA_LC_PERIOD_REGEX = re.compile(r"^([a-z])\.\s+(.*)", re.DOTALL)
ALPHA_LC_PAREN_R_REGEX = re.compile(r"^([a-z])\)\s+(.*)", re.DOTALL)
ALPHA_LC_PAREN_BOTH_REGEX = re.compile(r"^\(([a-z])\)\s+(.*)", re.DOTALL)

ALPHA_UC_PERIOD_REGEX = re.compile(r"^([A-Z])\.\s+(.*)", re.DOTALL)
ALPHA_UC_PAREN_R_REGEX = re.compile(r"^([A-Z])\)\s+(.*)", re.DOTALL)
ALPHA_UC_PAREN_BOTH_REGEX = re.compile(r"^\(([A-Z])\)\s+(.*)", re.DOTALL)

SUPPORTED_AUTONUM_TYPES = {
    "arabicPeriod", "arabicParenR", "arabicParenBoth", "arabicPlain",
    "alphaLcParenBoth", "alphaLcParenR", "alphaLcPeriod",
    "alphaUcParenBoth", "alphaUcParenR", "alphaUcPeriod",
    "romanLcParenBoth", "romanLcParenR", "romanLcPeriod",
    "romanUcParenBoth", "romanUcParenR", "romanUcPeriod",
    "circleNumDbPlain", "circleNumWdBlackPlain", "circleNumWdWhitePlain",
}


class ListCategory(str, Enum):
    """Semantic category of a paragraph list structure."""
    BULLET = "BULLET"
    NUMBERED = "NUMBERED"
    LETTERED = "LETTERED"
    MULTI_LEVEL = "MULTI_LEVEL"
    NONE = "NONE"
    INHERITED = "INHERITED"
    AMBIGUOUS = "AMBIGUOUS"
    UNSUPPORTED = "UNSUPPORTED"


class ListStatus(str, Enum):
    """Validation and remediation status."""
    VALID = "VALID"
    SAFE_TO_REMEDIATE = "SAFE_TO_REMEDIATE"
    MANUAL_REVIEW = "MANUAL_REVIEW"
    NOT_A_LIST = "NOT_A_LIST"


@dataclass
class ParagraphListAnalysis:
    """Detailed semantic assessment of an individual paragraph."""
    slide_number: int
    shape_id: int
    shape_name: str
    paragraph_index: int
    text_preview: str
    category: ListCategory
    status: ListStatus
    level: int = 0
    numbering_type: Optional[str] = None
    numbering_start: int = 1
    bullet_char: Optional[str] = None
    has_bu_none: bool = False
    is_inherited: bool = False
    inherited_from: str = "NONE"
    confidence: str = "HIGH"  # "HIGH", "MEDIUM", "LOW"
    source_info: str = "PARAGRAPH"
    proposed_action: str = "None"
    reason: str = ""
    pseudo_prefix: Optional[str] = None
    pseudo_clean_text: Optional[str] = None
    raw_p_element: Any = None
    raw_shape_element: Any = None

    @property
    def is_safe_to_remediate(self) -> bool:
        return self.status == ListStatus.SAFE_TO_REMEDIATE

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "paragraph_index": self.paragraph_index,
            "text_preview": self.text_preview,
            "category": self.category.value,
            "status": self.status.value,
            "level": self.level,
            "numbering_type": self.numbering_type,
            "numbering_start": self.numbering_start,
            "bullet_char": self.bullet_char,
            "has_bu_none": self.has_bu_none,
            "is_inherited": self.is_inherited,
            "inherited_from": self.inherited_from,
            "confidence": self.confidence,
            "source_info": self.source_info,
            "proposed_action": self.proposed_action,
            "reason": self.reason,
            "pseudo_prefix": self.pseudo_prefix,
        }


class ListClassifier:
    """Classifies PowerPoint text frames and paragraphs into semantic list models."""

    def __init__(self, resolver: Optional[ListInheritanceResolver] = None):
        self.resolver = resolver or ListInheritanceResolver()

    def analyze_slide_lists(
        self,
        raw_slide,
        slide_number: int,
        slide_model=None
    ) -> List[ParagraphListAnalysis]:
        """Analyzes all text shapes on a slide for native, inherited, and pseudo-list semantics."""
        analyses: List[ParagraphListAnalysis] = []
        if raw_slide is None or not hasattr(raw_slide, "shapes"):
            return analyses

        for shape in raw_slide.shapes:
            if not hasattr(shape, "_element"):
                continue
            sh_analyses = self.analyze_shape_lists(
                shape_elem=shape._element,
                slide_number=slide_number,
                raw_slide=raw_slide,
                shape_obj=shape
            )
            analyses.extend(sh_analyses)

        return analyses

    def analyze_shape_lists(
        self,
        shape_elem,
        slide_number: int,
        raw_slide=None,
        shape_obj=None
    ) -> List[ParagraphListAnalysis]:
        """Analyzes paragraphs within an individual shape element."""
        results: List[ParagraphListAnalysis] = []
        p_elems = _safe_xpath(shape_elem, ".//p:txBody/a:p")
        if not p_elems:
            return results

        # Shape identification
        cnv_prs = _safe_xpath(shape_elem, ".//p:cNvPr")
        shape_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
        shape_name = cnv_prs[0].get("name", f"Shape_{shape_id}") if cnv_prs else f"Shape_{shape_id}"

        # Title check: do not turn slide titles into lists
        ph_elems = _safe_xpath(shape_elem, ".//p:ph")
        ph_type = ph_elems[0].get("type") if ph_elems else None
        if ph_type in ("title", "ctrTitle"):
            # Native title placeholder
            for idx, p in enumerate(p_elems):
                text = "".join(t.text or "" for t in _safe_xpath(p, ".//a:t")).strip()
                results.append(
                    ParagraphListAnalysis(
                        slide_number=slide_number,
                        shape_id=shape_id,
                        shape_name=shape_name,
                        paragraph_index=idx,
                        text_preview=text[:60],
                        category=ListCategory.NONE,
                        status=ListStatus.NOT_A_LIST,
                        level=0,
                        reason="Slide title element is exempt from list formatting.",
                        raw_p_element=p,
                        raw_shape_element=shape_elem
                    )
                )
            return results

        # Multi-level detection prep
        levels_present = set()
        for p in p_elems:
            ppr = _safe_xpath(p, "./a:pPr")
            lvl = int(ppr[0].get("lvl", "0")) if ppr and ppr[0].get("lvl") else 0
            levels_present.add(lvl)
        is_shape_multilevel = len(levels_present) > 1

        # Phase 1: First-pass analysis of each paragraph
        raw_analyses = []
        pseudo_count = 0
        for idx, p in enumerate(p_elems):
            analysis = self._classify_single_paragraph(
                p_elem=p,
                shape_elem=shape_elem,
                slide_number=slide_number,
                shape_id=shape_id,
                shape_name=shape_name,
                paragraph_index=idx,
                raw_slide=raw_slide,
                is_shape_multilevel=is_shape_multilevel
            )
            raw_analyses.append(analysis)
            if analysis.pseudo_prefix:
                pseudo_count += 1

        # Phase 2: Contextual multi-paragraph correlation
        # A pseudo-list is confident if >= 2 paragraphs have consistent pseudo-prefixes
        # or if an isolated pseudo-prefix is surrounded by regular non-list sentences
        total_p = len(p_elems)
        for analysis in raw_analyses:
            if analysis.status == ListStatus.SAFE_TO_REMEDIATE:
                if pseudo_count < 2 and total_p > 1:
                    # Single isolated numbered/bulleted paragraph among plain text -> Ambiguous
                    analysis.status = ListStatus.MANUAL_REVIEW
                    analysis.category = ListCategory.AMBIGUOUS
                    analysis.confidence = "LOW"
                    analysis.reason = "Isolated visual prefix in multi-paragraph text requires manual confirmation."
                    analysis.proposed_action = "Review if paragraph was intended as a list item or plain text."

            results.append(analysis)

        return results

    def _classify_single_paragraph(
        self,
        p_elem,
        shape_elem,
        slide_number: int,
        shape_id: int,
        shape_name: str,
        paragraph_index: int,
        raw_slide,
        is_shape_multilevel: bool
    ) -> ParagraphListAnalysis:
        """Classifies an individual paragraph element using inheritance and content inspection."""
        text = "".join(t.text or "" for t in _safe_xpath(p_elem, ".//a:t")).strip()
        text_preview = text[:60]

        ppr = _safe_xpath(p_elem, "./a:pPr")
        level = int(ppr[0].get("lvl", "0")) if ppr and ppr[0].get("lvl") else 0

        # Step 1: Resolve inherited properties
        inherited: InheritedListProperties = self.resolver.resolve_paragraph_list_properties(
            p_elem=p_elem,
            shape_elem=shape_elem,
            raw_slide=raw_slide
        )

        # Case A: buNone explicitly set
        if inherited.has_bu_none:
            # Check if text nevertheless has pseudo-bullet prefix
            prefix, clean_text, p_type, start_val = self._detect_pseudo_prefix(text)
            if prefix:
                return ParagraphListAnalysis(
                    slide_number=slide_number,
                    shape_id=shape_id,
                    shape_name=shape_name,
                    paragraph_index=paragraph_index,
                    text_preview=text_preview,
                    category=ListCategory.NONE,
                    status=ListStatus.MANUAL_REVIEW,
                    level=level,
                    has_bu_none=True,
                    source_info=inherited.inherited_from,
                    reason="Paragraph has visual list prefix but bullet formatting is explicitly suppressed via buNone.",
                    proposed_action="Manual review: verify if buNone suppression was intentional.",
                    pseudo_prefix=prefix,
                    pseudo_clean_text=clean_text,
                    raw_p_element=p_elem,
                    raw_shape_element=shape_elem
                )
            else:
                return ParagraphListAnalysis(
                    slide_number=slide_number,
                    shape_id=shape_id,
                    shape_name=shape_name,
                    paragraph_index=paragraph_index,
                    text_preview=text_preview,
                    category=ListCategory.NONE,
                    status=ListStatus.NOT_A_LIST,
                    level=level,
                    has_bu_none=True,
                    source_info=inherited.inherited_from,
                    reason="Bullet formatting explicitly suppressed via buNone.",
                    proposed_action="None",
                    raw_p_element=p_elem,
                    raw_shape_element=shape_elem
                )

        # Case B: Native or inherited explicit bullet / numbering
        if inherited.has_explicit_bullet:
            cat = ListCategory.BULLET
            if is_shape_multilevel:
                cat = ListCategory.MULTI_LEVEL
            elif inherited.autonum_type:
                if "alpha" in inherited.autonum_type.lower():
                    cat = ListCategory.LETTERED
                else:
                    cat = ListCategory.NUMBERED

            is_inherited = inherited.inherited_from != "PARAGRAPH"
            status = ListStatus.VALID

            # Check if autonum_type is supported
            if inherited.autonum_type and inherited.autonum_type not in SUPPORTED_AUTONUM_TYPES:
                cat = ListCategory.UNSUPPORTED
                status = ListStatus.MANUAL_REVIEW

            return ParagraphListAnalysis(
                slide_number=slide_number,
                shape_id=shape_id,
                shape_name=shape_name,
                paragraph_index=paragraph_index,
                text_preview=text_preview,
                category=cat,
                status=status,
                level=inherited.level,
                numbering_type=inherited.autonum_type,
                numbering_start=inherited.autonum_start_at,
                bullet_char=inherited.bullet_char,
                is_inherited=is_inherited,
                inherited_from=inherited.inherited_from,
                source_info=inherited.inherited_from,
                confidence="HIGH",
                reason=f"Valid native list structure ({inherited.inherited_from}).",
                proposed_action="Preserve list semantics intact.",
                raw_p_element=p_elem,
                raw_shape_element=shape_elem
            )

        # Case C: No native list markup; inspect for pseudo-list prefixes in text
        if not text:
            # Empty paragraph
            return ParagraphListAnalysis(
                slide_number=slide_number,
                shape_id=shape_id,
                shape_name=shape_name,
                paragraph_index=paragraph_index,
                text_preview="",
                category=ListCategory.NONE,
                status=ListStatus.NOT_A_LIST,
                level=level,
                reason="Empty paragraph.",
                proposed_action="None",
                raw_p_element=p_elem,
                raw_shape_element=shape_elem
            )

        prefix, clean_text, p_type, start_val = self._detect_pseudo_prefix(text)
        if prefix:
            if p_type == "BULLET":
                cat = ListCategory.BULLET
                num_type = None
                b_char = prefix.strip()
            elif p_type.startswith("alpha"):
                cat = ListCategory.LETTERED
                num_type = p_type
                b_char = None
            else:
                cat = ListCategory.NUMBERED
                num_type = p_type
                b_char = None

            return ParagraphListAnalysis(
                slide_number=slide_number,
                shape_id=shape_id,
                shape_name=shape_name,
                paragraph_index=paragraph_index,
                text_preview=text_preview,
                category=cat,
                status=ListStatus.SAFE_TO_REMEDIATE,
                level=level,
                numbering_type=num_type,
                numbering_start=start_val,
                bullet_char=b_char,
                confidence="HIGH",
                source_info="PLAIN_TEXT_PREFIX",
                reason=f"Visual pseudo-list prefix '{prefix}' detected without native OpenXML markup.",
                proposed_action=f"Promote to native list ({cat.value}) with prefix removed from text.",
                pseudo_prefix=prefix,
                pseudo_clean_text=clean_text,
                raw_p_element=p_elem,
                raw_shape_element=shape_elem
            )

        # Plain text sentence
        return ParagraphListAnalysis(
            slide_number=slide_number,
            shape_id=shape_id,
            shape_name=shape_name,
            paragraph_index=paragraph_index,
            text_preview=text_preview,
            category=ListCategory.NONE,
            status=ListStatus.NOT_A_LIST,
            level=level,
            confidence="HIGH",
            source_info="NONE",
            reason="Standard unbulleted paragraph.",
            proposed_action="None",
            raw_p_element=p_elem,
            raw_shape_element=shape_elem
        )

    def _detect_pseudo_prefix(self, text: str) -> Tuple[Optional[str], Optional[str], Optional[str], int]:
        """
        Detects if text starts with a visual pseudo-bullet or numbering prefix.
        Returns: (prefix, clean_text, list_type, start_at)
        """
        if not text:
            return None, None, None, 1

        # 1. Pseudo bullet symbol
        m_bullet = PSEUDO_BULLET_REGEX.match(text)
        if m_bullet:
            symbol = m_bullet.group(1)
            rest = m_bullet.group(2)
            prefix = text[:len(text) - len(rest)]
            # Normalize bullet char if simple dash or asterisk
            char = "•" if symbol in ("*", "-", "o") else symbol
            return prefix, rest, "BULLET", 1

        # 2. Numeric Parentheses Both: (1) Text
        m_num_paren_both = NUMERIC_PAREN_BOTH_REGEX.match(text)
        if m_num_paren_both:
            num = int(m_num_paren_both.group(1))
            rest = m_num_paren_both.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "arabicParenBoth", num

        # 3. Numeric Parenthesis Right: 1) Text
        m_num_paren_r = NUMERIC_PAREN_R_REGEX.match(text)
        if m_num_paren_r:
            num = int(m_num_paren_r.group(1))
            rest = m_num_paren_r.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "arabicParenR", num

        # 4. Numeric Period: 1. Text
        m_num_period = NUMERIC_PERIOD_REGEX.match(text)
        if m_num_period:
            num = int(m_num_period.group(1))
            rest = m_num_period.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "arabicPeriod", num

        # 5. Alpha Lowercase Parentheses Both: (a) Text
        m_alpha_lc_paren_both = ALPHA_LC_PAREN_BOTH_REGEX.match(text)
        if m_alpha_lc_paren_both:
            char = m_alpha_lc_paren_both.group(1)
            num = ord(char) - ord('a') + 1
            rest = m_alpha_lc_paren_both.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "alphaLcParenBoth", num

        # 6. Alpha Lowercase Parenthesis Right: a) Text
        m_alpha_lc_paren_r = ALPHA_LC_PAREN_R_REGEX.match(text)
        if m_alpha_lc_paren_r:
            char = m_alpha_lc_paren_r.group(1)
            num = ord(char) - ord('a') + 1
            rest = m_alpha_lc_paren_r.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "alphaLcParenR", num

        # 7. Alpha Lowercase Period: a. Text
        m_alpha_lc_period = ALPHA_LC_PERIOD_REGEX.match(text)
        if m_alpha_lc_period:
            char = m_alpha_lc_period.group(1)
            num = ord(char) - ord('a') + 1
            rest = m_alpha_lc_period.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "alphaLcPeriod", num

        # 8. Alpha Uppercase Parentheses Both: (A) Text
        m_alpha_uc_paren_both = ALPHA_UC_PAREN_BOTH_REGEX.match(text)
        if m_alpha_uc_paren_both:
            char = m_alpha_uc_paren_both.group(1)
            num = ord(char) - ord('A') + 1
            rest = m_alpha_uc_paren_both.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "alphaUcParenBoth", num

        # 9. Alpha Uppercase Parenthesis Right: A) Text
        m_alpha_uc_paren_r = ALPHA_UC_PAREN_R_REGEX.match(text)
        if m_alpha_uc_paren_r:
            char = m_alpha_uc_paren_r.group(1)
            num = ord(char) - ord('A') + 1
            rest = m_alpha_uc_paren_r.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "alphaUcParenR", num

        # 10. Alpha Uppercase Period: A. Text
        m_alpha_uc_period = ALPHA_UC_PERIOD_REGEX.match(text)
        if m_alpha_uc_period:
            char = m_alpha_uc_period.group(1)
            num = ord(char) - ord('A') + 1
            rest = m_alpha_uc_period.group(2)
            prefix = text[:len(text) - len(rest)]
            return prefix, rest, "alphaUcPeriod", num

        return None, None, None, 1
