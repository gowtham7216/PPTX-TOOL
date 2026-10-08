"""
SlideGuard List Transformer.
Executes deterministic in-place OpenXML list transformations:
1. Promotes visual pseudo-bullets into native <a:buChar char="...">.
2. Promotes visual pseudo-numbers/letters into native <a:buAutoNum type="..." startAt="...">.
3. Preserves existing native and inherited list properties without visual redesign.
4. Preserves shape IDs, animations, hyperlinks, fonts, margins, alignments, and text content.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Set, Dict, Any, Tuple
import logging
from lxml import etree

from rules.lists.classifier import (
    ListClassifier,
    ParagraphListAnalysis,
    ListCategory,
    ListStatus,
    _safe_xpath,
    NSMAP,
)

logger = logging.getLogger(__name__)


@dataclass
class ListRemediationResult:
    """Outcome of list remediation across an individual slide."""
    slide_number: int
    success: bool = True
    total_paragraphs_analyzed: int = 0
    paragraphs_remediated: int = 0
    paragraphs_preserved: int = 0
    paragraphs_manual_review: int = 0
    remediated_targets: List[Dict[str, Any]] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


class ListTransformer:
    """Executes safe in-place OpenXML list transformations."""

    def __init__(self, classifier: Optional[ListClassifier] = None):
        self.classifier = classifier or ListClassifier()

    def remediate_slide(
        self,
        slide_elem,
        slide_number: int,
        raw_slide=None,
        approved_targets: Optional[Set[Tuple[int, int]]] = None
    ) -> ListRemediationResult:
        """
        Processes and remediates list semantics on a slide.
        approved_targets: set of (shape_id, paragraph_index) tuples approved for remediation.
        """
        result = ListRemediationResult(slide_number=slide_number, success=True)
        if slide_elem is None:
            result.errors.append("Slide element is None.")
            result.success = False
            return result

        analyses = self.classifier.analyze_slide_lists(
            raw_slide=raw_slide,
            slide_number=slide_number
        )
        result.total_paragraphs_analyzed = len(analyses)

        for analysis in analyses:
            target_key = (analysis.shape_id, analysis.paragraph_index)

            # Check approval gate
            if approved_targets is not None and target_key not in approved_targets:
                result.paragraphs_preserved += 1
                continue

            if analysis.status == ListStatus.VALID or analysis.status == ListStatus.NOT_A_LIST:
                result.paragraphs_preserved += 1
                continue

            if analysis.status == ListStatus.MANUAL_REVIEW:
                result.paragraphs_manual_review += 1
                result.warnings.append(
                    f"Slide {slide_number} Shape {analysis.shape_id} P{analysis.paragraph_index}: {analysis.reason}"
                )
                continue

            if analysis.status == ListStatus.SAFE_TO_REMEDIATE:
                try:
                    self.remediate_paragraph(analysis)
                    result.paragraphs_remediated += 1
                    result.remediated_targets.append({
                        "slide_number": slide_number,
                        "shape_id": analysis.shape_id,
                        "paragraph_index": analysis.paragraph_index,
                        "category": analysis.category.value,
                        "numbering_type": analysis.numbering_type,
                        "bullet_char": analysis.bullet_char,
                    })
                    logger.info(
                        "Remediated list on Slide %d Shape %d P%d (%s)",
                        slide_number, analysis.shape_id, analysis.paragraph_index, analysis.category.value
                    )
                except Exception as e:
                    logger.error(
                        "Failed to remediate list on Slide %d Shape %d P%d: %s",
                        slide_number, analysis.shape_id, analysis.paragraph_index, e, exc_info=True
                    )
                    result.errors.append(f"Shape {analysis.shape_id} P{analysis.paragraph_index}: {str(e)}")
                    result.paragraphs_preserved += 1

        return result

    def remediate_paragraph(self, analysis: ParagraphListAnalysis) -> None:
        """
        Remediates an individual paragraph element in-place:
        1. Ensures <a:pPr> exists.
        2. Adds <a:buChar> or <a:buAutoNum>.
        3. Strips the visual pseudo-prefix from the text runs so numbers/bullets aren't duplicated.
        4. Preserves formatting, fonts, margins, alignments.
        """
        p_elem = analysis.raw_p_element
        if p_elem is None:
            raise ValueError("Missing raw paragraph element for remediation.")

        # Step 1: Ensure <a:pPr> exists
        ppr_elems = _safe_xpath(p_elem, "./a:pPr")
        if not ppr_elems:
            p_pr = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/main}pPr")
            p_elem.insert(0, p_pr)
        else:
            p_pr = ppr_elems[0]

        # Step 2: Remove <a:buNone/> if present
        for bn in _safe_xpath(p_pr, "./a:buNone"):
            p_pr.remove(bn)

        # Step 3: Configure bullet / autonum markup
        if analysis.category == ListCategory.BULLET:
            char_to_use = analysis.bullet_char or "•"
            # Remove any existing buAutoNum
            for an in _safe_xpath(p_pr, "./a:buAutoNum"):
                p_pr.remove(an)

            existing_bc = _safe_xpath(p_pr, "./a:buChar")
            if existing_bc:
                existing_bc[0].set("char", char_to_use)
            else:
                bc = etree.Element(
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar",
                    attrib={"char": char_to_use}
                )
                self._insert_bullet_element(p_pr, bc)

        elif analysis.category in (ListCategory.NUMBERED, ListCategory.LETTERED):
            num_type = analysis.numbering_type or "arabicPeriod"
            start_at = str(analysis.numbering_start or 1)

            # Remove any existing buChar
            for bc in _safe_xpath(p_pr, "./a:buChar"):
                p_pr.remove(bc)

            existing_an = _safe_xpath(p_pr, "./a:buAutoNum")
            if existing_an:
                existing_an[0].set("type", num_type)
                existing_an[0].set("startAt", start_at)
            else:
                an = etree.Element(
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}buAutoNum",
                    attrib={"type": num_type, "startAt": start_at}
                )
                self._insert_bullet_element(p_pr, an)

        # Step 4: Ensure hanging indent if completely missing
        if not p_pr.get("marL") and not p_pr.get("indent"):
            p_pr.set("marL", "457200")    # 0.5 inches
            p_pr.set("indent", "-228600")  # -0.25 inches hanging indent

        # Step 5: Strip visual pseudo-prefix from text runs
        if analysis.pseudo_prefix:
            self._strip_pseudo_prefix_from_runs(p_elem, analysis.pseudo_prefix)

    def _insert_bullet_element(self, p_pr, bullet_elem) -> None:
        """
        Inserts buChar or buAutoNum into <a:pPr> respecting DrawingML schema sequence:
        lnSpc, spcBfr, spcAft, (bullet group), tabLst, defRPr, extLst.
        """
        # Place before tabLst, defRPr, extLst if they exist
        target_idx = len(p_pr)
        for i, child in enumerate(p_pr):
            tag_name = etree.QName(child.tag).localname
            if tag_name in ("tabLst", "defRPr", "extLst"):
                target_idx = i
                break
        p_pr.insert(target_idx, bullet_elem)

    def _strip_pseudo_prefix_from_runs(self, p_elem, prefix: str) -> None:
        """
        Removes the visual prefix string from the initial text runs of the paragraph,
        preserving all font styling (<a:rPr>) and subsequent text intact.
        """
        t_nodes = _safe_xpath(p_elem, ".//a:r/a:t")
        if not t_nodes:
            return

        remaining_to_strip = prefix

        for t_node in t_nodes:
            if not remaining_to_strip:
                break
            text = t_node.text or ""
            if not text:
                continue

            if text.startswith(remaining_to_strip):
                # Run starts with entire remaining prefix
                t_node.text = text[len(remaining_to_strip):]
                remaining_to_strip = ""
            elif remaining_to_strip.startswith(text):
                # Run is completely consumed by prefix
                remaining_to_strip = remaining_to_strip[len(text):]
                t_node.text = ""
            else:
                # Partial match (e.g. whitespace variation)
                # Strip leading whitespace and prefix
                stripped_prefix = remaining_to_strip.strip()
                if text.startswith(stripped_prefix):
                    t_node.text = text[len(stripped_prefix):].lstrip()
                remaining_to_strip = ""
