"""
SlideGuard Image In-Place OOXML Transformer (Task 3).
Applies approved alternative text and decorative accessibility metadata directly to existing
image shape XML nodes (<p:cNvPr>).
Strictly guarantees:
  - In-place modification ONLY (no object creation, deletion, or duplication)
  - Zero geometry drift (x/y/w/h/rot unchanged)
  - Zero relationship or image binary mutation
  - Shape ID stability
"""

import logging
from typing import Dict, Any, List, Optional, Set, Tuple
from lxml import etree

from rules.placeholders.classifier import _safe_xpath
from rules.accessibility.base import get_cnvpr_element, safe_xpath
from rules.alt_text.models import ImageDecision, ImageAction, ImageClassification

logger = logging.getLogger(__name__)

DECORATIVE_EXT_URI = "{63B4BB66-AC50-487E-9519-01E37F264883}"
A16_NS = "http://schemas.microsoft.com/office/drawing/2014/main"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


class ImageAltTextTransformer:
    """Deterministically transforms existing image shapes in-place in OOXML."""

    def apply_decisions(
        self,
        raw_slide: Any,
        decisions: List[ImageDecision],
        approved_only: bool = True
    ) -> Dict[str, Any]:
        """
        Applies a list of image accessibility decisions to a slide.
        If approved_only is True, applies only decisions with status in ('APPROVED', 'EDITED').
        """
        slide_elem = raw_slide._element if hasattr(raw_slide, "_element") else raw_slide
        sp_tree = slide_elem.xpath(".//p:spTree")
        if not sp_tree:
            return {"applied_count": 0, "failed_count": 0, "actions": []}

        sp_tree_elem = sp_tree[0]
        applied = 0
        failed = 0
        actions = []

        for dec in decisions:
            if approved_only and dec.status not in ("APPROVED", "EDITED"):
                logger.debug("Skipping unapproved decision for shape %d (status=%s)", dec.shape_id, dec.status)
                continue

            target_id = dec.shape_id
            found = False

            # Search spTree (including groups) for target shape
            for child in sp_tree_elem.iter():
                cnvpr = get_cnvpr_element(child)
                if cnvpr is not None and cnvpr.get("id") == str(target_id):
                    found = True
                    try:
                        action_taken = self._apply_to_cnvpr(cnvpr, dec)
                        applied += 1
                        dec.status = "APPLIED"
                        actions.append({
                            "shape_id": target_id,
                            "action": action_taken,
                            "final_alt": dec.effective_alt_text,
                            "success": True
                        })
                    except Exception as err:
                        logger.error("Error applying image decision to shape %d: %s", target_id, err)
                        failed += 1
                        dec.status = "FAILED"
                        actions.append({
                            "shape_id": target_id,
                            "error": str(err),
                            "success": False
                        })
                    break

            if not found:
                logger.warning("Target shape ID %d not found on slide %d", target_id, dec.slide_number)
                failed += 1
                dec.status = "FAILED"

        return {
            "applied_count": applied,
            "failed_count": failed,
            "actions": actions
        }

    def _apply_to_cnvpr(self, cnvpr: etree._Element, dec: ImageDecision) -> str:
        """Applies alt text or decorative metadata directly to the p:cNvPr element."""
        if dec.action == ImageAction.MARK_DECORATIVE or dec.classification == ImageClassification.DECORATIVE:
            # Mark Decorative: set empty descr and add a16:decorative extension
            cnvpr.set("descr", "")
            self._ensure_decorative_extension(cnvpr)
            return "MARKED_DECORATIVE"

        elif dec.action == ImageAction.GENERATE_ALT_TEXT or (dec.action == ImageAction.PRESERVE and dec.effective_alt_text):
            final_alt = dec.effective_alt_text.strip()
            cnvpr.set("descr", final_alt)
            # Remove decorative extension if present
            self._remove_decorative_extension(cnvpr)
            return "APPLIED_ALT_TEXT"

        elif dec.action == ImageAction.PRESERVE:
            return "PRESERVED_IN_PLACE"

        else:
            return "MANUAL_REVIEW_UNMODIFIED"

    def _ensure_decorative_extension(self, cnvpr: etree._Element):
        """Ensures <a16:decorative val='1'/> exists in <p:cNvPr>/<a:extLst>."""
        ext_lst_nodes = _safe_xpath(cnvpr, "./a:extLst")
        if not ext_lst_nodes:
            ext_lst = etree.SubElement(cnvpr, f"{{{DRAWING_NS}}}extLst")
        else:
            ext_lst = ext_lst_nodes[0]

        # Check existing decorative extension
        existing_ext = None
        for ext in ext_lst:
            if ext.get("uri") == DECORATIVE_EXT_URI:
                existing_ext = ext
                break

        if existing_ext is None:
            existing_ext = etree.SubElement(ext_lst, f"{{{DRAWING_NS}}}ext")
            existing_ext.set("uri", DECORATIVE_EXT_URI)

        # Check a16:decorative inside extension
        dec_nodes = safe_xpath(existing_ext, ".//*[local-name()='decorative']")
        if not dec_nodes:
            dec_elem = etree.SubElement(existing_ext, f"{{{A16_NS}}}decorative")
            dec_elem.set("val", "1")
        else:
            dec_nodes[0].set("val", "1")

    def _remove_decorative_extension(self, cnvpr: etree._Element):
        """Removes decorative extension if shape is given meaningful alt text."""
        ext_lst_nodes = _safe_xpath(cnvpr, "./a:extLst")
        if not ext_lst_nodes:
            return
        ext_lst = ext_lst_nodes[0]
        to_remove = []
        for ext in ext_lst:
            if ext.get("uri") == DECORATIVE_EXT_URI:
                to_remove.append(ext)
        for ext in to_remove:
            ext_lst.remove(ext)
        if len(ext_lst) == 0:
            cnvpr.remove(ext_lst)
