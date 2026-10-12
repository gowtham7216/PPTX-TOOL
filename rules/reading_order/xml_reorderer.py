"""
SlideGuard XML Reading Order In-Place Reorderer.
Physically reorders existing XML child nodes inside <p:spTree> without deleting or recreating elements.
Preserves all shape IDs, formatting, relationships, and animation targets.
"""

from typing import List, Dict, Any, Tuple, Optional
import logging
from pptx_engine.ooxml.xml_utils import find_elements_by_xpath

logger = logging.getLogger(__name__)


class XMLReadingOrderReorderer:
    """
    Safely reorders elements in <p:spTree> in-place.
    Never constructs new XML tags; moves existing nodes to maintain complete fidelity.
    """

    SHAPE_TAGS = ("sp", "grpSp", "pic", "graphicFrame", "cxnSp")

    def reorder_sp_tree(
        self,
        slide_elem,
        proposed_order: List[int]
    ) -> Tuple[bool, str, List[int]]:
        """
        Reorders top-level shape elements in slide_elem's <p:spTree> to match proposed_order.
        Returns: (success: bool, message: str, resulting_order: List[int])
        """
        sp_trees = slide_elem.xpath(".//p:spTree")
        if not sp_trees:
            return False, "No p:spTree container found on slide.", []

        sp_tree = sp_trees[0]

        # 1. Map existing shape elements by ID
        shape_elements: Dict[int, Any] = {}
        original_shape_elements_list = []

        for child in list(sp_tree):
            tag_name = child.tag.split("}")[-1]
            if tag_name in self.SHAPE_TAGS:
                cnv_prs = child.xpath(".//p:cNvPr")
                if cnv_prs:
                    s_id = int(cnv_prs[0].get("id", "0"))
                    shape_elements[s_id] = child
                    original_shape_elements_list.append((s_id, child))

        orig_ids = [s_id for s_id, _ in original_shape_elements_list]

        if orig_ids == proposed_order:
            logger.info("Current XML order already matches proposed order; no mutation required.")
            return True, "Order already matches proposed sequence.", orig_ids

        # Safely preserve any unlisted shape IDs to ensure zero dropped shapes
        unlisted_ids = [s_id for s_id in orig_ids if s_id not in proposed_order]
        if unlisted_ids:
            logger.info("Preserving unlisted shapes %s at original relative sequence.", unlisted_ids)
        full_proposed_order = [s_id for s_id in proposed_order if s_id in shape_elements] + unlisted_ids

        # 2. Reorder in-place
        # In DrawingML, the first non-shape properties in <p:spTree> (p:nvGrpSpPr, p:grpSpPr)
        # must remain at the top of <p:spTree>.
        # We move each shape in full_proposed_order to the end of sp_tree.
        try:
            for s_id in full_proposed_order:
                if s_id in shape_elements:
                    elem = shape_elements[s_id]
                    # In lxml, appending an existing child node moves it to the end
                    sp_tree.append(elem)

            # 3. Verify resulting sequence
            resulting_order = []
            for child in sp_tree:
                tag_name = child.tag.split("}")[-1]
                if tag_name in self.SHAPE_TAGS:
                    cnv_prs = child.xpath(".//p:cNvPr")
                    if cnv_prs:
                        resulting_order.append(int(cnv_prs[0].get("id", "0")))

            if resulting_order != full_proposed_order:
                return False, f"Reordering mismatch: expected {full_proposed_order}, got {resulting_order}", resulting_order

            logger.info("Successfully reordered <p:spTree> nodes in-place to: %s", resulting_order)
            return True, "Successfully reordered shape tree nodes in-place.", resulting_order

        except Exception as e:
            logger.error("Failed to reorder <p:spTree> elements: %s", e, exc_info=True)
            return False, f"Exception during in-place XML reordering: {str(e)}", orig_ids
