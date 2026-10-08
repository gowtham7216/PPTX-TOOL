"""
SlideGuard List Inheritance Resolver.
Traverses Paragraph -> Text Body lstStyle -> Layout Placeholder -> Master Placeholder -> Master txStyles
to resolve inherited bullet formatting, numbering schemes, indentations, and buNone suppression.
"""

from dataclasses import dataclass
from typing import Optional, Tuple, Dict, Any, List
import logging
from lxml import etree

logger = logging.getLogger(__name__)

NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _safe_xpath(elem, xpath_expr: str) -> list:
    """Safely executes xpath across python-pptx oxml elements and lxml elements."""
    if elem is None:
        return []
    try:
        return elem.xpath(xpath_expr, namespaces=NSMAP)
    except TypeError:
        return elem.xpath(xpath_expr)
    except Exception:
        try:
            return elem.xpath(xpath_expr)
        except Exception:
            return []


@dataclass
class InheritedListProperties:
    """Resolved list formatting properties for an individual paragraph."""
    has_explicit_bullet: bool = False
    has_bu_none: bool = False
    bullet_char: Optional[str] = None
    autonum_type: Optional[str] = None
    autonum_start_at: int = 1
    bullet_font: Optional[str] = None
    bullet_color: Optional[str] = None
    bullet_size_pct: Optional[int] = None
    mar_l_emu: Optional[int] = None
    indent_emu: Optional[int] = None
    level: int = 0
    inherited_from: str = "NONE"  # "PARAGRAPH", "LST_STYLE", "LAYOUT", "MASTER", "MASTER_STYLE", "NONE"


class ListInheritanceResolver:
    """Resolves hierarchical list semantics across PowerPoint OpenXML structures."""

    def resolve_paragraph_list_properties(
        self,
        p_elem,
        shape_elem=None,
        raw_slide=None
    ) -> InheritedListProperties:
        """
        Determines the effective bullet/numbering properties of a paragraph by checking:
        1. Paragraph's own <a:pPr>
        2. Shape text body <a:lstStyle>
        3. Slide Layout placeholder <a:lstStyle> or <a:pPr>
        4. Slide Master placeholder <a:lstStyle> or <a:pPr>
        5. Slide Master <p:txStyles> (titleStyle, bodyStyle, otherStyle)
        """
        if p_elem is None:
            return InheritedListProperties()

        # Step 1: Check paragraph's own <a:pPr>
        ppr_elems = _safe_xpath(p_elem, "./a:pPr")
        p_pr = ppr_elems[0] if ppr_elems else None
        level = int(p_pr.get("lvl", "0")) if p_pr is not None and p_pr.get("lvl") else 0

        if p_pr is not None:
            props = self._inspect_p_pr(p_pr, source="PARAGRAPH", level=level)
            if props.has_explicit_bullet or props.has_bu_none:
                return props

        # Step 2: Check shape text body <a:lstStyle>
        if shape_elem is not None:
            lst_style = _safe_xpath(shape_elem, ".//p:txBody/a:lstStyle")
            if lst_style:
                lvl_pr = _safe_xpath(lst_style[0], f"./a:lvl{level + 1}pPr")
                if lvl_pr:
                    props = self._inspect_p_pr(lvl_pr[0], source="LST_STYLE", level=level)
                    if props.has_explicit_bullet or props.has_bu_none:
                        return props

        # If shape is missing or raw_slide is missing, return baseline
        if shape_elem is None or raw_slide is None:
            return InheritedListProperties(level=level, inherited_from="NONE")

        # Step 3: Check placeholder inheritance
        ph_elems = _safe_xpath(shape_elem, ".//p:ph")
        ph_type = ph_elems[0].get("type") if ph_elems else None
        ph_idx = ph_elems[0].get("idx") if ph_elems else None
        is_placeholder = bool(ph_elems)

        # 3a. Layout placeholder
        layout = getattr(raw_slide, "slide_layout", None)
        if layout is not None and is_placeholder:
            layout_shape = self._find_matching_layout_placeholder(layout, ph_type, ph_idx)
            if layout_shape is not None and hasattr(layout_shape, "_element"):
                l_elem = layout_shape._element
                # Check layout shape's lstStyle
                l_lst = _safe_xpath(l_elem, ".//p:txBody/a:lstStyle")
                if l_lst:
                    l_lvl_pr = _safe_xpath(l_lst[0], f"./a:lvl{level + 1}pPr")
                    if l_lvl_pr:
                        props = self._inspect_p_pr(l_lvl_pr[0], source="LAYOUT", level=level)
                        if props.has_explicit_bullet or props.has_bu_none:
                            return props

        # 3b. Master placeholder
        master = getattr(layout, "slide_master", None) if layout is not None else None
        if master is not None and is_placeholder:
            master_shape = self._find_matching_layout_placeholder(master, ph_type, ph_idx)
            if master_shape is not None and hasattr(master_shape, "_element"):
                m_elem = master_shape._element
                m_lst = _safe_xpath(m_elem, ".//p:txBody/a:lstStyle")
                if m_lst:
                    m_lvl_pr = _safe_xpath(m_lst[0], f"./a:lvl{level + 1}pPr")
                    if m_lvl_pr:
                        props = self._inspect_p_pr(m_lvl_pr[0], source="MASTER", level=level)
                        if props.has_explicit_bullet or props.has_bu_none:
                            return props

        # Step 4: Slide Master <p:txStyles>
        if master is not None and hasattr(master, "_element"):
            master_elem = master._element
            target_style_tag = "otherStyle"
            if is_placeholder:
                if ph_type in ("title", "ctrTitle"):
                    target_style_tag = "titleStyle"
                elif ph_type in ("body", "obj") or not ph_type:
                    target_style_tag = "bodyStyle"

            tx_style = _safe_xpath(master_elem, f".//p:txStyles/p:{target_style_tag}")
            if tx_style:
                style_lvl_pr = _safe_xpath(tx_style[0], f"./a:lvl{level + 1}pPr")
                if style_lvl_pr:
                    props = self._inspect_p_pr(style_lvl_pr[0], source="MASTER_STYLE", level=level)
                    if props.has_explicit_bullet or props.has_bu_none:
                        return props

        return InheritedListProperties(level=level, inherited_from="NONE")

    def _inspect_p_pr(self, ppr_elem, source: str, level: int) -> InheritedListProperties:
        """Inspects bullet, numbering, and indentation attributes from a pPr or lvlNpPr node."""
        has_bu_none = bool(_safe_xpath(ppr_elem, "./a:buNone"))
        bu_char_elems = _safe_xpath(ppr_elem, "./a:buChar")
        bu_autonum_elems = _safe_xpath(ppr_elem, "./a:buAutoNum")

        bu_char = bu_char_elems[0].get("char") if bu_char_elems else None
        an_type = None
        an_start = 1

        if bu_autonum_elems:
            an_type = bu_autonum_elems[0].get("type")
            start_str = bu_autonum_elems[0].get("startAt")
            if start_str and start_str.isdigit():
                an_start = int(start_str)

        has_explicit = bool(bu_char or an_type)
        if has_bu_none:
            has_explicit = False

        # Fonts, colors, sizes
        font_elems = _safe_xpath(ppr_elem, "./a:buFont")
        bu_font = font_elems[0].get("typeface") if font_elems else None

        clr_elems = _safe_xpath(ppr_elem, "./a:buClr//a:srgbClr")
        bu_clr = clr_elems[0].get("val") if clr_elems else None

        sz_elems = _safe_xpath(ppr_elem, "./a:buSzPct")
        bu_sz = int(sz_elems[0].get("val")) if sz_elems and sz_elems[0].get("val", "").isdigit() else None

        mar_l = int(ppr_elem.get("marL")) if ppr_elem.get("marL", "").isdigit() else None
        indent_str = ppr_elem.get("indent")
        indent = int(indent_str) if indent_str and (indent_str.isdigit() or (indent_str.startswith("-") and indent_str[1:].isdigit())) else None

        return InheritedListProperties(
            has_explicit_bullet=has_explicit,
            has_bu_none=has_bu_none,
            bullet_char=bu_char,
            autonum_type=an_type,
            autonum_start_at=an_start,
            bullet_font=bu_font,
            bullet_color=bu_clr,
            bullet_size_pct=bu_sz,
            mar_l_emu=mar_l,
            indent_emu=indent,
            level=level,
            inherited_from=source
        )

    def _find_matching_layout_placeholder(self, container, target_type: Optional[str], target_idx: Optional[str]):
        """Finds placeholder shape in layout/master matching type or idx."""
        if not hasattr(container, "shapes"):
            return None

        # Pass 1: exact idx match
        if target_idx:
            for s in container.shapes:
                if hasattr(s, "_element"):
                    ph = _safe_xpath(s._element, ".//p:ph")
                    if ph and ph[0].get("idx") == target_idx:
                        return s

        # Pass 2: exact type match
        if target_type:
            for s in container.shapes:
                if hasattr(s, "_element"):
                    ph = _safe_xpath(s._element, ".//p:ph")
                    if ph and ph[0].get("type") == target_type:
                        return s

        return None
