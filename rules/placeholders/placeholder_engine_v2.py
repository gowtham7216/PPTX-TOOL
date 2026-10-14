"""
Deterministic placeholder conversion engine incorporating the recovered ConvertPlaceholder_v2 logic.

This module provides native OpenXML placeholder insertion and styling isolation:
1. Adds native placeholder metadata (<p:ph>) to eligible loose shapes.
2. Allocates non-colliding placeholder indices avoiding duplicate idx values.
3. Preserves existing native placeholders (title, subtitle, footer, date, slide-number)
   and existing decorative markings without detaching or rewriting them.
4. Freezes paragraph properties (<a:buNone/>, 100% line spacing, zero paragraph spacing)
   to prevent converted shapes from inheriting unwanted list bullets from layout/master styles.
5. Freezes autofit (<a:noAutofit/>) and backfills run sizes (<a:rPr sz="...">) to preserve
   exact typography and layout fidelity.
6. Normalizes content placeholder names to "Content Placeholder {idx}" for PowerPoint
   Selection Pane and Reading Order screen-reader accessibility.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Set, List, Tuple, Dict, Any
from lxml import etree

import sys
from pathlib import Path

try:
    import ConvertPlaceholder_v2 as cp_v2
except ImportError:
    root_dir = str(Path(__file__).resolve().parents[2])
    if root_dir not in sys.path:
        sys.path.insert(0, root_dir)
    import ConvertPlaceholder_v2 as cp_v2

# -----------------------------------------------------------------------------
# XML Namespaces and Constants (sourced directly from ConvertPlaceholder_v2)
# -----------------------------------------------------------------------------
P_NS = cp_v2.P_NS
A_NS = cp_v2.A_NS
MC_NS = cp_v2.MC_NS
ADEC_NS = cp_v2.ADEC_NS
DECORATIVE_EXT_URI = cp_v2.DECORATIVE_EXT_URI
PLAIN_TEXTBOX_DEFAULT_SZ = cp_v2.PLAIN_TEXTBOX_DEFAULT_SZ

NSMAP = cp_v2.NS

CONTENT_PLACEHOLDER_LABEL = cp_v2.CONTENT_PLACEHOLDER_LABEL
CONTENT_PH_TYPES = cp_v2.CONTENT_PH_TYPES
NV_TAG_MAP = cp_v2.NV_TAG_MAP
PPR_CHILD_ORDER = cp_v2.PPR_CHILD_ORDER
RPR_CHILD_ORDER = cp_v2.RPR_CHILD_ORDER


def qn(prefix: str, tag: str) -> str:
    """Format Clark notation {ns}tag for prefix."""
    return cp_v2.qn(prefix, tag)


def _get_localname(elem: Any) -> str:
    if hasattr(elem, "tag") and isinstance(elem.tag, str):
        return etree.QName(elem).localname
    return ""


def get_nv_container(el: Any) -> Optional[Any]:
    """Retrieve the non-visual property container for a shape."""
    return cp_v2.get_nv_container(el)


def has_ph(el: Any) -> bool:
    """True if shape already contains a native <p:ph> element."""
    return cp_v2.has_ph(el)


def get_existing_ph(el: Any) -> Optional[Any]:
    """Retrieve existing <p:ph> element if present."""
    nv = get_nv_container(el)
    if nv is None:
        return None
    nvpr = nv.find(qn("p", "nvPr"))
    if nvpr is None:
        return None
    return nvpr.find(qn("p", "ph"))


def has_existing_decorative_mark(el: Any) -> bool:
    """
    True if this shape was already marked Decorative via Office's native extension.
    These shapes must be left alone.
    """
    if el is None:
        return False
    tag = etree.QName(el).localname if hasattr(el, "tag") else ""
    if tag not in ("sp", "pic", "graphicFrame", "cxnSp"):
        return False
    return cp_v2.has_existing_decorative_mark(el)


def mark_decorative(el: Any) -> None:
    """Mark a shape as decorative in OpenXML metadata."""
    cp_v2.mark_decorative(el)


def normalize_content_placeholder_names(sp_tree: Any) -> List[Tuple[str, str]]:
    """
    Post-processing name normalization pass: ensure any content-type placeholder
    has 'Placeholder' in its name. Returns list of (shape_id, new_name) renames.
    """
    renamed: List[Tuple[str, str]] = []
    if sp_tree is None:
        return renamed

    for child in list(sp_tree):
        if not has_ph(child):
            continue
        nv = get_nv_container(child)
        if nv is None:
            continue
        nvpr = nv.find(qn("p", "nvPr"))
        ph = nvpr.find(qn("p", "ph")) if nvpr is not None else None
        if ph is None:
            continue
        ph_type = ph.get("type")
        cnvpr = nv.find(qn("p", "cNvPr"))
        if cnvpr is None:
            continue
        current_name = cnvpr.get("name", "")
        shape_id = cnvpr.get("id", "")
        idx = ph.get("idx", shape_id)

        if ph_type in ("title", "ctrTitle"):
            if "Title" not in current_name:
                new_name = "Title 1"
                cnvpr.set("name", new_name)
                renamed.append((shape_id, new_name))
        elif ph_type == "subTitle":
            if "Subtitle" not in current_name:
                new_name = f"Subtitle {idx}"
                cnvpr.set("name", new_name)
                renamed.append((shape_id, new_name))
        elif ph_type in CONTENT_PH_TYPES:
            if "Placeholder" not in current_name:
                new_name = f"{CONTENT_PLACEHOLDER_LABEL} {idx}"
                cnvpr.set("name", new_name)
                renamed.append((shape_id, new_name))

    return renamed


def resolve_shape_font(sp_el: Any, primary_font: str = "Times New Roman", fallback_font: str = "Calibri") -> str:
    """
    Standardizes font to primary_font (Times New Roman) by default.
    If applying primary_font causes contents to overflow the placeholder box, falls back to fallback_font (Calibri).
    """
    return cp_v2.resolve_shape_font(sp_el, primary_font=primary_font, fallback_font=fallback_font)


def preserve_and_backfill_run_fonts(
    sp_el: Any,
    fallback_font: Optional[str] = None,
    standardize_fonts: bool = True
) -> str:
    """
    Standardizes font typeface according to design rules:
    - Primary standard: Times New Roman
    - Fallback standard: Calibri if text content overflows placeholder bounding box
    - If standardize_fonts is False: preserves whichever font was originally applied.
    """
    if sp_el is None:
        return ""

    if standardize_fonts:
        target_font = resolve_shape_font(sp_el, primary_font="Times New Roman", fallback_font="Calibri")
    else:
        # Detect shape-level dominant font
        target_font = fallback_font
        for p in sp_el.findall(f'.//{cp_v2.qn("a", "p")}'):
            ppr = p.find(cp_v2.qn("a", "pPr"))
            if ppr is not None:
                def_rpr = ppr.find(cp_v2.qn("a", "defRPr"))
                if def_rpr is not None:
                    lat = def_rpr.find(cp_v2.qn("a", "latin"))
                    if lat is not None and lat.get("typeface"):
                        target_font = lat.get("typeface")
                        break
            for r in p.findall(cp_v2.qn("a", "r")):
                rpr = r.find(cp_v2.qn("a", "rPr"))
                if rpr is not None:
                    lat = rpr.find(cp_v2.qn("a", "latin"))
                    if lat is not None and lat.get("typeface"):
                        target_font = lat.get("typeface")
                        break
            if target_font:
                break
        if not target_font:
            target_font = "Times New Roman"

    # Apply target font to all runs and paragraph ends
    for p in sp_el.findall(f'.//{cp_v2.qn("a", "p")}'):
        for r in p.findall(cp_v2.qn("a", "r")):
            rpr = r.find(cp_v2.qn("a", "rPr"))
            if rpr is None:
                rpr = etree.Element(cp_v2.qn("a", "rPr"))
                rpr.set("lang", "en-US")
                r.insert(0, rpr)
            cp_v2.set_latin_font(rpr, target_font)

        end_rpr = p.find(cp_v2.qn("a", "endParaRPr"))
        if end_rpr is not None:
            cp_v2.set_latin_font(end_rpr, target_font)

    return target_font


@dataclass(frozen=True)
class PlaceholderInsertionResult:
    inserted: bool
    placeholder_type: Optional[str]
    placeholder_idx: Optional[str]
    reason: str


class PlaceholderEngineV2:
    """
    Deterministic placeholder insertion and style preservation engine interface.
    Implementation removed awaiting ConvertPlaceholder_v2 integration.
    """

    SUPPORTED_TYPES = {"title", "ctrTitle", "subTitle", "body", "obj", "pic", "tbl", "chart"}

    def ensure_placeholder(
        self,
        shape_elem: Any,
        placeholder_type: str,
        sp_tree: Optional[Any] = None,
        preferred_idx: Optional[str] = None,
        freeze_styles: bool = True,
        rename_content: bool = True,
    ) -> PlaceholderInsertionResult:
        """
        Adds native placeholder metadata to eligible shape using ConvertPlaceholder_v2 logic.
        """
        if placeholder_type not in self.SUPPORTED_TYPES:
            raise ValueError(f"Unsupported placeholder type: {placeholder_type}")

        # Check existing placeholder
        if cp_v2.has_ph(shape_elem):
            nv = cp_v2.get_nv_container(shape_elem)
            ph = nv.find(cp_v2.qn("p", "nvPr")).find(cp_v2.qn("p", "ph")) if nv is not None else None
            return PlaceholderInsertionResult(
                inserted=False,
                placeholder_type=ph.get("type", "obj") if ph is not None else None,
                placeholder_idx=ph.get("idx") if ph is not None else None,
                reason="existing-placeholder-preserved",
            )

        # Check existing decorative mark
        if cp_v2.has_existing_decorative_mark(shape_elem):
            return PlaceholderInsertionResult(
                inserted=False,
                placeholder_type=None,
                placeholder_idx=None,
                reason="decorative-shape-preserved",
            )

        # Determine non-colliding index
        idx_val = None
        if placeholder_type not in ("title", "ctrTitle"):
            used = set()
            if sp_tree is not None:
                for ph_node in sp_tree.xpath(".//p:ph[@idx]", namespaces=cp_v2.NS):
                    v = ph_node.get("idx")
                    if v and v.isdigit():
                        used.add(int(v))
            if preferred_idx is not None and str(preferred_idx).isdigit() and int(preferred_idx) not in used:
                idx_val = str(int(preferred_idx))
            else:
                cand = 1
                while cand in used:
                    cand += 1
                idx_val = str(cand)

        tag = etree.QName(shape_elem).localname if hasattr(shape_elem, "tag") else ""

        if placeholder_type in ("body", "pic", "tbl", "obj", "chart"):
            if freeze_styles and tag in ("sp", "cxnSp"):
                cp_v2.freeze_paragraph_properties(shape_elem)
                cp_v2.backfill_run_sizes(shape_elem)
                preserve_and_backfill_run_fonts(shape_elem)
            cp_v2.freeze_autofit(shape_elem)
            cp_v2.add_placeholder(shape_elem, placeholder_type, idx_val or "1")
            return PlaceholderInsertionResult(True, placeholder_type, idx_val, "placeholder-inserted")

        elif placeholder_type in ("title", "ctrTitle"):
            if freeze_styles and tag in ("sp", "cxnSp"):
                cp_v2.freeze_paragraph_properties(shape_elem)
                cp_v2.freeze_autofit(shape_elem)
                cp_v2.backfill_run_sizes(shape_elem)
                preserve_and_backfill_run_fonts(shape_elem)
            nv = cp_v2.get_nv_container(shape_elem)
            if nv is not None:
                nvpr = nv.find(cp_v2.qn("p", "nvPr"))
                if nvpr is None:
                    nvpr = etree.SubElement(nv, cp_v2.qn("p", "nvPr"))
                if nvpr.find(cp_v2.qn("p", "ph")) is None:
                    ph = etree.Element(cp_v2.qn("p", "ph"), attrib={"type": placeholder_type})
                    nvpr.insert(0, ph)
                cnvpr = nv.find(cp_v2.qn("p", "cNvPr"))
                if cnvpr is not None:
                    curr_name = cnvpr.get("name", "")
                    if "Title" not in curr_name:
                        cnvpr.set("name", "Title 1")
            return PlaceholderInsertionResult(True, placeholder_type, None, "placeholder-inserted")

        elif placeholder_type == "subTitle":
            if freeze_styles and tag in ("sp", "cxnSp"):
                cp_v2.freeze_paragraph_properties(shape_elem)
                cp_v2.freeze_autofit(shape_elem)
                cp_v2.backfill_run_sizes(shape_elem)
                preserve_and_backfill_run_fonts(shape_elem)
            nv = cp_v2.get_nv_container(shape_elem)
            if nv is not None:
                nvpr = nv.find(cp_v2.qn("p", "nvPr"))
                if nvpr is None:
                    nvpr = etree.SubElement(nv, cp_v2.qn("p", "nvPr"))
                if nvpr.find(cp_v2.qn("p", "ph")) is None:
                    ph = etree.Element(cp_v2.qn("p", "ph"), attrib={"type": "subTitle", "idx": str(idx_val or "1")})
                    nvpr.insert(0, ph)
                cnvpr = nv.find(cp_v2.qn("p", "cNvPr"))
                if cnvpr is not None:
                    curr_name = cnvpr.get("name", "")
                    if "Subtitle" not in curr_name:
                        cnvpr.set("name", f"Subtitle {idx_val or '1'}")
            return PlaceholderInsertionResult(True, placeholder_type, idx_val, "placeholder-inserted")

        return PlaceholderInsertionResult(False, placeholder_type, None, "unsupported-type")
