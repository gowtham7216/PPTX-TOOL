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

# -----------------------------------------------------------------------------
# XML Namespaces and Constants (recovered from ConvertPlaceholder_v2)
# -----------------------------------------------------------------------------
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
ADEC_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"
DECORATIVE_EXT_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
PLAIN_TEXTBOX_DEFAULT_SZ = "1800"

NSMAP = {
    "p": P_NS,
    "a": A_NS,
    "mc": MC_NS,
}

CONTENT_PLACEHOLDER_LABEL = "Content Placeholder"
# Types that our engine treats as "content" and standardizes names for.
# We deliberately never rename title, ctrTitle, subTitle, ftr, dt, sldNum.
CONTENT_PH_TYPES = (None, "body", "pic", "tbl", "obj")

NV_TAG_MAP = {
    "sp": "nvSpPr",
    "pic": "nvPicPr",
    "graphicFrame": "nvGraphicFramePr",
    "cxnSp": "nvCxnSpPr",
}

PPR_CHILD_ORDER = [
    "lnSpc", "spcBef", "spcAft",
    "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont",
    "buNone", "buAutoNum", "buChar", "buBlip",
    "tabLst", "defRPr", "extLst",
]

RPR_CHILD_ORDER = [
    "ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
    "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
    "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst",
]


def qn(prefix: str, tag: str) -> str:
    """Format Clark notation {ns}tag for prefix."""
    return f"{{{NSMAP[prefix]}}}{tag}"


def _get_localname(elem: Any) -> str:
    if hasattr(elem, "tag") and isinstance(elem.tag, str):
        return etree.QName(elem).localname
    return ""


def get_nv_container(el: Any) -> Optional[Any]:
    """Retrieve the non-visual property container for a shape."""
    local = _get_localname(el)
    target_tag = NV_TAG_MAP.get(local, "nvSpPr")
    container = el.find(qn("p", target_tag))
    if container is None:
        # Fallback to general search
        for tag in ("nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvCxnSpPr"):
            found = el.find(qn("p", tag))
            if found is not None:
                return found
    return container


def has_ph(el: Any) -> bool:
    """True if shape already contains a native <p:ph> element."""
    nv = get_nv_container(el)
    if nv is None:
        return False
    nvpr = nv.find(qn("p", "nvPr"))
    if nvpr is None:
        return False
    return nvpr.find(qn("p", "ph")) is not None


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
    nv = get_nv_container(el)
    if nv is None:
        return False
    cnvpr = nv.find(qn("p", "cNvPr"))
    if cnvpr is None:
        return False
    extlst = cnvpr.find(qn("a", "extLst"))
    if extlst is None:
        return False
    for ext in extlst.findall(qn("a", "ext")):
        if ext.get("uri") == DECORATIVE_EXT_URI:
            dec = ext.find(f"{{{ADEC_NS}}}decorative")
            if dec is not None and dec.get("val") == "1":
                return True
    return False


def mark_decorative(el: Any) -> None:
    """Mark a shape as decorative in OpenXML metadata."""
    nv = get_nv_container(el)
    if nv is None:
        return
    cnvpr = nv.find(qn("p", "cNvPr"))
    if cnvpr is None:
        return
    extlst = cnvpr.find(qn("a", "extLst"))
    if extlst is None:
        extlst = etree.SubElement(cnvpr, qn("a", "extLst"))
    for ext in extlst.findall(qn("a", "ext")):
        if ext.get("uri") == DECORATIVE_EXT_URI:
            return
    ext = etree.SubElement(extlst, qn("a", "ext"))
    ext.set("uri", DECORATIVE_EXT_URI)
    dec = etree.SubElement(ext, f"{{{ADEC_NS}}}decorative")
    dec.set("val", "1")


def _insert_in_schema_order(parent: Any, new_el: Any, tag: str, order_list: List[str]) -> None:
    """Insert a child element in strict OpenXML schema order."""
    if tag not in order_list:
        parent.append(new_el)
        return
    my_rank = order_list.index(tag)
    insert_at = len(list(parent))
    for i, child in enumerate(parent):
        local = _get_localname(child)
        if local in order_list and order_list.index(local) > my_rank:
            insert_at = i
            break
    parent.insert(insert_at, new_el)


def _ensure_para_child(ppr: Any, tag: str, build_fn: Any) -> None:
    """Ensure a specific <a:pPr> child element exists in schema order."""
    if ppr.find(qn("a", tag)) is not None:
        return
    _insert_in_schema_order(ppr, build_fn(), tag, PPR_CHILD_ORDER)


def freeze_paragraph_properties(sp_el: Any) -> None:
    """
    Freeze paragraph spacing and bullet suppression (<a:buNone/>) so that converted
    placeholders do NOT inherit unintended bullets or spacing from layout/master styles.
    Existing deliberate bullet lists (buChar, buAutoNum) are strictly preserved.
    """
    for p in sp_el.findall(f".//{qn('a','p')}"):
        ppr = p.find(qn("a", "pPr"))
        if ppr is None:
            ppr = etree.Element(qn("a", "pPr"))
            p.insert(0, ppr)

        # A paragraph that carries a REAL bullet/number marker is an intentional list item;
        # leave its formatting completely alone.
        has_real_bullet = (
            ppr.find(qn("a", "buChar")) is not None
            or ppr.find(qn("a", "buAutoNum")) is not None
        )
        if has_real_bullet:
            continue

        if ppr.get("algn") is None:
            ppr.set("algn", "l")

        def _lnspc():
            el = etree.Element(qn("a", "lnSpc"))
            etree.SubElement(el, qn("a", "spcPct")).set("val", "100000")
            return el

        def _spcbef():
            el = etree.Element(qn("a", "spcBef"))
            etree.SubElement(el, qn("a", "spcPts")).set("val", "0")
            return el

        def _spcaft():
            el = etree.Element(qn("a", "spcAft"))
            etree.SubElement(el, qn("a", "spcPts")).set("val", "0")
            return el

        def _bunone():
            return etree.Element(qn("a", "buNone"))

        _ensure_para_child(ppr, "lnSpc", _lnspc)
        _ensure_para_child(ppr, "spcBef", _spcbef)
        _ensure_para_child(ppr, "spcAft", _spcaft)
        # Explicitly say "no bullet" so this paragraph can never inherit an unwanted bullet
        _ensure_para_child(ppr, "buNone", _bunone)

        sp_lststyle = sp_el.find(f".//{qn('a','lstStyle')}")
        if sp_lststyle is not None and len(sp_lststyle) == 0:
            lvl1 = etree.SubElement(sp_lststyle, qn("a", "lvl1pPr"))
            lvl1.set("algn", "l")
            etree.SubElement(etree.SubElement(lvl1, qn("a", "lnSpc")), qn("a", "spcPct")).set("val", "100000")
            etree.SubElement(etree.SubElement(lvl1, qn("a", "spcBef")), qn("a", "spcPts")).set("val", "0")


def backfill_run_sizes(sp_el: Any) -> None:
    """
    Ensure runs have explicit size <a:rPr sz="..."> so converting to a placeholder
    does not cause text to resize to layout master defaults.
    """
    for p in sp_el.findall(f".//{qn('a','p')}"):
        ppr = p.find(qn("a", "pPr"))
        para_defrpr = ppr.find(qn("a", "defRPr")) if ppr is not None else None
        para_sz = para_defrpr.get("sz") if para_defrpr is not None else None
        fallback_sz = para_sz if para_sz else PLAIN_TEXTBOX_DEFAULT_SZ

        for r in p.findall(qn("a", "r")):
            rpr = r.find(qn("a", "rPr"))
            if rpr is None:
                rpr = etree.Element(qn("a", "rPr"))
                r.insert(0, rpr)
            if rpr.get("sz") is None:
                rpr.set("sz", fallback_sz)

        end_rpr = p.find(qn("a", "endParaRPr"))
        if end_rpr is not None and end_rpr.get("sz") is None:
            end_rpr.set("sz", fallback_sz)


def freeze_autofit(el: Any) -> None:
    """Freeze autofit to <a:noAutofit/> preventing unexpected text resizing."""
    if _get_localname(el) != "sp":
        return
    bodypr = el.find(f".//{qn('a','bodyPr')}")
    if bodypr is None:
        return
    autofit = bodypr.find(qn("a", "spAutoFit"))
    if autofit is not None:
        idx = list(bodypr).index(autofit)
        bodypr.remove(autofit)
        bodypr.insert(idx, etree.Element(qn("a", "noAutofit")))


def reset_autofit_scale(sp_el: Any) -> None:
    """Remove font scale reductions from normal autofit."""
    bodypr = sp_el.find(f".//{qn('a','bodyPr')}")
    if bodypr is None:
        return
    autofit = bodypr.find(qn("a", "normAutofit"))
    if autofit is not None and autofit.get("fontScale") is not None:
        del autofit.attrib["fontScale"]
        if autofit.get("lnSpcReduction") is not None:
            del autofit.attrib["lnSpcReduction"]


def rename_shape_if_content(el: Any, ph_type: Optional[str], idx: Optional[str]) -> None:
    """
    Rename shape to 'Content Placeholder {idx}' in <p:cNvPr> so Reading Order / Selection
    panes show standard names. Title/subtitle/footer/slide-number are left untouched.
    """
    if ph_type not in CONTENT_PH_TYPES:
        return
    nv = get_nv_container(el)
    if nv is None:
        return
    cnvpr = nv.find(qn("p", "cNvPr"))
    if cnvpr is not None:
        label_idx = idx or cnvpr.get("id", "1")
        cnvpr.set("name", f"{CONTENT_PLACEHOLDER_LABEL} {label_idx}")


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
        if ph_type not in CONTENT_PH_TYPES:
            continue
        cnvpr = nv.find(qn("p", "cNvPr"))
        if cnvpr is None:
            continue
        current_name = cnvpr.get("name", "")
        if "Placeholder" in current_name:
            continue
        shape_id = cnvpr.get("id", "")
        idx = ph.get("idx", shape_id)
        new_name = f"{CONTENT_PLACEHOLDER_LABEL} {idx}"
        cnvpr.set("name", new_name)
        renamed.append((shape_id, new_name))

    return renamed


@dataclass(frozen=True)
class PlaceholderInsertionResult:
    inserted: bool
    placeholder_type: Optional[str]
    placeholder_idx: Optional[str]
    reason: str


class PlaceholderEngineV2:
    """
    Deterministic placeholder insertion and style preservation engine.
    Directly incorporates the recovered ConvertPlaceholder_v2 logic into SlideGuard V2.
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
        Adds native placeholder metadata to eligible shape without detaching or rewriting
        existing placeholders.
        """
        if placeholder_type not in self.SUPPORTED_TYPES:
            raise ValueError(f"Unsupported placeholder type: {placeholder_type}")

        # 1. Existing placeholder check: NEVER overwrite or detach existing placeholder
        nv = get_nv_container(shape_elem)
        if nv is not None:
            nvpr = nv.find(qn("p", "nvPr"))
            if nvpr is not None:
                existing_ph = nvpr.find(qn("p", "ph"))
                if existing_ph is not None:
                    return PlaceholderInsertionResult(
                        inserted=False,
                        placeholder_type=existing_ph.get("type", "obj"),
                        placeholder_idx=existing_ph.get("idx"),
                        reason="existing-placeholder-preserved",
                    )

        # 2. Existing decorative mark check: leave alone
        if has_existing_decorative_mark(shape_elem):
            return PlaceholderInsertionResult(
                inserted=False,
                placeholder_type=None,
                placeholder_idx=None,
                reason="decorative-shape-preserved",
            )

        # 3. Obtain or create <p:nvPr> container
        if nv is None:
            local = _get_localname(shape_elem)
            target_tag = NV_TAG_MAP.get(local, "nvSpPr")
            nv = etree.SubElement(shape_elem, qn("p", target_tag))
            nvpr = etree.SubElement(nv, qn("p", "nvPr"))
        else:
            nvpr = nv.find(qn("p", "nvPr"))
            if nvpr is None:
                nvpr = etree.SubElement(nv, qn("p", "nvPr"))

        # 4. Allocate non-colliding placeholder index
        idx: Optional[str] = None
        # Title placeholders conventionally omit idx.
        if placeholder_type not in {"title", "ctrTitle"}:
            used = self._used_indices(sp_tree) if sp_tree is not None else set()
            if preferred_idx is not None and preferred_idx.isdigit() and int(preferred_idx) not in used:
                idx = str(int(preferred_idx))
            else:
                candidate = 1
                while candidate in used:
                    candidate += 1
                idx = str(candidate)

        # 5. Insert <p:ph> element at beginning of <p:nvPr>
        attrs: Dict[str, str] = {}
        if placeholder_type in ("body", "pic", "tbl", "obj", "chart", "subTitle"):
            attrs["type"] = placeholder_type
        elif placeholder_type in ("title", "ctrTitle"):
            attrs["type"] = placeholder_type

        if idx is not None:
            attrs["idx"] = idx

        ph_el = etree.Element(qn("p", "ph"), attrib=attrs)
        nvpr.insert(0, ph_el)

        # 6. Apply style and bullet preservation if requested (text shapes)
        if freeze_styles and _get_localname(shape_elem) in ("sp", "cxnSp"):
            freeze_paragraph_properties(shape_elem)
            freeze_autofit(shape_elem)
            backfill_run_sizes(shape_elem)

        # 7. Normalize content placeholder label if requested
        if rename_content:
            rename_shape_if_content(shape_elem, placeholder_type, idx)

        return PlaceholderInsertionResult(True, placeholder_type, idx, "placeholder-inserted")

    @staticmethod
    def _xpath(element: Any, expression: str) -> List[Any]:
        try:
            return element.xpath(expression, namespaces=NSMAP)
        except TypeError:
            return element.xpath(expression)

    @classmethod
    def _used_indices(cls, sp_tree: Any) -> Set[int]:
        """Collect all existing numeric placeholder idx values from shape tree."""
        used: Set[int] = set()
        if sp_tree is None:
            return used
        try:
            nodes = sp_tree.xpath(".//p:ph[@idx]", namespaces=NSMAP)
        except TypeError:
            nodes = sp_tree.xpath(".//p:ph[@idx]")
        for node in nodes:
            value = node.get("idx", "")
            if value.isdigit():
                used.add(int(value))
        return used
