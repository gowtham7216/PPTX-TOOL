"""
SlideGuard Placeholder Inheritance Resolver.
Traverses Slide -> Slide Layout -> Slide Master hierarchies to resolve inherited
geometry, layout bounds, formatting, and semantic placeholder types.
"""

from dataclasses import dataclass
from typing import Optional, Tuple, Dict, Any, List
import logging
from lxml import etree

from pptx_engine.utilities.geometry import emu_to_points

logger = logging.getLogger(__name__)

NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _safe_xpath(elem, xpath_expr: str) -> list:
    """Executes xpath safely across python-pptx BaseOxmlElement and lxml elements."""
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
class InheritedProperties:
    """Resolved inherited properties for a placeholder shape."""
    has_slide_xfrm: bool
    inherited_from: str  # "SLIDE", "LAYOUT", "MASTER", "NONE"
    xfrm_emu: Optional[Tuple[int, int, int, int]] = None  # (off_x, off_y, ext_cx, ext_cy)
    xfrm_pt: Optional[Tuple[float, float, float, float]] = None  # (left, top, width, height)
    rotation: int = 0
    flip_h: bool = False
    flip_v: bool = False
    matched_layout_shape_id: Optional[int] = None
    matched_master_shape_id: Optional[int] = None
    font_family: Optional[str] = None
    font_size_pt: Optional[float] = None
    color_val: Optional[str] = None
    placeholder_type_resolved: str = "body"
    placeholder_idx_resolved: Optional[str] = None


class PlaceholderInheritanceResolver:
    """
    Resolves OpenXML placeholder geometry and formatting down the layout and master chain.
    Ensures that shapes with inherited geometry can be materialized without visual drift.
    """

    def resolve_placeholder(
        self,
        shape_elem,
        raw_slide=None,
        layout=None,
        master=None
    ) -> InheritedProperties:
        """
        Resolves complete geometric and stylistic properties of a placeholder element.
        """
        # 1. Inspect placeholder tag on slide element
        ph_elems = _safe_xpath(shape_elem, ".//p:ph")
        ph_type = ph_elems[0].get("type") if ph_elems else None
        ph_idx = ph_elems[0].get("idx") if ph_elems else None

        # 2. Check if shape has explicit <a:xfrm> on slide
        local_xfrm = _safe_xpath(shape_elem, ".//p:spPr/a:xfrm") or _safe_xpath(shape_elem, ".//p:picPr/a:xfrm")
        if not local_xfrm:
            # Also check direct child a:xfrm for graphicFrame
            local_xfrm = _safe_xpath(shape_elem, ".//p:xfrm") or _safe_xpath(shape_elem, ".//a:xfrm")

        if local_xfrm:
            off_x, off_y, ext_cx, ext_cy, rot, flip_h, flip_v = self._parse_xfrm(local_xfrm[0])
            # Slide has explicit geometry
            return InheritedProperties(
                has_slide_xfrm=True,
                inherited_from="SLIDE",
                xfrm_emu=(off_x, off_y, ext_cx, ext_cy),
                xfrm_pt=(emu_to_points(off_x), emu_to_points(off_y), emu_to_points(ext_cx), emu_to_points(ext_cy)),
                rotation=rot,
                flip_h=flip_h,
                flip_v=flip_v,
                placeholder_type_resolved=ph_type or "body",
                placeholder_idx_resolved=ph_idx
            )

        # 3. Shape lacks slide-level <a:xfrm>; resolve via Layout
        resolved_layout = layout
        if resolved_layout is None and raw_slide is not None:
            resolved_layout = getattr(raw_slide, "slide_layout", None)

        layout_shape = None
        matched_layout_id = None
        if resolved_layout is not None:
            layout_shape = self._find_matching_placeholder(
                container=resolved_layout,
                target_idx=ph_idx,
                target_type=ph_type
            )
            if layout_shape is not None:
                matched_layout_id = getattr(layout_shape, "shape_id", None)
                l_xfrm = _safe_xpath(layout_shape._element, ".//p:spPr/a:xfrm") or _safe_xpath(layout_shape._element, ".//a:xfrm")
                if l_xfrm:
                    off_x, off_y, ext_cx, ext_cy, rot, flip_h, flip_v = self._parse_xfrm(l_xfrm[0])
                    font_family, font_size_pt, color_val = self._extract_text_style(layout_shape._element)
                    return InheritedProperties(
                        has_slide_xfrm=False,
                        inherited_from="LAYOUT",
                        xfrm_emu=(off_x, off_y, ext_cx, ext_cy),
                        xfrm_pt=(emu_to_points(off_x), emu_to_points(off_y), emu_to_points(ext_cx), emu_to_points(ext_cy)),
                        rotation=rot,
                        flip_h=flip_h,
                        flip_v=flip_v,
                        matched_layout_shape_id=matched_layout_id,
                        font_family=font_family,
                        font_size_pt=font_size_pt,
                        color_val=color_val,
                        placeholder_type_resolved=ph_type or "body",
                        placeholder_idx_resolved=ph_idx
                    )

        # 4. Resolve via Master if Layout lacked geometry
        resolved_master = master
        if resolved_master is None and resolved_layout is not None:
            resolved_master = getattr(resolved_layout, "slide_master", None)

        if resolved_master is not None:
            master_shape = self._find_matching_placeholder(
                container=resolved_master,
                target_idx=ph_idx,
                target_type=ph_type
            )
            if master_shape is not None:
                matched_master_id = getattr(master_shape, "shape_id", None)
                m_xfrm = _safe_xpath(master_shape._element, ".//p:spPr/a:xfrm") or _safe_xpath(master_shape._element, ".//a:xfrm")
                if m_xfrm:
                    off_x, off_y, ext_cx, ext_cy, rot, flip_h, flip_v = self._parse_xfrm(m_xfrm[0])
                    font_family, font_size_pt, color_val = self._extract_text_style(master_shape._element)
                    return InheritedProperties(
                        has_slide_xfrm=False,
                        inherited_from="MASTER",
                        xfrm_emu=(off_x, off_y, ext_cx, ext_cy),
                        xfrm_pt=(emu_to_points(off_x), emu_to_points(off_y), emu_to_points(ext_cx), emu_to_points(ext_cy)),
                        rotation=rot,
                        flip_h=flip_h,
                        flip_v=flip_v,
                        matched_layout_shape_id=matched_layout_id,
                        matched_master_shape_id=matched_master_id,
                        font_family=font_family,
                        font_size_pt=font_size_pt,
                        color_val=color_val,
                        placeholder_type_resolved=ph_type or "body",
                        placeholder_idx_resolved=ph_idx
                    )

        # Geometry could not be resolved from layout or master
        return InheritedProperties(
            has_slide_xfrm=False,
            inherited_from="NONE",
            xfrm_emu=(0, 0, 0, 0),
            xfrm_pt=(0.0, 0.0, 0.0, 0.0),
            placeholder_type_resolved=ph_type or "body",
            placeholder_idx_resolved=ph_idx
        )

    def _find_matching_placeholder(
        self,
        container,
        target_idx: Optional[str],
        target_type: Optional[str]
    ):
        """Finds matching placeholder in layout or master shapes."""
        if not hasattr(container, "shapes"):
            return None

        # Pass 1: Match by exact idx if idx specified
        if target_idx is not None:
            for s in container.shapes:
                ph_xml = _safe_xpath(s._element, ".//p:ph")
                if ph_xml and ph_xml[0].get("idx") == str(target_idx):
                    return s

        # Pass 2: Match by semantic type
        norm_target_type = self._normalize_type(target_type)
        for s in container.shapes:
            ph_xml = _safe_xpath(s._element, ".//p:ph")
            if ph_xml:
                l_type = ph_xml[0].get("type")
                if self._normalize_type(l_type) == norm_target_type:
                    return s

        # Pass 3: If target is title/ctrTitle, match first title shape
        if norm_target_type == "title":
            for s in container.shapes:
                ph_xml = _safe_xpath(s._element, ".//p:ph")
                if ph_xml and ph_xml[0].get("type") in ("title", "ctrTitle"):
                    return s

        return None

    def _normalize_type(self, ph_type: Optional[str]) -> str:
        """Normalizes placeholder types for inheritance matching."""
        if not ph_type:
            return "body"
        p = ph_type.lower()
        if p in ("title", "ctrtitle"):
            return "title"
        if p in ("body", "obj"):
            return "body"
        return p

    def _parse_xfrm(self, xfrm_elem) -> Tuple[int, int, int, int, int, bool, bool]:
        """Extracts off, ext, rot, and flips from <a:xfrm>."""
        off = _safe_xpath(xfrm_elem, ".//a:off")
        ext = _safe_xpath(xfrm_elem, ".//a:ext")
        off_x = int(off[0].get("x", "0")) if off else 0
        off_y = int(off[0].get("y", "0")) if off else 0
        ext_cx = int(ext[0].get("cx", "0")) if ext else 0
        ext_cy = int(ext[0].get("cy", "0")) if ext else 0

        rot = int(xfrm_elem.get("rot", "0"))
        flip_h = xfrm_elem.get("flipH") in ("1", "true")
        flip_v = xfrm_elem.get("flipV") in ("1", "true")

        return off_x, off_y, ext_cx, ext_cy, rot, flip_h, flip_v

    def _extract_text_style(self, elem) -> Tuple[Optional[str], Optional[float], Optional[str]]:
        """Extracts font family, font size, and color from layout/master text properties."""
        font_family = None
        font_size_pt = None
        color_val = None

        r_prs = _safe_xpath(elem, ".//a:rPr") or _safe_xpath(elem, ".//a:defRPr")
        if r_prs:
            r_pr = r_prs[0]
            # Font size (sz in 100ths of pt)
            sz = r_pr.get("sz")
            if sz:
                try:
                    font_size_pt = float(sz) / 100.0
                except ValueError:
                    pass

            # Typeface
            latin = _safe_xpath(r_pr, ".//a:latin")
            if latin and latin[0].get("typeface"):
                font_family = latin[0].get("typeface")

            # Color
            srgb = _safe_xpath(r_pr, ".//a:srgbClr")
            if srgb and srgb[0].get("val"):
                color_val = srgb[0].get("val")

        return font_family, font_size_pt, color_val
