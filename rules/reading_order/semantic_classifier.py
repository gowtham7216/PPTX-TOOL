"""
SlideGuard Semantic Role Classifier.
Deterministically classifies shapes into semantic roles (Title, Body, Image, Table, Footer, etc.)
using OpenXML structures, placeholder metadata, text characteristics, and spatial layout.
"""

from enum import Enum
from dataclasses import dataclass, field
from typing import Optional, List, Tuple, Dict, Any
import re
from pptx_engine.utilities.geometry import is_in_footer_region

NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _safe_xpath(elem, xpath_expr: str) -> list:
    """Executes xpath query safely on either python-pptx BaseOxmlElement or raw lxml Element."""
    try:
        return elem.xpath(xpath_expr)
    except TypeError:
        return elem.xpath(xpath_expr, namespaces=NSMAP)
    except Exception:
        try:
            return elem.xpath(xpath_expr, namespaces=NSMAP)
        except Exception:
            return []


# Content headings and structural labels that must NEVER be classified as footers
CONTENT_HEADING_PATTERNS = [
    re.compile(r"^\s*chapter\b", re.IGNORECASE),
    re.compile(r"^\s*section\b", re.IGNORECASE),
    re.compile(r"^\s*unit\b", re.IGNORECASE),
    re.compile(r"^\s*topic\b", re.IGNORECASE),
    re.compile(r"^\s*part\b", re.IGNORECASE),
    re.compile(r"^\s*module\b", re.IGNORECASE),
    re.compile(r"^\s*lesson\b", re.IGNORECASE),
]

# Strongly verified corporate/copyright footer patterns
FOOTER_TEXT_PATTERNS = [
    re.compile(r"^\s*wiley\s*$", re.IGNORECASE),
    re.compile(r"\bwiley\s*&?\s*sons\b", re.IGNORECASE),
    re.compile(r"\bjohn\s+wiley\b", re.IGNORECASE),
    re.compile(r"\bwiley-blackwell\b", re.IGNORECASE),
    re.compile(r"(©|\(c\)|copyright)\s*\d{4}", re.IGNORECASE),
    re.compile(r"\ball\s+rights\s+reserved\b", re.IGNORECASE),
    re.compile(r"^\s*(page\s*)?\d+(\s*(of|/)\s*\d+)?\s*$", re.IGNORECASE),
    re.compile(r"^\s*confidential\s*$", re.IGNORECASE),
    re.compile(r"^\s*draft\s*$", re.IGNORECASE),
    re.compile(r"^\s*internal\s+use\s+only\s*$", re.IGNORECASE),
    re.compile(r"^\s*proprietary\s*$", re.IGNORECASE),
]


class SemanticRole(str, Enum):
    """Semantic role assigned to a slide shape."""
    TITLE = "TITLE"
    SUBTITLE = "SUBTITLE"
    BODY_TEXT = "BODY_TEXT"
    CARD_BADGE = "CARD_BADGE"
    CARD_HEADER = "CARD_HEADER"
    CARD_BODY = "CARD_BODY"
    LIST = "LIST"
    IMAGE = "IMAGE"
    TABLE = "TABLE"
    CHART = "CHART"
    GROUP = "GROUP"
    CAPTION = "CAPTION"
    FOOTER = "FOOTER"
    SLIDE_NUMBER = "SLIDE_NUMBER"
    DECORATIVE = "DECORATIVE"
    UNKNOWN = "UNKNOWN"



@dataclass
class SemanticItem:
    """Classified semantic properties of an individual shape."""
    shape_id: int
    name: str
    role: SemanticRole
    confidence: float  # 0.0 to 1.0
    rationale: str
    is_footer: bool
    is_title: bool
    bbox_pt: Tuple[float, float, float, float]  # (left, top, width, height)
    z_order: int
    text_preview: str = ""
    is_placeholder: bool = False
    placeholder_type: Optional[str] = None
    placeholder_idx: Optional[str] = None
    parent_group_id: Optional[int] = None

    @property
    def left(self) -> float:
        return self.bbox_pt[0]

    @property
    def top(self) -> float:
        return self.bbox_pt[1]

    @property
    def width(self) -> float:
        return self.bbox_pt[2]

    @property
    def height(self) -> float:
        return self.bbox_pt[3]

    @property
    def right(self) -> float:
        return self.left + self.width

    @property
    def bottom(self) -> float:
        return self.top + self.height


class SemanticRoleClassifier:
    """
    Deterministic semantic role classifier for PowerPoint slide shapes.
    Never uses probabilistic cloud AI; relies on verifiable OOXML and geometric rules.
    """

    def __init__(self, footer_threshold_ratio: float = 0.88):
        self.footer_threshold_ratio = footer_threshold_ratio

    def classify_shape(
        self,
        shape_elem,
        shape_obj,
        slide_h_pt: float,
        slide_w_pt: float,
        z_order: int,
        parent_group_id: Optional[int] = None,
        surrounding_items: Optional[List[SemanticItem]] = None
    ) -> SemanticItem:
        """Classifies an individual shape element based on its OOXML and content properties."""
        tag = shape_elem.tag.split("}")[-1]

        # Extract basic ID and Name from p:cNvPr
        cnv_prs = _safe_xpath(shape_elem, ".//p:cNvPr")
        shape_id = int(cnv_prs[0].get("id", "0")) if cnv_prs else 0
        name = cnv_prs[0].get("name", f"Shape_{shape_id}") if cnv_prs else f"Shape_{shape_id}"

        # Placeholder info
        ph_elems = _safe_xpath(shape_elem, ".//p:ph")
        is_placeholder = len(ph_elems) > 0
        ph_type = ph_elems[0].get("type") if is_placeholder else None
        ph_idx = ph_elems[0].get("idx") if is_placeholder else None

        # Text content
        text = ""
        if hasattr(shape_obj, "has_text_frame") and shape_obj.has_text_frame:
            text = shape_obj.text_frame.text.strip()

        # Bounding box in points
        bbox_pt = (0.0, 0.0, 0.0, 0.0)
        if hasattr(shape_obj, "left") and hasattr(shape_obj, "top"):
            try:
                from pptx_engine.utilities.geometry import emu_to_points
                bbox_pt = (
                    emu_to_points(shape_obj.left),
                    emu_to_points(shape_obj.top),
                    emu_to_points(shape_obj.width),
                    emu_to_points(shape_obj.height)
                )
            except Exception:
                pass

        # 1. Check Placeholder Roles (Highest Priority)
        if is_placeholder:
            if ph_type in ("title", "ctrTitle"):
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.TITLE,
                    confidence=1.0,
                    rationale=f"Native title placeholder (type='{ph_type}')",
                    is_footer=False,
                    is_title=True,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=True,
                    placeholder_type=ph_type,
                    placeholder_idx=ph_idx,
                    parent_group_id=parent_group_id
                )
            elif ph_type == "subTitle":
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.SUBTITLE,
                    confidence=0.95,
                    rationale="Native subtitle placeholder",
                    is_footer=False,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=True,
                    placeholder_type=ph_type,
                    placeholder_idx=ph_idx,
                    parent_group_id=parent_group_id
                )
            elif ph_type == "ftr":
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.FOOTER,
                    confidence=1.0,
                    rationale="Native footer placeholder (type='ftr')",
                    is_footer=True,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=True,
                    placeholder_type=ph_type,
                    placeholder_idx=ph_idx,
                    parent_group_id=parent_group_id
                )
            elif ph_type == "sldNum":
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.SLIDE_NUMBER,
                    confidence=1.0,
                    rationale="Native slide number placeholder (type='sldNum')",
                    is_footer=True,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=True,
                    placeholder_type=ph_type,
                    placeholder_idx=ph_idx,
                    parent_group_id=parent_group_id
                )
            elif ph_type == "dt":
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.FOOTER,
                    confidence=1.0,
                    rationale="Native date/time footer placeholder (type='dt')",
                    is_footer=True,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=True,
                    placeholder_type=ph_type,
                    placeholder_idx=ph_idx,
                    parent_group_id=parent_group_id
                )
            elif ph_type in ("body", "obj") or ph_type is None:
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.BODY_TEXT,
                    confidence=0.95,
                    rationale=f"Native body/object placeholder (idx={ph_idx})",
                    is_footer=False,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=True,
                    placeholder_type=ph_type or "obj",
                    placeholder_idx=ph_idx,
                    parent_group_id=parent_group_id
                )

        # 2. Check XML Tag Structure (Groups, Pictures, Tables)
        if tag == "grpSp":
            return SemanticItem(
                shape_id=shape_id,
                name=name,
                role=SemanticRole.GROUP,
                confidence=1.0,
                rationale="Composite group shape (p:grpSp)",
                is_footer=False,
                is_title=False,
                bbox_pt=bbox_pt,
                z_order=z_order,
                text_preview="",
                is_placeholder=False,
                parent_group_id=parent_group_id
            )
        elif tag == "pic":
            return SemanticItem(
                shape_id=shape_id,
                name=name,
                role=SemanticRole.IMAGE,
                confidence=1.0,
                rationale="Picture element (p:pic)",
                is_footer=False,
                is_title=False,
                bbox_pt=bbox_pt,
                z_order=z_order,
                text_preview="",
                is_placeholder=False,
                parent_group_id=parent_group_id
            )
        elif tag == "graphicFrame":
            if _safe_xpath(shape_elem, ".//a:tbl"):
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.TABLE,
                    confidence=1.0,
                    rationale="Table graphic frame (a:tbl)",
                    is_footer=False,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview="",
                    is_placeholder=False,
                    parent_group_id=parent_group_id
                )
            elif _safe_xpath(shape_elem, ".//c:chart"):
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.CHART,
                    confidence=1.0,
                    rationale="Chart graphic frame (c:chart)",
                    is_footer=False,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview="",
                    is_placeholder=False,
                    parent_group_id=parent_group_id
                )

        # Check for full-slide or large background decorative shapes and empty containers
        is_bg = (bbox_pt[2] >= slide_w_pt * 0.85 and bbox_pt[3] >= slide_h_pt * 0.60 and len(text) == 0)
        is_thin_divider = (bbox_pt[2] <= 5.0 or bbox_pt[3] <= 5.0) and len(text) == 0
        is_empty_container = (len(text) == 0 and tag in ("sp", "cxnSp"))
        if is_bg or is_thin_divider or is_empty_container:
            return SemanticItem(
                shape_id=shape_id,
                name=name,
                role=SemanticRole.DECORATIVE,
                confidence=0.95,
                rationale="Visual background, container or divider shape excluded from content sequence",
                is_footer=False,
                is_title=False,
                bbox_pt=bbox_pt,
                z_order=z_order,
                text_preview="",
                is_placeholder=False,
                parent_group_id=parent_group_id
            )

        # Check for Card Badges (e.g., '01', '02', '03' in content region)
        badge_match = re.match(r"^\s*(0?[1-9]|[1-9][0-9]|step\s*\d+|part\s*\d+)\s*$", text, re.IGNORECASE)
        if badge_match and bbox_pt[1] < slide_h_pt * 0.85 and bbox_pt[2] < 120.0 and bbox_pt[3] < 80.0:
            return SemanticItem(
                shape_id=shape_id,
                name=name,
                role=SemanticRole.CARD_BADGE,
                confidence=0.92,
                rationale="Card badge or step indicator in content region",
                is_footer=False,
                is_title=False,
                bbox_pt=bbox_pt,
                z_order=z_order,
                text_preview=text[:60],
                is_placeholder=False,
                parent_group_id=parent_group_id
            )

        # Check if text is an educational heading or structural label that must never be a footer
        is_content_heading = bool(text and any(p.search(text) for p in CONTENT_HEADING_PATTERNS))

        # 3. Check Strongly Verified Footer Patterns & Keywords
        if text and not is_content_heading:
            for pattern in FOOTER_TEXT_PATTERNS:
                if pattern.search(text):
                    is_page_num = bool(re.search(r"^\s*(page\s*)?\d+(\s*(of|/)\s*\d+)?\s*$", text, re.IGNORECASE))
                    # Page numbers ONLY match if actually located in footer region or confirmed by shape name
                    if is_page_num and not is_in_footer_region(bbox_pt[1], bbox_pt[3], slide_h_pt, self.footer_threshold_ratio):
                        if "foot" not in name.lower() and "num" not in name.lower():
                            continue
                    return SemanticItem(
                        shape_id=shape_id,
                        name=name,
                        role=SemanticRole.SLIDE_NUMBER if is_page_num else SemanticRole.FOOTER,
                        confidence=0.95,
                        rationale=f"Matches verified footer/copyright pattern: '{pattern.pattern}'",
                        is_footer=True,
                        is_title=False,
                        bbox_pt=bbox_pt,
                        z_order=z_order,
                        text_preview=text[:60],
                        is_placeholder=False,
                        parent_group_id=parent_group_id
                    )


        # 4. Check Spatial Bottom Footer Candidate (ONLY when combined with strong footer evidence)
        # Position alone is NOT sufficient: real educational content frequently exists in bottom margins.
        if (
            bbox_pt[1] > 0
            and is_in_footer_region(bbox_pt[1], bbox_pt[3], slide_h_pt, self.footer_threshold_ratio)
            and not is_content_heading
        ):
            # Evidence A: Shape name explicitly indicates footer/slide number
            has_footer_name = bool(
                re.search(r"\b(footer|slide\s*num(ber)?|page\s*num(ber)?|ftr|sldnum|datetime|date\s*placeholder)\b", name, re.IGNORECASE)
            )
            # Evidence B: Compact box containing standalone slide number digits (e.g., '14' in bottom corner)
            is_standalone_number = bool(
                text and re.match(r"^\s*\d{1,4}\s*$", text) and bbox_pt[2] < 120.0 and bbox_pt[3] < 50.0
            )

            if has_footer_name or is_standalone_number:
                role = SemanticRole.SLIDE_NUMBER if (is_standalone_number or "num" in name.lower()) else SemanticRole.FOOTER
                rationale = (
                    "Spatial footer candidate confirmed by explicit footer element name"
                    if has_footer_name
                    else "Spatial footer candidate confirmed by standalone page number in footer margin"
                )
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=role,
                    confidence=0.90,
                    rationale=rationale,
                    is_footer=True,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=False,
                    parent_group_id=parent_group_id
                )

        # 5. Check Non-Placeholder Title Candidates
        # Standalone text box near top of slide with short single-paragraph text and large font
        if bbox_pt[1] > 0 and bbox_pt[1] < slide_h_pt * 0.22 and text and len(text) < 120:
            # If no other shape above it
            if not any(item.is_title for item in (surrounding_items or [])):
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.TITLE,
                    confidence=0.80,
                    rationale="Text box positioned in header region acting as slide title",
                    is_footer=False,
                    is_title=True,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=False,
                    parent_group_id=parent_group_id
                )

        # 6. Check for Lists
        if hasattr(shape_obj, "has_text_frame") and shape_obj.has_text_frame:
            tf = shape_obj.text_frame
            has_list = any(p.level > 0 for p in getattr(tf, "paragraphs", [])) or _safe_xpath(shape_elem, ".//a:buChar") or _safe_xpath(shape_elem, ".//a:buAutoNum")
            if has_list:
                return SemanticItem(
                    shape_id=shape_id,
                    name=name,
                    role=SemanticRole.LIST,
                    confidence=0.90,
                    rationale="Text frame contains structured list markup (bullets or numbering)",
                    is_footer=False,
                    is_title=False,
                    bbox_pt=bbox_pt,
                    z_order=z_order,
                    text_preview=text[:60],
                    is_placeholder=False,
                    parent_group_id=parent_group_id
                )

        # 7. Check for Image Caption (short text placed immediately below or above an image)
        if text and len(text) < 100 and surrounding_items:
            for item in surrounding_items:
                if item.role == SemanticRole.IMAGE:
                    # Directly below image: top is within 20pt of image.bottom
                    if 0 <= (bbox_pt[1] - item.bottom) <= 24.0 and abs(bbox_pt[0] - item.left) < 100.0:
                        return SemanticItem(
                            shape_id=shape_id,
                            name=name,
                            role=SemanticRole.CAPTION,
                            confidence=0.85,
                            rationale=f"Caption positioned immediately adjacent to image '{item.name}'",
                            is_footer=False,
                            is_title=False,
                            bbox_pt=bbox_pt,
                            z_order=z_order,
                            text_preview=text[:60],
                            is_placeholder=False,
                            parent_group_id=parent_group_id
                        )

        # 8. Decorative shapes (empty thin dividers, accent bars)
        if not text and (bbox_pt[2] <= 5.0 or bbox_pt[3] <= 5.0):
            return SemanticItem(
                shape_id=shape_id,
                name=name,
                role=SemanticRole.DECORATIVE,
                confidence=0.80,
                rationale="Non-text narrow line/divider shape",
                is_footer=False,
                is_title=False,
                bbox_pt=bbox_pt,
                z_order=z_order,
                text_preview="",
                is_placeholder=False,
                parent_group_id=parent_group_id
            )

        # Default: Body text or generic object
        if text:
            return SemanticItem(
                shape_id=shape_id,
                name=name,
                role=SemanticRole.BODY_TEXT,
                confidence=0.75,
                rationale="Standard slide body text content",
                is_footer=False,
                is_title=False,
                bbox_pt=bbox_pt,
                z_order=z_order,
                text_preview=text[:60],
                is_placeholder=False,
                parent_group_id=parent_group_id
            )

        return SemanticItem(
            shape_id=shape_id,
            name=name,
            role=SemanticRole.UNKNOWN,
            confidence=0.50,
            rationale="Unclassified shape object",
            is_footer=False,
            is_title=False,
            bbox_pt=bbox_pt,
            z_order=z_order,
            text_preview="",
            is_placeholder=False,
            parent_group_id=parent_group_id
        )
