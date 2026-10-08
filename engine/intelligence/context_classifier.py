"""
SlideGuard Context-Aware Issue Classifier (Phase 8).
Combines multi-modal structural evidence across shapes, slides, and layouts
to produce rich contextual classifications rather than isolated rule evaluations.
"""

from typing import Optional, Dict, Any, List
from dataclasses import dataclass, field

from models.enums import ShapeType, PlaceholderType
from models.slide import SlideModel
from models.shape import ShapeModel
from rules.accessibility.base import safe_xpath, get_cnvpr_element


@dataclass
class TitleContext:
    """Contextual classification for slide title analysis."""
    classification: str  # "NATIVE_TITLE", "OBVIOUS_TITLE_CANDIDATE", "AMBIGUOUS_TITLE", "NO_TITLE"
    is_native: bool
    candidate_shape_id: Optional[int] = None
    candidate_text: str = ""
    font_size_pt: Optional[float] = None
    is_bold: bool = False
    top_pt: Optional[float] = None
    text_length: int = 0
    reading_order_rank: int = -1
    confidence_score: float = 0.0
    rationale: str = ""


@dataclass
class ImageContext:
    """Contextual classification for visual and graphical objects."""
    classification: str  # "DECORATIVE_DIVIDER", "DECORATIVE_EXPLICIT", "MEANINGFUL_PICTURE", "COMPLEX_DIAGRAM", "UNKNOWN"
    shape_id: int
    has_alt_text: bool
    alt_text: str = ""
    is_generic_alt: bool = False
    is_divider: bool = False
    width_pt: float = 0.0
    height_pt: float = 0.0
    surrounding_text_count: int = 0
    confidence_score: float = 0.0
    rationale: str = ""


@dataclass
class TextContext:
    """Contextual classification for textual content and typographic legibility."""
    classification: str  # "TITLE_HEADING", "SECTION_HEADING", "BODY_CONTENT", "CAPTION", "FOOTER", "EMPTY_FRAME"
    shape_id: int
    text_length: int
    font_size_pt: Optional[float]
    semantic_role: str
    is_empty: bool
    is_small_font: bool
    confidence_score: float = 0.0
    rationale: str = ""


@dataclass
class TableContext:
    """Contextual classification for tabular structures."""
    classification: str  # "ACCESSIBLE_DATA_TABLE", "REMEDIABLE_HEADER", "AMBIGUOUS_HEADER", "COMPLEX_MERGED_TABLE"
    shape_id: int
    has_designated_header: bool
    row_count: int
    col_count: int
    has_merged_cells: bool
    row_0_complete: bool
    confidence_score: float = 0.0
    rationale: str = ""


class ContextualClassifier:
    """
    Centralized contextual analysis layer.
    Synthesizes geometry, typography, reading flow, and OpenXML metadata
    to provide holistic semantic decisions.
    """

    def analyze_title_context(self, slide: SlideModel) -> TitleContext:
        """Evaluates whether the slide possesses a native title or an eligible top heading."""
        # 1. Native title check
        if slide.title and slide.title.strip():
            return TitleContext(
                classification="NATIVE_TITLE",
                is_native=True,
                candidate_shape_id=slide.title_shape_id,
                candidate_text=slide.title.strip(),
                confidence_score=1.0,
                rationale="Slide possesses a verified native title placeholder."
            )

        # 2. Check for candidate standalone top text boxes
        candidates = []
        for rank, s in enumerate(slide.shapes):
            if s.shape_type in (ShapeType.TEXT_BOX, ShapeType.AUTO_SHAPE) and s.has_text:
                if s.top_pt <= 160.0:
                    text = "".join(r.text for p in s.paragraphs for r in p.runs).strip()
                    chars = len(text)
                    if 1 <= chars <= 140:
                        # Extract typographic dominance
                        max_font_size = 0.0
                        has_bold = False
                        for p in s.paragraphs:
                            for r in p.runs:
                                if r.font_size_pt and r.font_size_pt > max_font_size:
                                    max_font_size = r.font_size_pt
                                if getattr(r, "font_bold", False):
                                    has_bold = True

                        # Score candidate based on position, font prominence, and brevity
                        pos_score = max(0.0, 1.0 - (s.top_pt / 160.0))
                        font_score = min(1.0, max_font_size / 24.0) if max_font_size > 0 else 0.5
                        bold_score = 0.2 if has_bold else 0.0
                        brevity_score = 0.2 if chars <= 60 else 0.1
                        total_score = min(1.0, 0.4 * pos_score + 0.3 * font_score + bold_score + brevity_score)

                        candidates.append((total_score, s, text, max_font_size, has_bold, rank))

        if not candidates:
            return TitleContext(
                classification="NO_TITLE",
                is_native=False,
                confidence_score=0.95,
                rationale="Slide lacks a title and no suitable top text boxes exist."
            )

        # Sort candidates by contextual score
        candidates.sort(key=lambda x: x[0], reverse=True)
        best_score, best_shape, best_text, font_sz, bold, rank = candidates[0]

        # If single dominant candidate or top candidate significantly outperforms others
        is_clear_leader = (len(candidates) == 1) or (best_score - candidates[1][0] >= 0.15)
        if is_clear_leader and best_score >= 0.70:
            return TitleContext(
                classification="OBVIOUS_TITLE_CANDIDATE",
                is_native=False,
                candidate_shape_id=best_shape.shape_id,
                candidate_text=best_text,
                font_size_pt=font_sz if font_sz > 0 else None,
                is_bold=bold,
                top_pt=best_shape.top_pt,
                text_length=len(best_text),
                reading_order_rank=rank,
                confidence_score=round(best_score, 2),
                rationale=f"Shape '{best_shape.name}' near top ({best_shape.top_pt:.1f}pt) with prominent text ({len(best_text)} chars)."
            )
        else:
            return TitleContext(
                classification="AMBIGUOUS_TITLE",
                is_native=False,
                candidate_shape_id=best_shape.shape_id,
                candidate_text=best_text,
                font_size_pt=font_sz if font_sz > 0 else None,
                is_bold=bold,
                top_pt=best_shape.top_pt,
                text_length=len(best_text),
                reading_order_rank=rank,
                confidence_score=round(best_score, 2),
                rationale=f"Multiple competing text candidates near slide top; manual review required."
            )

    def analyze_image_context(self, shape: ShapeModel, slide: SlideModel) -> ImageContext:
        """Evaluates image or graphic frame role and decorative potential."""
        # Check decorative divider geometry
        is_divider = (
            shape.shape_type == ShapeType.AUTO_SHAPE
            and not shape.has_text
            and (shape.height_pt <= 4.0 or shape.width_pt <= 4.0)
        )

        # Inspect raw OOXML properties if available
        alt_descr = shape.alt_text or ""
        is_explicit_dec = shape.is_decorative

        if shape.raw_element is not None:
            cnvpr = get_cnvpr_element(shape.raw_element)
            if cnvpr is not None:
                descr = cnvpr.get("descr")
                if descr is not None:
                    alt_descr = descr
                if descr == "" and not cnvpr.get("title"):
                    is_explicit_dec = True
                dec_nodes = safe_xpath(cnvpr, ".//*[local-name()='decorative']")
                if dec_nodes and dec_nodes[0].get("val") in ("1", "true"):
                    is_explicit_dec = True

        # Check generic alt text
        generic_tokens = {"image", "picture", "photo", "graphic", "figure", "placeholder", "drawing"}
        is_generic = alt_descr.strip().lower() in generic_tokens

        # Surrounding text objects on slide
        text_count = sum(1 for s in slide.shapes if s.has_text)

        if is_divider:
            return ImageContext(
                classification="DECORATIVE_DIVIDER",
                shape_id=shape.shape_id,
                has_alt_text=bool(alt_descr),
                alt_text=alt_descr,
                is_generic_alt=is_generic,
                is_divider=True,
                width_pt=shape.width_pt,
                height_pt=shape.height_pt,
                surrounding_text_count=text_count,
                confidence_score=0.98,
                rationale=f"AutoShape with zero text and narrow line dimension ({shape.width_pt}x{shape.height_pt}pt)."
            )
        elif is_explicit_dec:
            return ImageContext(
                classification="DECORATIVE_EXPLICIT",
                shape_id=shape.shape_id,
                has_alt_text=False,
                alt_text="",
                is_divider=False,
                width_pt=shape.width_pt,
                height_pt=shape.height_pt,
                surrounding_text_count=text_count,
                confidence_score=0.99,
                rationale="Visual object explicitly tagged as decorative in OpenXML metadata."
            )
        elif shape.shape_type == ShapeType.PICTURE or (shape.raw_element is not None and shape.raw_element.tag.endswith("pic")):
            # Distinguish simple picture vs large diagram
            is_large = (shape.width_pt >= 300.0 and shape.height_pt >= 200.0)
            cls_name = "COMPLEX_DIAGRAM" if is_large else "MEANINGFUL_PICTURE"
            conf = 0.95 if bool(alt_descr and not is_generic) else 0.85
            return ImageContext(
                classification=cls_name,
                shape_id=shape.shape_id,
                has_alt_text=bool(alt_descr),
                alt_text=alt_descr,
                is_generic_alt=is_generic,
                is_divider=False,
                width_pt=shape.width_pt,
                height_pt=shape.height_pt,
                surrounding_text_count=text_count,
                confidence_score=conf,
                rationale="Meaningful pictorial content requiring verified descriptive alternative text."
            )
        else:
            return ImageContext(
                classification="UNKNOWN",
                shape_id=shape.shape_id,
                has_alt_text=bool(alt_descr),
                alt_text=alt_descr,
                is_generic_alt=is_generic,
                is_divider=False,
                width_pt=shape.width_pt,
                height_pt=shape.height_pt,
                surrounding_text_count=text_count,
                confidence_score=0.60,
                rationale="Unclassified visual shape."
            )

    def analyze_table_context(self, shape: ShapeModel) -> TableContext:
        """Evaluates table headers, merged cells, and data completeness."""
        if shape.shape_type != ShapeType.TABLE and not (shape.raw_element is not None and bool(safe_xpath(shape.raw_element, ".//a:tbl"))):
            return TableContext(
                classification="NOT_A_TABLE",
                shape_id=shape.shape_id,
                has_designated_header=False,
                row_count=0,
                col_count=0,
                has_merged_cells=False,
                row_0_complete=False,
                confidence_score=1.0,
                rationale="Shape is not a table."
            )

        tbl = None
        if shape.raw_element is not None:
            tbl_nodes = safe_xpath(shape.raw_element, ".//a:tbl")
            if tbl_nodes:
                tbl = tbl_nodes[0]

        if tbl is None:
            return TableContext(
                classification="AMBIGUOUS_HEADER",
                shape_id=shape.shape_id,
                has_designated_header=False,
                row_count=0,
                col_count=0,
                has_merged_cells=False,
                row_0_complete=False,
                confidence_score=0.5,
                rationale="Unable to parse table OpenXML structure."
            )

        tbl_prs = safe_xpath(tbl, "./a:tblPr")
        has_first_row = bool(tbl_prs and tbl_prs[0].get("firstRow") in ("1", "true"))

        rows = safe_xpath(tbl, "./a:tr")
        row_count = len(rows)
        has_merged = False
        row_0_complete = True
        col_count = 0

        if rows:
            row_0_cells = safe_xpath(rows[0], "./a:tc")
            col_count = len(row_0_cells)
            for tc in row_0_cells:
                text = "".join(safe_xpath(tc, ".//a:t/text()")).strip()
                if not text:
                    row_0_complete = False
                if tc.get("gridSpan") or tc.get("rowSpan") or tc.get("hMerge") or tc.get("vMerge"):
                    has_merged = True

            # Also check subsequent rows for merged cells
            for r in rows[1:]:
                for tc in safe_xpath(r, "./a:tc"):
                    if tc.get("gridSpan") or tc.get("rowSpan") or tc.get("hMerge") or tc.get("vMerge"):
                        has_merged = True

        if has_merged:
            return TableContext(
                classification="COMPLEX_MERGED_TABLE",
                shape_id=shape.shape_id,
                has_designated_header=has_first_row,
                row_count=row_count,
                col_count=col_count,
                has_merged_cells=True,
                row_0_complete=row_0_complete,
                confidence_score=0.90,
                rationale="Table contains merged cells (gridSpan/rowSpan), making linear navigation ambiguous."
            )
        elif has_first_row and row_0_complete:
            return TableContext(
                classification="ACCESSIBLE_DATA_TABLE",
                shape_id=shape.shape_id,
                has_designated_header=True,
                row_count=row_count,
                col_count=col_count,
                has_merged_cells=False,
                row_0_complete=True,
                confidence_score=0.98,
                rationale="Table has explicit firstRow markup and all header cells populated."
            )
        elif not has_first_row and row_0_complete:
            return TableContext(
                classification="REMEDIABLE_HEADER",
                shape_id=shape.shape_id,
                has_designated_header=False,
                row_count=row_count,
                col_count=col_count,
                has_merged_cells=False,
                row_0_complete=True,
                confidence_score=0.95,
                rationale="Row 0 contains complete header text but firstRow attribute is absent."
            )
        else:
            return TableContext(
                classification="AMBIGUOUS_HEADER",
                shape_id=shape.shape_id,
                has_designated_header=has_first_row,
                row_count=row_count,
                col_count=col_count,
                has_merged_cells=False,
                row_0_complete=False,
                confidence_score=0.85,
                rationale="Row 0 contains blank cells; header structure requires human confirmation."
            )
