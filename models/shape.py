"""
SlideGuard Shape Data Model.
Encapsulates metadata, geometric bounding box, content, and styling for PowerPoint shapes.
"""

from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any
from models.enums import ShapeType, PlaceholderType, ListType


@dataclass
class TextRunModel:
    """Represents a text run within a paragraph."""
    text: str = ""
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    font_bold: Optional[bool] = None
    font_italic: Optional[bool] = None
    font_color_rgb: Optional[str] = None  # Hex format RRGGBB


@dataclass
class TextParagraphModel:
    """Represents a paragraph in a text frame."""
    text: str = ""
    level: int = 0
    list_type: ListType = ListType.NONE
    runs: List[TextRunModel] = field(default_factory=list)


@dataclass
class ShapeModel:
    """
    Unified representation of an individual slide object.
    Preserves exact geometry and visual references without modifying original OOXML.
    """
    shape_id: int
    name: str
    shape_type: ShapeType
    left_pt: float
    top_pt: float
    width_pt: float
    height_pt: float
    z_order: int = 0
    parent_group_id: Optional[int] = None
    is_placeholder: bool = False
    placeholder_type: PlaceholderType = PlaceholderType.NONE
    text_content: str = ""
    paragraphs: List[TextParagraphModel] = field(default_factory=list)
    alt_text: Optional[str] = None
    is_decorative: bool = False
    is_hidden: bool = False
    table_row_count: int = 0
    table_col_count: int = 0
    children: List["ShapeModel"] = field(default_factory=list)  # For group shapes
    raw_element: Any = None  # Reference to lxml/python-pptx element for low-level inspection

    @property
    def right_pt(self) -> float:
        return self.left_pt + self.width_pt

    @property
    def bottom_pt(self) -> float:
        return self.top_pt + self.height_pt

    @property
    def has_text(self) -> bool:
        return bool(self.text_content and self.text_content.strip())
