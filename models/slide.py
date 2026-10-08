"""
SlideGuard Slide Data Model.
Encapsulates slide hierarchy, title, shapes, layout metadata, and footers.
"""

from dataclasses import dataclass, field
from typing import List, Optional, Any
from models.shape import ShapeModel


@dataclass
class SlideModel:
    """Represents an individual slide within the presentation."""
    slide_id: int
    slide_number: int  # 1-indexed
    title: Optional[str] = None
    title_shape_id: Optional[int] = None
    layout_name: str = "Blank"
    shapes: List[ShapeModel] = field(default_factory=list)
    has_footer: bool = False
    footer_text: Optional[str] = None
    footer_shape_ids: List[int] = field(default_factory=list)
    notes_text: Optional[str] = None
    raw_slide: Any = None  # Reference to python-pptx Slide object

    @property
    def shape_count(self) -> int:
        return len(self.shapes)

    def get_shape_by_id(self, shape_id: int) -> Optional[ShapeModel]:
        """Finds a shape by ID, traversing top-level and nested groups."""
        def _search(shapes: List[ShapeModel]) -> Optional[ShapeModel]:
            for s in shapes:
                if s.shape_id == shape_id:
                    return s
                if s.children:
                    res = _search(s.children)
                    if res:
                        return res
            return None
        return _search(self.shapes)
