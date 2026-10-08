"""
SlideGuard Presentation Data Model.
Encapsulates global presentation properties, slide list, dimensions, and metadata.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from pathlib import Path
from models.slide import SlideModel


@dataclass
class DocumentMetadata:
    """Core document properties for presentation."""
    title: Optional[str] = None
    subject: Optional[str] = None
    author: Optional[str] = None
    language: Optional[str] = None
    keywords: Optional[str] = None
    comments: Optional[str] = None
    created: Optional[str] = None
    modified: Optional[str] = None


@dataclass
class PresentationModel:
    """Root model for an analyzed or remediated presentation."""
    file_path: Path
    slide_width_pt: float = 720.0   # Standard 10 inches in points
    slide_height_pt: float = 540.0  # Standard 7.5 inches in points
    slides: List[SlideModel] = field(default_factory=list)
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    raw_presentation: Any = None

    @property
    def total_slides(self) -> int:
        return len(self.slides)

    def get_slide_by_number(self, slide_number: int) -> Optional[SlideModel]:
        for s in self.slides:
            if s.slide_number == slide_number:
                return s
        return None
