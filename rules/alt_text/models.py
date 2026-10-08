"""
SlideGuard AI-Assisted Alternative Text & Decorative Image Detection Models (Task 3).
Defines structured data classes, enums, decision representations, and audit structures.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, Any, List, Optional, Tuple


class ImageClassification(str, Enum):
    """Semantic accessibility classification of an image."""
    MEANINGFUL = "MEANINGFUL"
    DECORATIVE = "DECORATIVE"
    UNKNOWN = "UNKNOWN"


class ImageAction(str, Enum):
    """Allowed actions for image accessibility remediation."""
    PRESERVE = "PRESERVE"
    GENERATE_ALT_TEXT = "GENERATE_ALT_TEXT"
    MARK_DECORATIVE = "MARK_DECORATIVE"
    MANUAL_REVIEW = "MANUAL_REVIEW"


@dataclass
class ImageCandidate:
    """Detailed structural and contextual extraction of an individual presentation image."""
    slide_number: int
    shape_id: int
    shape_name: str
    image_rel_id: Optional[str] = None
    image_format: str = "unknown"
    image_filename: str = ""
    dimensions_pt: Tuple[float, float] = (0.0, 0.0)  # (width_pt, height_pt)
    bbox_pt: Tuple[float, float, float, float] = (0.0, 0.0, 0.0, 0.0)  # (left_pt, top_pt, width_pt, height_pt)
    z_order: int = 0
    containing_group_id: Optional[int] = None
    group_hierarchy: List[int] = field(default_factory=list)
    nearby_text: List[str] = field(default_factory=list)
    nearby_title: str = ""
    nearby_caption: Optional[str] = None
    surrounding_shapes: List[Dict[str, Any]] = field(default_factory=list)
    has_hyperlink: bool = False
    hyperlink_url: Optional[str] = None
    is_animated: bool = False
    is_already_decorative: bool = False
    existing_alt_text: str = ""
    existing_title: str = ""
    is_generic_alt_text: bool = False
    image_hash: Optional[str] = None
    is_repeated_across_slides: bool = False
    raw_element: Any = None

    def to_context_dict(self) -> Dict[str, Any]:
        """Returns structured minimal privacy-safe context payload for Gemini evaluation."""
        return {
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "image_format": self.image_format,
            "dimensions": f"{self.dimensions_pt[0]:.0f}x{self.dimensions_pt[1]:.0f}pt",
            "position": {
                "left_pt": round(self.bbox_pt[0], 1),
                "top_pt": round(self.bbox_pt[1], 1),
                "width_pt": round(self.bbox_pt[2], 1),
                "height_pt": round(self.bbox_pt[3], 1),
            },
            "slide_title": self.nearby_title,
            "nearby_text": self.nearby_text[:3],
            "nearby_caption": self.nearby_caption or "",
            "is_grouped": bool(self.containing_group_id),
            "group_hierarchy": self.group_hierarchy,
            "has_hyperlink": self.has_hyperlink,
            "is_animated": self.is_animated,
            "is_repeated": self.is_repeated_across_slides,
            "existing_alt_text": self.existing_alt_text,
            "is_generic_alt": self.is_generic_alt_text,
        }


class GenerationBasis(str, Enum):
    VISUAL_ANALYSIS = "VISUAL_ANALYSIS"
    STRUCTURAL_CONTEXT = "STRUCTURAL_CONTEXT"
    EXISTING_AUTHOR_TEXT = "EXISTING_AUTHOR_TEXT"
    DETERMINISTIC = "DETERMINISTIC"


class QualityClass(str, Enum):
    SPECIFIC = "SPECIFIC"
    CONTEXTUAL = "CONTEXTUAL"
    GENERIC = "GENERIC"
    MANUAL_REVIEW = "MANUAL_REVIEW"


@dataclass
class ImageDecision:
    """Structured decision produced by deterministic safety gates and Gemini 3.5."""
    slide_number: int
    shape_id: int
    shape_name: str
    classification: ImageClassification
    action: ImageAction
    alt_text: str = ""
    original_alt_text: str = ""
    user_edited_alt_text: Optional[str] = None
    confidence: float = 0.85
    reason: str = ""
    is_animated: bool = False
    is_grouped: bool = False
    is_decorative_tagged: bool = False
    generation_basis: str = "STRUCTURAL_CONTEXT"  # VISUAL_ANALYSIS, STRUCTURAL_CONTEXT, EXISTING_AUTHOR_TEXT, DETERMINISTIC
    quality_class: str = "MANUAL_REVIEW"          # SPECIFIC, CONTEXTUAL, GENERIC, MANUAL_REVIEW
    status: str = "PENDING_REVIEW"  # PENDING_REVIEW, APPROVED, EDITED, REJECTED, APPLIED
    model_traceability: Dict[str, str] = field(default_factory=lambda: {
        "antigravity_coding_model": "Gemini 3.8 Flash High",
        "slideguard_runtime_model": "Gemini 3.5",
        "runtime_model_identifier": "gemini-3.5"
    })


    @property
    def effective_alt_text(self) -> str:
        """Returns the user-edited alt text if available, otherwise proposed alt text."""
        if self.user_edited_alt_text is not None:
            return self.user_edited_alt_text
        return self.alt_text

    def approve(self):
        self.status = "APPROVED"

    def edit(self, new_text: str):
        self.user_edited_alt_text = new_text
        self.status = "EDITED"

    def reject(self):
        self.status = "REJECTED"

    def mark_decorative_action(self):
        self.classification = ImageClassification.DECORATIVE
        self.action = ImageAction.MARK_DECORATIVE
        self.alt_text = ""
        self.status = "APPROVED"

    def keep_existing_action(self):
        self.action = ImageAction.PRESERVE
        self.effective_alt_text_val = self.original_alt_text
        self.status = "APPROVED"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "slide_number": self.slide_number,
            "shape_id": self.shape_id,
            "shape_name": self.shape_name,
            "classification": self.classification.value,
            "action": self.action.value,
            "alt_text": self.alt_text,
            "original_alt_text": self.original_alt_text,
            "effective_alt_text": self.effective_alt_text,
            "confidence": round(self.confidence, 2),
            "reason": self.reason,
            "is_animated": self.is_animated,
            "is_grouped": self.is_grouped,
            "generation_basis": self.generation_basis,
            "quality_class": self.quality_class,
            "status": self.status,
            "model_traceability": self.model_traceability,
        }

