"""
SlideGuard Alternative Text Rule (Phase 7).
Detects images, charts, and graphical objects lacking descriptive alternative text.
Inspects low-level OOXML description attributes (p:cNvPr/@descr, @title)
and flags generic or missing descriptions.
"""

from typing import List, Optional, Set
from lxml import etree

from rules.base import BaseAccessibilityRule
from rules.accessibility.base import get_cnvpr_element, safe_xpath
from models.enums import RuleBehavior, IssueSeverity, ShapeType, AccessibilityStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue
from models.shape import ShapeModel

GENERIC_ALT_TEXTS: Set[str] = {
    "picture",
    "image",
    "graphic",
    "photo",
    "photograph",
    "diagram",
    "figure",
    "placeholder",
    "untitled",
    "drawing",
    "clipart",
    "clip art",
}


class AltTextRule(BaseAccessibilityRule):
    """
    Validates that non-decorative visual elements possess meaningful alternative text.
    Classified as AI_ASSISTED_REVIEW / DETECT_ONLY because descriptive alt text
    must not be generated from unverified guesswork.
    """

    @property
    def rule_id(self) -> str:
        return "SG-ALT-01"

    @property
    def name(self) -> str:
        return "Missing Alternative Text"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.AI_ASSISTED_REVIEW

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.1.1 Non-text Content (Level A)"

    @property
    def description(self) -> str:
        return "Flags meaningful pictures and diagrams missing descriptive alternative text."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        def _check_shapes(shapes: List[ShapeModel]):
            for s in shapes:
                # Pictures, charts, media frames, and shapes with image blips
                is_visual = (
                    s.shape_type in (ShapeType.PICTURE, ShapeType.MEDIA)
                    or (s.raw_element is not None and (
                        s.raw_element.tag.endswith(("pic", "graphicFrame"))
                        or bool(s.raw_element.xpath(".//a:blip"))
                    ))
                )


                if is_visual:
                    alt_descr, alt_title, is_explicit_dec = self._inspect_ooxml_alt(s)

                    # If explicitly marked decorative, it passes WCAG 1.1.1
                    if s.is_decorative or is_explicit_dec:
                        continue

                    # If alt text exists, check if it is generic or informative
                    effective_alt = (alt_descr or alt_title or s.alt_text or "").strip()

                    if not effective_alt:
                        issues.append(
                            Issue(
                                issue_id=f"ALT_MISSING_{slide.slide_number}_{s.shape_id}",
                                rule_id=self.rule_id,
                                rule_name=self.name,
                                slide_number=slide.slide_number,
                                severity=IssueSeverity.ERROR,
                                shape_id=s.shape_id,
                                shape_name=s.name,
                                wcag_mapping=self.wcag_mapping,
                                description=f"Picture '{s.name}' lacks alternative text description.",
                                recommendation="Provide descriptive alternative text or mark as decorative if purely visual.",
                                behavior=self.behavior,
                                status=AccessibilityStatus.DETECTED,
                                confidence="HIGH",
                                explanation="Visual image element has no description or title attribute in OOXML.",
                                extra_data={
                                    "classification": "UNKNOWN",
                                    "existing_descr": alt_descr,
                                    "existing_title": alt_title,
                                }
                            )
                        )
                    else:
                        normalized = effective_alt.lower().strip()
                        if normalized in GENERIC_ALT_TEXTS:
                            issues.append(
                                Issue(
                                    issue_id=f"ALT_GENERIC_{slide.slide_number}_{s.shape_id}",
                                    rule_id=self.rule_id,
                                    rule_name=self.name,
                                    slide_number=slide.slide_number,
                                    severity=IssueSeverity.WARNING,
                                    shape_id=s.shape_id,
                                    shape_name=s.name,
                                    wcag_mapping=self.wcag_mapping,
                                    description=f"Picture '{s.name}' uses generic/uninformative alternative text '{effective_alt}'.",
                                    recommendation="Replace generic placeholder with meaningful description conveying content and purpose.",
                                    behavior=RuleBehavior.DETECT_ONLY,
                                    status=AccessibilityStatus.MANUAL_REVIEW,
                                    confidence="HIGH",
                                    explanation=f"Text '{effective_alt}' matches generic placeholder vocabulary.",
                                    extra_data={
                                        "classification": "MEANINGFUL",
                                        "existing_alt": effective_alt,
                                    }
                                )
                            )

                if s.children:
                    _check_shapes(s.children)

        _check_shapes(slide.shapes)
        return issues

    def _inspect_ooxml_alt(self, shape: ShapeModel) -> tuple[Optional[str], Optional[str], bool]:
        """Inspects raw OOXML for descr, title, and decorative marker tags."""
        if shape.raw_element is None:
            return None, None, False

        cnvpr = get_cnvpr_element(shape.raw_element)
        if cnvpr is None:
            return None, None, False

        descr = cnvpr.get("descr")
        title = cnvpr.get("title")

        # Check modern Office 365 decorative tag in extensions
        is_dec = False
        dec_nodes = safe_xpath(cnvpr, ".//*[local-name()='decorative']")
        if dec_nodes and dec_nodes[0].get("val") in ("1", "true"):
            is_dec = True

        # In OOXML, empty descr="" explicitly set is often used to mark decorative
        if descr == "" and not title:
            is_dec = True

        return descr, title, is_dec
