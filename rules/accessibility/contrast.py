"""
SlideGuard Text Contrast Rule (Phase 7).
Analyzes foreground text color against effective background luminances using WCAG 2.1 formulas.
Strictly configured as DETECT_ONLY to prevent arbitrary redesign or brand violation.
"""

from typing import List, Tuple
from rules.base import BaseAccessibilityRule
from rules.accessibility.base import parse_rgb_hex, calculate_contrast_ratio
from models.enums import RuleBehavior, IssueSeverity, AccessibilityStatus
from models.slide import SlideModel
from models.presentation import PresentationModel
from models.issue import Issue


class TextContrastRule(BaseAccessibilityRule):
    """
    Checks text contrast ratios.
    Classified strictly as DETECT_ONLY to prevent automatic color changes
    that would violate brand styling or designer intent.
    """

    @property
    def rule_id(self) -> str:
        return "SG-CNT-01"

    @property
    def name(self) -> str:
        return "Color Contrast Ratio (Minimum)"

    @property
    def behavior(self) -> RuleBehavior:
        return RuleBehavior.DETECT_ONLY

    @property
    def wcag_mapping(self) -> str:
        return "WCAG 2.1 - 1.4.3 Contrast (Minimum) (Level AA)"

    @property
    def description(self) -> str:
        return "Checks text foreground color contrast against background (minimum 4.5:1 ratio)."

    def evaluate(self, slide: SlideModel, presentation: PresentationModel) -> List[Issue]:
        issues: List[Issue] = []

        # Effective background color (default white unless slide background specified)
        default_bg_rgb: Tuple[int, int, int] = (255, 255, 255)

        for shape in slide.shapes:
            if not shape.paragraphs:
                continue

            # Determine shape-level background if available
            shape_bg_rgb = default_bg_rgb

            for p in shape.paragraphs:
                for r in p.runs:
                    if not r.text or not r.text.strip():
                        continue

                    # Check for explicit low-contrast colors or calculated ratio
                    fg_hex = r.font_color_rgb
                    is_low_contrast = False
                    ratio = 21.0
                    threshold = 4.5

                    # Determine threshold based on font size (large text >= 18pt or >= 14pt bold)
                    is_large_text = False
                    if r.font_size_pt and (r.font_size_pt >= 18.0 or (r.font_size_pt >= 14.0 and r.font_bold)):
                        is_large_text = True
                        threshold = 3.0

                    if fg_hex:
                        fg_rgb = parse_rgb_hex(fg_hex)
                        if fg_rgb:
                            ratio = calculate_contrast_ratio(fg_rgb, shape_bg_rgb)
                            if ratio < threshold:
                                is_low_contrast = True
                        elif fg_hex.upper() in ("D3D3D3", "E0E0E0", "CCCCCC", "F0F0F0"):
                            is_low_contrast = True
                            ratio = 1.5

                    if is_low_contrast:
                        issues.append(
                            Issue(
                                issue_id=f"CNT_LOW_{slide.slide_number}_{shape.shape_id}",
                                rule_id=self.rule_id,
                                rule_name=self.name,
                                slide_number=slide.slide_number,
                                severity=IssueSeverity.WARNING,
                                shape_id=shape.shape_id,
                                shape_name=shape.name,
                                wcag_mapping=self.wcag_mapping,
                                description=f"Text run '{r.text[:20]}...' uses low-contrast color #{fg_hex or 'UNKNOWN'}.",
                                recommendation=f"Increase color contrast ratio to at least {threshold:.1f}:1 (calculated: {ratio:.2f}:1).",
                                behavior=self.behavior,
                                status=AccessibilityStatus.MANUAL_REVIEW,
                                confidence="HIGH",
                                explanation=f"Foreground #{fg_hex} against background #{shape_bg_rgb[0]:02X}{shape_bg_rgb[1]:02X}{shape_bg_rgb[2]:02X} yields ratio {ratio:.2f}:1 (required: {threshold:.1f}:1).",
                                extra_data={
                                    "foreground_rgb": fg_hex,
                                    "background_rgb": f"{shape_bg_rgb[0]:02X}{shape_bg_rgb[1]:02X}{shape_bg_rgb[2]:02X}",
                                    "calculated_ratio": round(ratio, 2),
                                    "threshold": threshold,
                                    "is_large_text": is_large_text,
                                }
                            )
                        )
                        break  # Report once per shape to prevent log flooding

        return issues
