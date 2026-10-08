"""
SlideGuard Accessibility Rule Base and Utilities.
Extends BaseAccessibilityRule with accessibility-specific OOXML helpers,
luminance calculation, and color parsing utilities.
"""

import math
from typing import Optional, Tuple
from lxml import etree

from rules.base import BaseAccessibilityRule
from models.enums import AccessibilityStatus, RuleBehavior, IssueSeverity


def parse_rgb_hex(hex_str: Optional[str]) -> Optional[Tuple[int, int, int]]:
    """Parses a 6-character hex string into an RGB tuple (r, g, b)."""
    if not hex_str:
        return None
    cleaned = hex_str.strip().lstrip("#").upper()
    if len(cleaned) == 6:
        try:
            return (
                int(cleaned[0:2], 16),
                int(cleaned[2:4], 16),
                int(cleaned[4:6], 16),
            )
        except ValueError:
            return None
    return None


def calculate_relative_luminance(rgb: Tuple[int, int, int]) -> float:
    """
    Calculates WCAG 2.1 relative luminance for an RGB tuple.
    L = 0.2126 * R + 0.7152 * G + 0.0722 * B
    """
    def _channel_lum(c_8bit: int) -> float:
        s = c_8bit / 255.0
        if s <= 0.04045:
            return s / 12.92
        else:
            return math.pow((s + 0.055) / 1.055, 2.4)

    r_lum = _channel_lum(rgb[0])
    g_lum = _channel_lum(rgb[1])
    b_lum = _channel_lum(rgb[2])
    return 0.2126 * r_lum + 0.7152 * g_lum + 0.0722 * b_lum


def calculate_contrast_ratio(fg_rgb: Tuple[int, int, int], bg_rgb: Tuple[int, int, int]) -> float:
    """Calculates WCAG contrast ratio between foreground and background colors."""
    l1 = calculate_relative_luminance(fg_rgb)
    l2 = calculate_relative_luminance(bg_rgb)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


NSMAP = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}


def safe_xpath(elem, xpath_expr: str) -> list:
    """Executes xpath query safely across python-pptx BaseOxmlElement and raw lxml Element."""
    if elem is None:
        return []
    try:
        return elem.xpath(xpath_expr)
    except (TypeError, etree.XPathError):
        try:
            return elem.xpath(xpath_expr, namespaces=NSMAP)
        except Exception:
            return []
    except Exception:
        try:
            return elem.xpath(xpath_expr, namespaces=NSMAP)
        except Exception:
            return []


def get_cnvpr_element(shape_element) -> Optional[etree._Element]:
    """Retrieves the non-visual shape properties node (p:cNvPr) from any shape element."""
    matches = safe_xpath(
        shape_element,
        ".//p:nvSpPr/p:cNvPr | .//p:nvPicPr/p:cNvPr | .//p:nvGraphicFramePr/p:cNvPr | .//p:nvGrpSpPr/p:cNvPr",
    )
    if matches:
        return matches[0]
    return None
