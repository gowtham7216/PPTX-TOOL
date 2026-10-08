"""
SlideGuard OOXML Utility Functions.
Direct lxml operations for low-level PowerPoint inspection and manipulation.
"""

from typing import Optional, List
from lxml import etree
from pptx_engine.ooxml.xml_namespaces import NAMESPACES, P_CNVPR


def find_elements_by_xpath(element, xpath_expr: str) -> List[etree._Element]:
    """Evaluates an xpath expression with predefined OpenXML namespaces."""
    if element is None:
        return []
    try:
        # python-pptx oxml elements pre-bind OpenXML prefixes
        return element.xpath(xpath_expr)
    except (TypeError, etree.XPathError):
        # Fallback for native lxml etree elements requiring explicit namespaces
        try:
            return element.xpath(xpath_expr, namespaces=NAMESPACES)
        except Exception:
            return []
    except Exception:
        try:
            return element.xpath(xpath_expr, namespaces=NAMESPACES)
        except Exception:
            return []


def get_shape_alt_text(shape_element: etree._Element) -> Optional[str]:
    """
    Extracts alternative text directly from p:cNvPr descr or title attribute.
    """
    if shape_element is None:
        return None
    cnv_prs = find_elements_by_xpath(shape_element, ".//p:cNvPr")
    if cnv_prs:
        descr = cnv_prs[0].get("descr")
        if descr and descr.strip():
            return descr.strip()
        title = cnv_prs[0].get("title")
        if title and title.strip():
            return title.strip()
    return None


def get_shape_decorative_status(shape_element: etree._Element) -> bool:
    """
    Checks if shape is marked as decorative in OpenXML metadata.
    Often represented as empty/decorative alt text or specific ext tags.
    """
    if shape_element is None:
        return False
    cnv_prs = find_elements_by_xpath(shape_element, ".//p:cNvPr")
    if cnv_prs:
        # Check decorative attribute if present (Office 365 extension or empty descr marker)
        if cnv_prs[0].get("decorative") == "1":
            return True
        # Check decorative extLst
        exts = find_elements_by_xpath(cnv_prs[0], ".//a:extLst//a:ext")
        for ext in exts:
            if "decorative" in etree.tostring(ext, encoding="unicode").lower():
                return True
    return False


def get_element_attribute(element: etree._Element, attr_name: str) -> Optional[str]:
    """Safely retrieves an attribute value from an XML element."""
    if element is None:
        return None
    return element.get(attr_name)
