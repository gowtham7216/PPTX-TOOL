"""
SlideGuard OOXML Package.
"""

from pptx_engine.ooxml.xml_namespaces import NAMESPACES
from pptx_engine.ooxml.xml_utils import (
    find_elements_by_xpath,
    get_shape_alt_text,
    get_shape_decorative_status,
    get_element_attribute,
)
from pptx_engine.ooxml.deep_inspector import OOXMLDeepInspector

__all__ = [
    "NAMESPACES",
    "find_elements_by_xpath",
    "get_shape_alt_text",
    "get_shape_decorative_status",
    "get_element_attribute",
    "OOXMLDeepInspector",
]
