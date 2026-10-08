"""
SlideGuard OOXML Namespace Definitions.
Defines standard OpenXML namespaces for PowerPoint PresentationML and DrawingML.
"""

NAMESPACES = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}

# Common QNames
P_SP = f"{{{NAMESPACES['p']}}}sp"
P_GRPSP = f"{{{NAMESPACES['p']}}}grpSp"
P_PIC = f"{{{NAMESPACES['p']}}}pic"
P_GRAPHICFRAME = f"{{{NAMESPACES['p']}}}graphicFrame"
P_CXNSP = f"{{{NAMESPACES['p']}}}cxnSp"
P_NVSPPR = f"{{{NAMESPACES['p']}}}nvSpPr"
P_CNVPR = f"{{{NAMESPACES['p']}}}cNvPr"
P_CNVSPPR = f"{{{NAMESPACES['p']}}}cNvSpPr"
P_SPPR = f"{{{NAMESPACES['p']}}}spPr"
P_TXBODY = f"{{{NAMESPACES['p']}}}txBody"

A_P = f"{{{NAMESPACES['a']}}}p"
A_R = f"{{{NAMESPACES['a']}}}r"
A_T = f"{{{NAMESPACES['a']}}}t"
A_BLIP = f"{{{NAMESPACES['a']}}}blip"
A_EXT = f"{{{NAMESPACES['a']}}}ext"
A_OFF = f"{{{NAMESPACES['a']}}}off"
A_XFRM = f"{{{NAMESPACES['a']}}}xfrm"
