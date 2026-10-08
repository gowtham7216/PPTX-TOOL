"""
SlideGuard Image Candidate Extractor (Task 3).
Extracts all image/picture elements from slides, including grouped and nested elements,
with complete contextual, geometric, and relationship metadata.
Strictly operates on existing objects: does NOT create, duplicate, or delete image objects.
"""

import hashlib
import re
from typing import Dict, Any, List, Optional, Set, Tuple
from lxml import etree

from rules.placeholders.classifier import _safe_xpath
from rules.accessibility.base import get_cnvpr_element, safe_xpath
from rules.accessibility.alt_text import GENERIC_ALT_TEXTS
from rules.alt_text.models import ImageCandidate


# Weak/generic alt text regex patterns (e.g., IMG_1234, Image 1, Picture 5)
GENERIC_PATTERNS = [
    re.compile(r"^img[_\-\s]?\d+$", re.IGNORECASE),
    re.compile(r"^image\s*\d*$", re.IGNORECASE),
    re.compile(r"^picture\s*\d*$", re.IGNORECASE),
    re.compile(r"^photo\s*\d*$", re.IGNORECASE),
    re.compile(r"^graphic\s*\d*$", re.IGNORECASE),
    re.compile(r"^figure\s*\d*$", re.IGNORECASE),
    re.compile(r"^untitled\s*\d*$", re.IGNORECASE),
    re.compile(r"^drawing\s*\d*$", re.IGNORECASE),
]


def is_weak_or_generic_alt_text(text: Optional[str]) -> bool:
    """Detects empty or uninformative placeholder alt text."""
    if not text:
        return True
    cleaned = text.strip().lower()
    if not cleaned:
        return True
    if cleaned in GENERIC_ALT_TEXTS:
        return True
    for pat in GENERIC_PATTERNS:
        if pat.match(cleaned):
            return True
    return False


class ImageExtractor:
    """Extracts existing image candidates from slide OOXML trees with surrounding context."""

    def extract_slide_images(
        self,
        raw_slide: Any,
        slide_number: int,
        animation_target_ids: Optional[Set[int]] = None
    ) -> List[ImageCandidate]:
        """Extracts all image candidates from an individual slide."""
        anim_set = animation_target_ids or set()
        slide_elem = raw_slide._element if hasattr(raw_slide, "_element") else raw_slide
        sp_tree = slide_elem.xpath(".//p:spTree")
        if not sp_tree:
            return []

        sp_tree_elem = sp_tree[0]

        # 1. Extract slide text and title for context
        slide_texts: List[Tuple[int, str, Tuple[float, float, float, float]]] = []
        slide_title = ""

        # Collect all texts on slide for spatial & semantic context
        for shape_elem in sp_tree_elem.xpath(".//*[local-name()='sp']"):
            cnv = get_cnvpr_element(shape_elem)
            sid = int(cnv.get("id", "0")) if cnv is not None else 0
            t_nodes = _safe_xpath(shape_elem, ".//a:t")
            t_str = " ".join([t.text for t in t_nodes if t.text]).strip()
            if t_str:
                bbox = self._extract_bbox_pt(shape_elem)
                slide_texts.append((sid, t_str, bbox))
                # Check if title placeholder
                ph = _safe_xpath(shape_elem, ".//p:ph")
                if ph and ph[0].get("type") in ("title", "ctrTitle") and not slide_title:
                    slide_title = t_str

        # If title not found via placeholder, look for top heading
        if not slide_title and slide_texts:
            sorted_by_top = sorted(slide_texts, key=lambda x: x[2][1])
            if sorted_by_top and len(sorted_by_top[0][1]) <= 150:
                slide_title = sorted_by_top[0][1]

        # 2. Extract image candidates recursively
        candidates: List[ImageCandidate] = []
        z_index = [0]

        def _traverse(node, group_id: Optional[int], hierarchy: List[int]):
            for child in node.iterchildren():
                tag = child.tag.split("}")[-1]
                if tag in ("nvGrpSpPr", "grpSpPr"):
                    continue

                if tag == "grpSp":
                    cnv = get_cnvpr_element(child)
                    gid = int(cnv.get("id", "0")) if cnv is not None else 0
                    new_hierarchy = hierarchy + [gid] if gid else hierarchy
                    _traverse(child, gid, new_hierarchy)
                    continue

                # Check if child is an image (<p:pic> or shape with <a:blip>)
                blip_nodes = _safe_xpath(child, ".//a:blip")
                is_pic_tag = (tag == "pic")

                if is_pic_tag or blip_nodes:
                    cand = self._build_candidate(
                        child,
                        slide_number=slide_number,
                        raw_slide=raw_slide,
                        z_order=z_index[0],
                        containing_group_id=group_id,
                        group_hierarchy=hierarchy,
                        anim_set=anim_set,
                        slide_texts=slide_texts,
                        slide_title=slide_title
                    )
                    if cand:
                        candidates.append(cand)
                    z_index[0] += 1

        _traverse(sp_tree_elem, None, [])
        return candidates

    def _build_candidate(
        self,
        elem: Any,
        slide_number: int,
        raw_slide: Any,
        z_order: int,
        containing_group_id: Optional[int],
        group_hierarchy: List[int],
        anim_set: Set[int],
        slide_texts: List[Tuple[int, str, Tuple[float, float, float, float]]],
        slide_title: str
    ) -> Optional[ImageCandidate]:
        """Constructs an ImageCandidate dataclass from an OOXML shape element."""
        cnvpr = get_cnvpr_element(elem)
        if cnvpr is None:
            return None

        shape_id = int(cnvpr.get("id", "0"))
        shape_name = cnvpr.get("name", f"Picture {shape_id}")
        existing_descr = cnvpr.get("descr") or ""
        existing_title = cnvpr.get("title") or ""

        # Check modern Office 365 decorative tag
        is_already_decorative = False
        dec_nodes = safe_xpath(cnvpr, ".//*[local-name()='decorative']")
        if dec_nodes and dec_nodes[0].get("val") in ("1", "true"):
            is_already_decorative = True
        elif existing_descr == "" and cnvpr.get("descr") is not None and not existing_title:
            is_already_decorative = True

        effective_alt = existing_descr or existing_title
        is_generic = is_weak_or_generic_alt_text(effective_alt)

        # Image relationship and format
        blip_nodes = _safe_xpath(elem, ".//a:blip")
        rel_id = blip_nodes[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed") if blip_nodes else None

        image_format = "unknown"
        image_filename = ""
        image_bytes = None
        image_hash = None

        if rel_id and hasattr(raw_slide, "part") and hasattr(raw_slide.part, "rels") and rel_id in raw_slide.part.rels:
            try:
                rel = raw_slide.part.rels[rel_id]
                image_filename = rel.target_ref
                image_format = image_filename.split(".")[-1].lower() if "." in image_filename else "unknown"
                if hasattr(rel, "target_part") and hasattr(rel.target_part, "blob"):
                    image_bytes = rel.target_part.blob
                    image_hash = hashlib.md5(image_bytes).hexdigest()
            except Exception:
                pass

        # Dimensions & position
        bbox_pt = self._extract_bbox_pt(elem)
        dims_pt = (bbox_pt[2], bbox_pt[3])

        # Nearby text and captions
        nearby_texts: List[str] = []
        nearby_caption: Optional[str] = None
        for sid, txt, (sl_l, sl_t, sl_w, sl_h) in slide_texts:
            if sid == shape_id:
                continue
            # Vertical proximity check
            v_dist = sl_t - (bbox_pt[1] + bbox_pt[3])  # Distance below image
            if 0.0 <= v_dist <= 40.0 and abs(sl_l - bbox_pt[0]) <= max(bbox_pt[2], 100.0):
                # Strongly plausible caption immediately below image
                if not nearby_caption:
                    nearby_caption = txt
            elif -40.0 <= (bbox_pt[1] - (sl_t + sl_h)) <= 0.0 and abs(sl_l - bbox_pt[0]) <= max(bbox_pt[2], 100.0):
                # Text immediately above image
                if not nearby_caption and any(w in txt.lower() for w in ("figure", "photo", "image", "diagram", "source")):
                    nearby_caption = txt
            nearby_texts.append(txt)

        # Hyperlinks
        hl_nodes = _safe_xpath(elem, ".//a:hlinkClick")
        has_hyperlink = bool(hl_nodes)
        hl_url = hl_nodes[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") if hl_nodes else None

        is_animated = (shape_id in anim_set)

        return ImageCandidate(
            slide_number=slide_number,
            shape_id=shape_id,
            shape_name=shape_name,
            image_rel_id=rel_id,
            image_format=image_format,
            image_filename=image_filename,
            dimensions_pt=dims_pt,
            bbox_pt=bbox_pt,
            z_order=z_order,
            containing_group_id=containing_group_id,
            group_hierarchy=group_hierarchy,
            nearby_text=nearby_texts,
            nearby_title=slide_title,
            nearby_caption=nearby_caption,
            has_hyperlink=has_hyperlink,
            hyperlink_url=hl_url,
            is_animated=is_animated,
            is_already_decorative=is_already_decorative,
            existing_alt_text=existing_descr or existing_title,
            existing_title=existing_title,
            is_generic_alt_text=is_generic,
            image_hash=image_hash,
            raw_element=elem
        )

    def _extract_bbox_pt(self, elem: Any) -> Tuple[float, float, float, float]:
        """Extracts (left_pt, top_pt, width_pt, height_pt) from shape transforms."""
        xfrm = _safe_xpath(elem, ".//p:spPr/a:xfrm")
        if not xfrm:
            return (0.0, 0.0, 0.0, 0.0)
        off = _safe_xpath(xfrm[0], "a:off")
        ext = _safe_xpath(xfrm[0], "a:ext")
        l_pt = int(off[0].get("x", "0")) / 12700.0 if off else 0.0
        t_pt = int(off[0].get("y", "0")) / 12700.0 if off else 0.0
        w_pt = int(ext[0].get("cx", "0")) / 12700.0 if ext else 0.0
        h_pt = int(ext[0].get("cy", "0")) / 12700.0 if ext else 0.0
        return (l_pt, t_pt, w_pt, h_pt)

    def extract_presentation_images(
        self,
        presentation: Any,
        animation_target_ids_by_slide: Optional[Dict[int, Set[int]]] = None
    ) -> List[ImageCandidate]:
        """Extracts images across all slides in a presentation, resolving cross-slide repetitions."""
        anim_map = animation_target_ids_by_slide or {}
        all_candidates: List[ImageCandidate] = []

        slides = presentation.slides if hasattr(presentation, "slides") else []
        for idx, slide in enumerate(slides):
            s_num = idx + 1
            s_anim = anim_map.get(s_num, set())
            raw_slide = slide.raw_slide if hasattr(slide, "raw_slide") else slide
            cands = self.extract_slide_images(raw_slide, s_num, s_anim)
            all_candidates.extend(cands)

        # Mark cross-slide repetition (same image hash on > 2 slides)
        hash_counts: Dict[str, Set[int]] = {}
        for c in all_candidates:
            if c.image_hash:
                hash_counts.setdefault(c.image_hash, set()).add(c.slide_number)

        for c in all_candidates:
            if c.image_hash and len(hash_counts.get(c.image_hash, set())) >= 3:
                c.is_repeated_across_slides = True

        return all_candidates
