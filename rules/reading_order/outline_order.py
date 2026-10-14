"""
rules/reading_order/outline_order.py

Fixes PowerPoint's Outline View text order so it matches each slide's
existing Reading Order (the order shapes appear in the slide XML / the
order screen readers and the Tab key use).

WHY THIS IS NEEDED
------------------
- "Reading Order" = the order shapes appear in <p:spTree> (document order).
  This is what Tab-cycling and the accessibility "Reading Order" pane use.
- "Outline View" text order = placeholders sorted by their `idx` attribute
  (ascending), NOT by document order.
- When a deck was built by copy/pasting placeholders (common with decks
  converted from Google Slides, or heavily edited files), the `idx` values
  can end up out of sequence, or duplicated, relative to the actual
  document/reading order. That makes Outline View show text in the wrong
  order even though the slide reads correctly top-to-bottom.

WHAT THIS MODULE DOES
---------------------
For each slide:
  1. Walks every <p:ph> element (placeholder marker) in true document
     order (including placeholders nested inside <mc:AlternateContent> blocks).
  2. Checks whether the sequence of `idx` values is already
     non-decreasing (i.e., Outline order already equals Reading order).
  3. If NOT, renumbers just the `idx` attribute values to 1, 2, 3, ...
     following their existing document order. This does NOT move,
     delete, or retype any shape — only the idx number changes.
     Reading Order (document order) is therefore left completely
     untouched, and no visual property (position, size, font,
     formatting, layout) is affected.
  4. Keeps shape names synchronized (e.g. Content Placeholder 1, 2, 3).
"""

from typing import List, Tuple, Dict, Any, Optional
from pathlib import Path
import logging
import re
import shutil
import zipfile
import tempfile
from lxml import etree

from models.presentation import PresentationModel
from models.slide import SlideModel

logger = logging.getLogger(__name__)

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}
PH_TAG = f"{{{NS['p']}}}ph"
PH_PATTERN = re.compile(r'<p:ph([^>]*)/>')


def _safe_xpath(elem, xpath_expr: str) -> list:
    """Executes xpath query safely on either python-pptx BaseOxmlElement or raw lxml Element."""
    try:
        return elem.xpath(xpath_expr)
    except TypeError:
        return elem.xpath(xpath_expr, namespaces=NS)
    except Exception:
        try:
            return elem.xpath(xpath_expr, namespaces=NS)
        except Exception:
            return []


def get_slide_ph_idx_sequence(slide_elem) -> List[int]:
    """
    Returns the list of integer idx values from <p:ph> elements
    in true document order inside the slide's <p:spTree>.
    """
    if slide_elem is None:
        return []
    
    # Locate spTree container
    sp_trees = _safe_xpath(slide_elem, ".//p:spTree")
    sp_tree = sp_trees[0] if sp_trees else slide_elem

    seq: List[int] = []
    for ph in sp_tree.iter(PH_TAG):
        idx_val = ph.get("idx")
        if idx_val is not None:
            try:
                seq.append(int(idx_val))
            except (ValueError, TypeError):
                pass
    return seq


def is_outline_order_monotonic(seq: List[int]) -> bool:
    """Returns True if the sequence of idx values is non-decreasing."""
    if not seq or len(seq) <= 1:
        return True
    return all(seq[i] <= seq[i + 1] for i in range(len(seq) - 1))


def renumber_slide_outline_order(slide_elem) -> int:
    """
    Renumbers the idx attribute of all <p:ph> elements having an idx
    sequentially (1, 2, 3, ...) in true document order.
    Also synchronizes any 'Content Placeholder {old}' or 'Subtitle {old}'
    names on <p:cNvPr> to match the new idx.
    Leaves title without idx untouched.
    Returns the count of placeholders renumbered.
    """
    if slide_elem is None:
        return 0

    sp_trees = _safe_xpath(slide_elem, ".//p:spTree")
    sp_tree = sp_trees[0] if sp_trees else slide_elem

    counter = 0
    # Walk all shape elements in spTree in document order
    for child in sp_tree.iter():
        if child.tag == PH_TAG:
            idx_val = child.get("idx")
            if idx_val is not None:
                counter += 1
                new_idx_str = str(counter)
                child.set("idx", new_idx_str)

                # Synchronize companion cNvPr name if shape container has one
                # Look upwards to find shape container: sp, pic, graphicFrame
                parent_shape = child
                while parent_shape is not None and parent_shape != sp_tree:
                    tag_name = parent_shape.tag.split("}")[-1]
                    if tag_name in ("sp", "pic", "graphicFrame", "cxnSp"):
                        break
                    parent_shape = parent_shape.getparent()

                if parent_shape is not None and parent_shape != sp_tree:
                    cnvprs = _safe_xpath(parent_shape, ".//p:cNvPr")
                    if cnvprs:
                        c_name = cnvprs[0].get("name", "")
                        # If name matches Content Placeholder \d+
                        if re.match(r"^Content Placeholder \d+$", c_name):
                            cnvprs[0].set("name", f"Content Placeholder {new_idx_str}")
                        elif re.match(r"^Subtitle \d+$", c_name):
                            cnvprs[0].set("name", f"Subtitle {new_idx_str}")

    return counter


def fix_outline_order_on_slide(slide_model: SlideModel) -> Tuple[bool, int, List[int], List[int]]:
    """
    Checks and fixes outline order on a single SlideModel in memory.
    Returns: (was_fixed: bool, count_renumbered: int, old_seq: List[int], new_seq: List[int])
    """
    raw_s = slide_model.raw_slide
    if raw_s is None or not hasattr(raw_s, "_element"):
        return False, 0, [], []

    elem = raw_s._element
    seq = get_slide_ph_idx_sequence(elem)
    if seq and not is_outline_order_monotonic(seq):
        count = renumber_slide_outline_order(elem)
        new_seq = get_slide_ph_idx_sequence(elem)
        logger.info(
            "Slide %d: Aligned Outline Order to match Reading Order (%s -> %s, %d placeholders).",
            slide_model.slide_number, seq, new_seq, count
        )
        return True, count, seq, new_seq

    return False, 0, seq, seq


def fix_outline_order_on_presentation(presentation: PresentationModel) -> Dict[str, Any]:
    """
    Evaluates and fixes Outline Order across all slides in a PresentationModel in-place.
    Returns a detailed summary report dictionary.
    """
    fixed_slides: List[Dict[str, Any]] = []
    total_renumbered = 0
    total_slides = len(presentation.slides)

    for slide in presentation.slides:
        was_fixed, count, old_seq, new_seq = fix_outline_order_on_slide(slide)
        if was_fixed:
            total_renumbered += count
            fixed_slides.append({
                "slide_number": slide.slide_number,
                "renumbered_count": count,
                "old_idx_sequence": old_seq,
                "new_idx_sequence": new_seq,
            })

    return {
        "total_slides": total_slides,
        "slides_evaluated": total_slides,
        "slides_checked": total_slides,
        "slides_fixed": len(fixed_slides),
        "total_placeholders_renumbered": total_renumbered,
        "details": fixed_slides,
    }


def fix_outline_order_pptx(input_path: Path, output_path: Optional[Path] = None) -> Dict[str, Any]:
    """
    Standalone / CLI compatible implementation that extracts, inspects, and renumbers
    a PPTX file directly via ZIP manipulation.
    """
    inp = Path(input_path).resolve()
    if not inp.exists():
        raise FileNotFoundError(f"Input presentation not found: {inp}")

    out = (
        Path(output_path).resolve()
        if output_path is not None
        else inp.with_name(f"{inp.stem}_outline_fixed.pptx")
    )

    fixed: List[Tuple[str, int]] = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(inp) as z:
            z.extractall(tmp)

        slides_dir = tmp / "ppt" / "slides"
        slide_files = sorted(
            slides_dir.glob("slide*.xml"),
            key=lambda p: int(re.search(r'\d+', p.stem).group()) if re.search(r'\d+', p.stem) else 0,
        )

        for slide_path in slide_files:
            tree = etree.parse(str(slide_path))
            root = tree.getroot()
            spTree = root.find('.//p:cSld/p:spTree', NS)
            if spTree is None:
                continue

            seq = []
            for ph in spTree.iter(PH_TAG):
                idx = ph.get('idx')
                if idx is not None:
                    try:
                        seq.append(int(idx))
                    except (ValueError, TypeError):
                        pass

            if seq and not is_outline_order_monotonic(seq):
                raw = slide_path.read_bytes()
                data = raw.decode('utf-8')
                counter = {'n': 0}

                def repl(m):
                    attrs = m.group(1)
                    if 'idx=' not in attrs:
                        return m.group(0)
                    counter['n'] += 1
                    new_attrs = re.sub(r'idx="\d+"', f'idx="{counter["n"]}"', attrs)
                    return f'<p:ph{new_attrs}/>'

                new_data = PH_PATTERN.sub(repl, data)
                slide_path.write_bytes(new_data.encode('utf-8'))
                fixed.append((slide_path.name, counter['n']))

        if not fixed:
            shutil.copy(inp, out)
        else:
            if out.exists():
                out.unlink()
            with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
                for file in tmp.rglob("*"):
                    if file.is_file():
                        zf.write(file, file.relative_to(tmp))

    return {
        "input_path": str(inp),
        "output_path": str(out),
        "slides_fixed": len(fixed),
        "details": fixed,
    }
