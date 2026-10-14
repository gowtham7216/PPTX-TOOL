#!/usr/bin/env python3
"""
ConvertPlaceholder
======================
GUI wrapper around the pptx placeholder-classification/conversion logic.
Works on a SINGLE .pptx file or an entire FOLDER (with subfolders) of them.

Build to a Windows .exe with (run in VS Code / terminal on Windows, with
pyinstaller and the deps below installed):

    pip install lxml python-pptx pyinstaller
    pyinstaller --onefile --windowed --name ConvertPlaceholder ConvertPlaceholder.py

The resulting exe will be in the "dist" folder.
"""
import os
import sys
import glob
import re
import shutil
import zipfile
import tempfile
import traceback
import threading
import queue

from lxml import etree

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --------------------------------------------------------------------------
# Namespaces / constants
# --------------------------------------------------------------------------
P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
ADEC_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative'
DECORATIVE_EXT_URI = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}'
PLAIN_TEXTBOX_DEFAULT_SZ = '1800'

NS = {'p': P_NS, 'a': A_NS, 'mc': MC_NS}


def qn(prefix, tag):
    return f'{{{NS[prefix]}}}{tag}'


NV_TAG_MAP = {'sp': 'nvSpPr', 'pic': 'nvPicPr', 'graphicFrame': 'nvGraphicFramePr', 'cxnSp': 'nvCxnSpPr'}


def get_nv_container(el):
    return el.find(qn('p', NV_TAG_MAP[etree.QName(el).localname]))


def has_ph(el):
    nv = get_nv_container(el)
    nvpr = nv.find(qn('p', 'nvPr'))
    if nvpr is None:
        return False
    return nvpr.find(qn('p', 'ph')) is not None


def has_existing_decorative_mark(el):
    """True only if this shape was ALREADY marked Decorative (e.g. via
    PowerPoint's own 'Mark as Decorative' button) before we touched it.
    These are the only shapes we leave alone / skip."""
    nv = get_nv_container(el)
    cnvpr = nv.find(qn('p', 'cNvPr'))
    extlst = cnvpr.find(qn('a', 'extLst'))
    if extlst is None:
        return False
    for ext in extlst.findall(qn('a', 'ext')):
        if ext.get('uri') == DECORATIVE_EXT_URI:
            dec = ext.find(f'{{{ADEC_NS}}}decorative')
            if dec is not None and dec.get('val') == '1':
                return True
    return False


def shape_text(el):
    texts = [t.text for t in el.findall(f'.//{qn("a","t")}') if t.text]
    return ' '.join(texts).strip()


def is_table(el):
    gd = el.find(f'.//{qn("a","graphicData")}')
    return gd is not None and 'table' in gd.get('uri', '')


# --------------------------------------------------------------------------
# Walk a slide's shape tree
# --------------------------------------------------------------------------
def classify_slide_shapes(sptree, skip_empty_text=False):
    classification = {}
    flagged = []
    pics_no_alt = []
    seen_ids = set()

    def walk(container):
        for child in list(container):
            tag = etree.QName(child).localname
            if tag in ('sp', 'pic', 'graphicFrame', 'cxnSp'):
                cnvpr = get_nv_container(child).find(qn('p', 'cNvPr'))
                _id = cnvpr.get('id')
                if _id in seen_ids:
                    continue
                seen_ids.add(_id)
                if has_ph(child):
                    continue  # already a real placeholder - nothing to do
                if has_existing_decorative_mark(child):
                    continue  # already manually marked Decorative - leave as is

                if tag == 'graphicFrame':
                    cls = 'tbl' if is_table(child) else None
                    if cls is None:
                        flagged.append((_id, 'graphicFrame(non-table)', cnvpr.get('name', '')))
                elif tag == 'pic':
                    cls = 'pic'
                    descr = cnvpr.get('descr', '')
                    if not descr.strip():
                        pics_no_alt.append((_id, cnvpr.get('name', '')))
                elif tag in ('sp', 'cxnSp'):
                    if skip_empty_text and not shape_text(child):
                        cls = None  # user chose to leave empty text boxes untouched
                    else:
                        cls = 'body'  # convert every non-decorative-marked shape,
                                      # even ones with no text in them
                else:
                    cls = None

                if cls is not None:
                    classification[int(_id)] = cls
            elif tag == 'AlternateContent':
                choice = child.find(qn('mc', 'Choice'))
                fallback = child.find(qn('mc', 'Fallback'))
                if choice is not None:
                    walk(choice)
                if fallback is not None:
                    walk(fallback)
            elif tag == 'grpSp':
                walk(child)

    walk(sptree)
    return classification, flagged, pics_no_alt


# --------------------------------------------------------------------------
# XML shape-editing helpers
# --------------------------------------------------------------------------
def mark_decorative(el):
    nv = get_nv_container(el)
    cnvpr = nv.find(qn('p', 'cNvPr'))
    extlst = cnvpr.find(qn('a', 'extLst'))
    if extlst is None:
        extlst = etree.SubElement(cnvpr, qn('a', 'extLst'))
    for ext in extlst.findall(qn('a', 'ext')):
        if ext.get('uri') == DECORATIVE_EXT_URI:
            return
    ext = etree.SubElement(extlst, qn('a', 'ext'))
    ext.set('uri', DECORATIVE_EXT_URI)
    dec = etree.SubElement(ext, f'{{{ADEC_NS}}}decorative')
    dec.set('val', '1')


CONTENT_PLACEHOLDER_LABEL = 'Content Placeholder'
# ph types that our tool treats as "content" and is allowed to rename/retype.
# We deliberately never touch title/subTitle/ftr/dt/sldNum etc.
CONTENT_PH_TYPES = (None, 'body', 'pic', 'tbl', 'obj')


def add_placeholder(el, ph_type, idx):
    nv = get_nv_container(el)
    nvpr = nv.find(qn('p', 'nvPr'))
    if nvpr is None:
        nvpr = etree.SubElement(nv, qn('p', 'nvPr'))
    if nvpr.find(qn('p', 'ph')) is not None:
        return
    ph = etree.Element(qn('p', 'ph'))
    # Keep the functional placeholder TYPE (body/pic/tbl) - PowerPoint's
    # Outline view and some accessibility tooling only recognize text in
    # type="body" (or "title") placeholders, and pictures/tables need their
    # own type too. The shape's visible NAME is still unified below.
    if ph_type in ('body', 'pic', 'tbl'):
        ph.set('type', ph_type)
    ph.set('idx', str(idx))
    nvpr.insert(0, ph)

    # Rename the shape itself so PowerPoint's Reading Order / Selection
    # panes show "Content Placeholder N" instead of the old "Text Box N" /
    # "TextBox N" / "Picture N" name - regardless of its functional type.
    cnvpr = nv.find(qn('p', 'cNvPr'))
    if cnvpr is not None:
        cnvpr.set('name', f'{CONTENT_PLACEHOLDER_LABEL} {idx}')


PPR_CHILD_ORDER = [
    'lnSpc', 'spcBef', 'spcAft',
    'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont',
    'buNone', 'buAutoNum', 'buChar', 'buBlip',
    'tabLst', 'defRPr', 'extLst',
]
RPR_CHILD_ORDER = [
    'ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill',
    'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill',
    'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst',
]


def _insert_in_schema_order(parent, new_el, tag, order_list):
    my_rank = order_list.index(tag)
    insert_at = len(list(parent))
    for i, child in enumerate(parent):
        local = etree.QName(child).localname
        if local in order_list and order_list.index(local) > my_rank:
            insert_at = i
            break
    parent.insert(insert_at, new_el)


def _ensure_para_child(ppr, tag, build_fn):
    if ppr.find(qn('a', tag)) is not None:
        return
    _insert_in_schema_order(ppr, build_fn(), tag, PPR_CHILD_ORDER)


def freeze_paragraph_properties(sp_el):
    for p in sp_el.findall(f'.//{qn("a","p")}'):
        ppr = p.find(qn('a', 'pPr'))
        if ppr is None:
            ppr = etree.Element(qn('a', 'pPr'))
            p.insert(0, ppr)

        # A paragraph that already carries a REAL bullet/number marker is a
        # deliberate list item - leave its formatting completely alone so we
        # never flatten an intentional list's own spacing/indent.
        has_real_bullet = (ppr.find(qn('a', 'buChar')) is not None
                            or ppr.find(qn('a', 'buAutoNum')) is not None)
        if has_real_bullet:
            continue

        if ppr.get('algn') is None:
            ppr.set('algn', 'l')

        def _lnspc():
            el = etree.Element(qn('a', 'lnSpc'))
            etree.SubElement(el, qn('a', 'spcPct')).set('val', '100000')
            return el

        def _spcbef():
            el = etree.Element(qn('a', 'spcBef'))
            etree.SubElement(el, qn('a', 'spcPts')).set('val', '0')
            return el

        def _spcaft():
            el = etree.Element(qn('a', 'spcAft'))
            etree.SubElement(el, qn('a', 'spcPts')).set('val', '0')
            return el

        def _bunone():
            return etree.Element(qn('a', 'buNone'))

        _ensure_para_child(ppr, 'lnSpc', _lnspc)
        _ensure_para_child(ppr, 'spcBef', _spcbef)
        _ensure_para_child(ppr, 'spcAft', _spcaft)
        # Explicitly say "no bullet" so this paragraph can never inherit an
        # unwanted bullet from the slide layout's placeholder-level list
        # style once this shape becomes a real placeholder.
        _ensure_para_child(ppr, 'buNone', _bunone)

        sp_lststyle = sp_el.find(f'.//{qn("a","lstStyle")}')
        if sp_lststyle is not None and len(sp_lststyle) == 0:
            lvl1 = etree.SubElement(sp_lststyle, qn('a', 'lvl1pPr'))
            lvl1.set('algn', 'l')
            etree.SubElement(etree.SubElement(lvl1, qn('a', 'lnSpc')), qn('a', 'spcPct')).set('val', '100000')
            etree.SubElement(etree.SubElement(lvl1, qn('a', 'spcBef')), qn('a', 'spcPts')).set('val', '0')


def backfill_run_sizes(sp_el):
    for p in sp_el.findall(f'.//{qn("a","p")}'):
        ppr = p.find(qn('a', 'pPr'))
        para_defrpr = ppr.find(qn('a', 'defRPr')) if ppr is not None else None
        para_sz = para_defrpr.get('sz') if para_defrpr is not None else None
        fallback_sz = para_sz if para_sz else PLAIN_TEXTBOX_DEFAULT_SZ

        for r in p.findall(qn('a', 'r')):
            rpr = r.find(qn('a', 'rPr'))
            if rpr is None:
                rpr = etree.Element(qn('a', 'rPr'))
                rpr.set('lang', 'en-US')
                r.insert(0, rpr)
            if rpr.get('sz') is None:
                rpr.set('sz', fallback_sz)

        end_rpr = p.find(qn('a', 'endParaRPr'))
        if end_rpr is not None and end_rpr.get('sz') is None:
            end_rpr.set('sz', fallback_sz)

        for rpr in p.findall(f'.//{qn("a","rPr")}'):
            if rpr.get('sz') is None:
                rpr.set('sz', fallback_sz)


def freeze_autofit(el):
    if etree.QName(el).localname != 'sp':
        return
    bodypr = el.find(f'.//{qn("a","bodyPr")}')
    if bodypr is None:
        return
    autofit = bodypr.find(qn('a', 'spAutoFit'))
    if autofit is not None:
        idx = list(bodypr).index(autofit)
        bodypr.remove(autofit)
        bodypr.insert(idx, etree.Element(qn('a', 'noAutofit')))


TIMES_NEW_ROMAN = 'Times New Roman'
CALIBRI = 'Calibri'
TITLE_SZ = '4000'
BODY_SZ_MIN = 1800
BODY_SZ_MAX = 2600

_FONT_CACHE = {}


def get_font_object(font_name: str, size_pt: float):
    cache_key = (font_name, int(round(size_pt)))
    if cache_key in _FONT_CACHE:
        return _FONT_CACHE[cache_key]

    font = None
    try:
        from PIL import ImageFont
        font_file = 'times.ttf' if 'times' in font_name.lower() else 'calibri.ttf'
        try:
            font = ImageFont.truetype(font_file, int(round(size_pt)))
        except Exception:
            windir = os.environ.get('WINDIR', r'C:\Windows')
            full_path = os.path.join(windir, 'Fonts', font_file)
            if os.path.exists(full_path):
                font = ImageFont.truetype(full_path, int(round(size_pt)))
    except Exception:
        font = None

    _FONT_CACHE[cache_key] = font
    return font


def check_text_overflow(sp_el, target_font: str = 'Times New Roman') -> bool:
    xfrm = sp_el.find(f'.//{qn("a", "xfrm")}')
    if xfrm is None:
        return False
    ext = xfrm.find(qn('a', 'ext'))
    if ext is None:
        return False

    cx = int(ext.get('cx', '0'))
    cy = int(ext.get('cy', '0'))
    if cx <= 0 or cy <= 0:
        return False

    w_pt = cx / 12700.0
    h_pt = cy / 12700.0

    bodypr = sp_el.find(f'.//{qn("a", "bodyPr")}')
    l_ins = int(bodypr.get('lIns', '91440')) / 12700.0 if bodypr is not None and bodypr.get('lIns') is not None else 7.2
    r_ins = int(bodypr.get('rIns', '91440')) / 12700.0 if bodypr is not None and bodypr.get('rIns') is not None else 7.2
    t_ins = int(bodypr.get('tIns', '45720')) / 12700.0 if bodypr is not None and bodypr.get('tIns') is not None else 3.6
    b_ins = int(bodypr.get('bIns', '45720')) / 12700.0 if bodypr is not None and bodypr.get('bIns') is not None else 3.6
    wrap_mode = bodypr.get('wrap', 'square') if bodypr is not None else 'square'

    avail_w = max(1.0, w_pt - (l_ins + r_ins))
    avail_h = max(1.0, h_pt - (t_ins + b_ins))

    total_h = 0.0

    paragraphs = sp_el.findall(f'.//{qn("a", "p")}')
    for p in paragraphs:
        p_text = ''.join(t.text for t in p.findall(f'.//{qn("a", "t")}') if t.text).strip()
        if not p_text:
            continue

        sz_val = None
        for r in p.findall(qn('a', 'r')):
            rpr = r.find(qn('a', 'rPr'))
            if rpr is not None and rpr.get('sz'):
                sz_val = rpr.get('sz')
                break
        if not sz_val:
            defrpr = p.find(f'.//{qn("a", "defRPr")}')
            if defrpr is not None and defrpr.get('sz'):
                sz_val = defrpr.get('sz')
        sz_pt = int(sz_val) / 100.0 if sz_val else 18.0

        ppr = p.find(qn('a', 'pPr'))
        mar_l = int(ppr.get('marL', '0')) / 12700.0 if ppr is not None and ppr.get('marL') is not None else 0.0
        mar_r = int(ppr.get('marR', '0')) / 12700.0 if ppr is not None and ppr.get('marR') is not None else 0.0
        p_avail_w = max(1.0, avail_w - (mar_l + mar_r))

        font_obj = get_font_object(target_font, sz_pt)

        if font_obj is not None:
            ascent, descent = font_obj.getmetrics()
            line_h = (ascent + descent) * 1.15
            if wrap_mode == 'none':
                bbox = font_obj.getbbox(p_text)
                if (bbox[2] - bbox[0]) > p_avail_w:
                    return True
                num_lines = 1
            else:
                words = p_text.split()
                lines = []
                curr_line = ''
                for w in words:
                    cand = f'{curr_line} {w}'.strip()
                    bbox = font_obj.getbbox(cand)
                    cand_w = bbox[2] - bbox[0]
                    if cand_w <= p_avail_w:
                        curr_line = cand
                    else:
                        if curr_line:
                            lines.append(curr_line)
                        w_bbox = font_obj.getbbox(w)
                        if (w_bbox[2] - w_bbox[0]) > p_avail_w:
                            return True
                        curr_line = w
                if curr_line:
                    lines.append(curr_line)
                num_lines = max(1, len(lines))
        else:
            char_w = sz_pt * (0.50 if 'times' in target_font.lower() else 0.46)
            line_h = sz_pt * (1.25 if 'times' in target_font.lower() else 1.15)
            if wrap_mode == 'none':
                if (len(p_text) * char_w) > p_avail_w:
                    return True
                num_lines = 1
            else:
                chars_per_line = max(1, int(p_avail_w / char_w))
                for w in p_text.split():
                    if len(w) > chars_per_line:
                        return True
                num_lines = max(1, (len(p_text) + chars_per_line - 1) // chars_per_line)

        total_h += (num_lines * line_h)

    if total_h > (avail_h * 1.02):
        return True

    return False


def resolve_shape_font(sp_el, primary_font: str = TIMES_NEW_ROMAN, fallback_font: str = CALIBRI) -> str:
    """
    Standardizes font to primary_font (Times New Roman) by default.
    If applying primary_font causes contents to overflow the placeholder box, falls back to fallback_font (Calibri).
    """
    if check_text_overflow(sp_el, target_font=primary_font):
        return fallback_font
    return primary_font


def set_latin_font(rpr, typeface):
    latin = rpr.find(qn('a', 'latin'))
    if latin is None:
        latin = etree.Element(qn('a', 'latin'))
        _insert_in_schema_order(rpr, latin, 'latin', RPR_CHILD_ORDER)
    latin.set('typeface', typeface)


def apply_size(rpr, is_title):
    if is_title:
        rpr.set('sz', TITLE_SZ)
        return
    sz = rpr.get('sz')
    if sz is not None:
        val = int(sz)
        if val > BODY_SZ_MAX:
            rpr.set('sz', str(BODY_SZ_MAX))
        elif val < BODY_SZ_MIN:
            rpr.set('sz', str(BODY_SZ_MIN))


def reset_autofit_scale(sp_el):
    bodypr = sp_el.find(f'.//{qn("a","bodyPr")}')
    if bodypr is None:
        return
    autofit = bodypr.find(qn('a', 'normAutofit'))
    if autofit is not None and autofit.get('fontScale') is not None:
        del autofit.attrib['fontScale']
        if autofit.get('lnSpcReduction') is not None:
            del autofit.attrib['lnSpcReduction']


def apply_font_and_size_rules(sp_el, is_title):
    chosen_font = resolve_shape_font(sp_el, primary_font=TIMES_NEW_ROMAN, fallback_font=CALIBRI)
    for p in sp_el.findall(f'.//{qn("a","p")}'):
        for r in p.findall(qn('a', 'r')):
            rpr = r.find(qn('a', 'rPr'))
            if rpr is None:
                rpr = etree.Element(qn('a', 'rPr'))
                rpr.set('lang', 'en-US')
                r.insert(0, rpr)
            set_latin_font(rpr, chosen_font)

        end_rpr = p.find(qn('a', 'endParaRPr'))
        if end_rpr is not None:
            set_latin_font(end_rpr, chosen_font)


# --------------------------------------------------------------------------
# Per-slide processing: classify + apply
# --------------------------------------------------------------------------
def process_slide(path, slide_num, skip_empty_text=False):
    tree = etree.parse(path)
    root = tree.getroot()
    sptree = root.find(f'.//{qn("p","cSld")}/{qn("p","spTree")}')

    classification, flagged, pics_no_alt = classify_slide_shapes(sptree, skip_empty_text=skip_empty_text)

    groups = {}
    order = []

    def walk(container):
        for child in list(container):
            tag = etree.QName(child).localname
            if tag in ('sp', 'pic', 'graphicFrame', 'cxnSp'):
                cnvpr = get_nv_container(child).find(qn('p', 'cNvPr'))
                _id = cnvpr.get('id')
                groups.setdefault(_id, []).append(child)
                if _id not in order:
                    order.append(_id)
            elif tag == 'AlternateContent':
                choice = child.find(qn('mc', 'Choice'))
                fallback = child.find(qn('mc', 'Fallback'))
                if choice is not None:
                    walk(choice)
                if fallback is not None:
                    walk(fallback)
            elif tag == 'grpSp':
                walk(child)

    walk(sptree)

    next_idx = 1
    log = []

    for _id in order:
        elems = groups[_id]
        first = elems[0]
        if has_ph(first):
            continue

        cls = classification.get(int(_id))
        if cls is None:
            log.append(f"  ! shape id={_id} not auto-classified (unrecognized kind) - left untouched")
            continue

        if cls == 'decorative':
            for el in elems:
                mark_decorative(el)
            log.append(f"  decorative: id={_id}")
        elif cls in ('body', 'pic', 'tbl'):
            idx = next_idx
            next_idx += 1
            primary = elems[0]
            if cls == 'body' and etree.QName(primary).localname == 'sp':
                freeze_paragraph_properties(primary)
                apply_font_and_size_rules(primary, is_title=False)
            freeze_autofit(primary)
            add_placeholder(primary, cls, idx)
            log.append(f"  {cls} placeholder idx={idx}: id={_id}")

    for _id in order:
        first = groups[_id][0]
        if etree.QName(first).localname != 'sp':
            continue
        nv = get_nv_container(first)
        nvpr = nv.find(qn('p', 'nvPr')) if nv is not None else None
        ph = nvpr.find(qn('p', 'ph')) if nvpr is not None else None
        if ph is None:
            continue
        is_title = ph.get('type') in ('title', 'ctrTitle')
        apply_font_and_size_rules(first, is_title=is_title)

    # Name-normalization pass: ANY shape that is (now, or already was) a
    # content-type placeholder gets checked - if its name doesn't already
    # say "Placeholder", fix the name so the Reading Order / Selection panes
    # are accurate. Runs on pre-existing placeholders too, not just ones we
    # just created this run. Title/subtitle/footer/date/slide-number
    # placeholders are left untouched.
    for _id in order:
        el = groups[_id][0]
        if not has_ph(el):
            continue
        nv = get_nv_container(el)
        nvpr = nv.find(qn('p', 'nvPr'))
        ph = nvpr.find(qn('p', 'ph'))
        ph_type = ph.get('type')
        if ph_type not in CONTENT_PH_TYPES:
            continue  # title/subtitle/etc already follow PowerPoint's own convention
        cnvpr = nv.find(qn('p', 'cNvPr'))
        current_name = cnvpr.get('name', '') if cnvpr is not None else ''
        if 'Placeholder' in current_name:
            continue
        idx = ph.get('idx', _id)
        cnvpr.set('name', f'{CONTENT_PLACEHOLDER_LABEL} {idx}')
        log.append(f"  renamed id={_id} -> '{CONTENT_PLACEHOLDER_LABEL} {idx}'")

    tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return log, pics_no_alt


REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def get_ordered_slide_paths(tmpdir):
    """Return slide XML paths in the ORDER THEY APPEAR IN THE PRESENTATION
    (i.e. matching the slide numbers you see in PowerPoint's slide panel),
    not just sorted by internal filename - those can differ once slides
    have been reordered, duplicated, or deleted."""
    pres_path = os.path.join(tmpdir, 'ppt', 'presentation.xml')
    rels_path = os.path.join(tmpdir, 'ppt', '_rels', 'presentation.xml.rels')

    def _fallback():
        files = glob.glob(os.path.join(tmpdir, 'ppt', 'slides', 'slide*.xml'))
        files.sort(key=lambda p: int(re.search(r'slide(\d+)\.xml', p).group(1)))
        return files

    if not os.path.exists(pres_path) or not os.path.exists(rels_path):
        return _fallback()

    rid_to_target = {}
    for rel in etree.parse(rels_path).getroot():
        rid_to_target[rel.get('Id')] = rel.get('Target')

    root = etree.parse(pres_path).getroot()
    sldidlst = root.find(qn('p', 'sldIdLst'))
    ordered = []
    if sldidlst is not None:
        for sldid in sldidlst:
            rid = sldid.get(f'{{{REL_NS}}}id')
            target = rid_to_target.get(rid)
            if target:
                ordered.append(os.path.normpath(os.path.join(tmpdir, 'ppt', target)))
    return ordered if ordered else _fallback()


def parse_slide_spec(spec):
    """Parse '1', '2-4', '2,7,8', '1, 3-5, 9' etc into a sorted set of ints.
    Returns None for an empty/blank spec (meaning: all slides)."""
    spec = (spec or '').strip()
    if not spec:
        return None
    result = set()
    for part in spec.split(','):
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            a, b = part.split('-', 1)
            a, b = int(a.strip()), int(b.strip())
            if a > b:
                a, b = b, a
            result.update(range(a, b + 1))
        else:
            result.add(int(part))
    return result


def convert_file(in_path, out_path, log_fn=None, slide_numbers=None, skip_empty_text=False):
    tmpdir = tempfile.mkdtemp(prefix='pptx_convert_')
    full_log = []
    all_no_alt = []
    try:
        with zipfile.ZipFile(in_path) as z:
            z.extractall(tmpdir)

        slide_paths = get_ordered_slide_paths(tmpdir)

        for display_num, f in enumerate(slide_paths, start=1):
            if slide_numbers is not None and display_num not in slide_numbers:
                continue
            header = f"=== Slide {display_num} ({os.path.basename(f)}) ==="
            lines, no_alt = process_slide(f, display_num, skip_empty_text=skip_empty_text)
            full_log.append(header)
            full_log.extend(lines)
            all_no_alt.extend((display_num, sid, name) for sid, name in no_alt)
            if log_fn:
                log_fn(header)
                for line in lines:
                    log_fn(line)

        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dirpath, _, filenames in os.walk(tmpdir):
                for fn in filenames:
                    full = os.path.join(dirpath, fn)
                    arcname = os.path.relpath(full, tmpdir)
                    zf.write(full, arcname)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return full_log, all_no_alt


IGNORE_PREFIXES = ('~$',)


def find_pptx_files(input_root):
    for dirpath, _dirnames, filenames in os.walk(input_root):
        for fn in filenames:
            if fn.lower().endswith('.pptx') and not fn.startswith(IGNORE_PREFIXES):
                yield os.path.join(dirpath, fn)


# --------------------------------------------------------------------------
# Batch / single-file runner used by the GUI (mirrors the old main())
# --------------------------------------------------------------------------
def run_batch(input_root, output_root, suffix, quiet, log_fn, slide_numbers=None, skip_empty_text=False):
    input_root = os.path.abspath(input_root)
    output_root = os.path.abspath(output_root)

    if not os.path.isdir(input_root):
        log_fn(f"ERROR: input folder not found: {input_root}")
        return False

    pptx_files = sorted(find_pptx_files(input_root))
    if not pptx_files:
        log_fn(f"No .pptx files found under {input_root}")
        return True

    log_fn(f"Found {len(pptx_files)} .pptx file(s) under {input_root}")
    log_fn(f"Output will be written under {output_root}\n")

    succeeded, failed = [], []
    total_no_alt = 0

    for in_path in pptx_files:
        rel_path = os.path.relpath(in_path, input_root)
        if suffix:
            rel_dir, rel_name = os.path.split(rel_path)
            base, ext = os.path.splitext(rel_name)
            rel_path = os.path.join(rel_dir, f"{base}{suffix}{ext}")
        out_path = os.path.join(output_root, rel_path)

        log_fn(f"--- {rel_path} ---")
        try:
            log, no_alt = convert_file(in_path, out_path, log_fn=None if quiet else log_fn,
                                        slide_numbers=slide_numbers, skip_empty_text=skip_empty_text)
            unclassified = [l for l in log if 'not auto-classified' in l]
            succeeded.append(rel_path)
            total_no_alt += len(no_alt)
            log_fn(f"    -> OK: {out_path}"
                    f"{f' | {len(unclassified)} shape(s) not auto-classified' if unclassified else ''}"
                    f"{f' | {len(no_alt)} picture(s) missing alt text' if no_alt else ''}\n")
        except Exception as e:
            failed.append((rel_path, str(e)))
            log_fn(f"    -> FAILED: {e}")
            if not quiet:
                log_fn(traceback.format_exc())
            log_fn("")

    log_fn("=" * 60)
    log_fn(f"Done. {len(succeeded)} succeeded, {len(failed)} failed. "
           f"{total_no_alt} picture(s) total missing alt text (still converted, just flagged).")
    if failed:
        log_fn("\nFailed files:")
        for rel_path, err in failed:
            log_fn(f"  {rel_path}: {err}")
    return len(failed) == 0


def run_single(in_path, out_path, quiet, log_fn, slide_numbers=None, skip_empty_text=False):
    log_fn(f"--- {os.path.basename(in_path)} ---")
    try:
        log, no_alt = convert_file(in_path, out_path, log_fn=None if quiet else log_fn,
                                    slide_numbers=slide_numbers, skip_empty_text=skip_empty_text)
        unclassified = [l for l in log if 'not auto-classified' in l]
        log_fn(f"    -> OK: {out_path}"
                f"{f' | {len(unclassified)} shape(s) not auto-classified' if unclassified else ''}"
                f"{f' | {len(no_alt)} picture(s) missing alt text' if no_alt else ''}\n")
        log_fn("Done. 1 succeeded, 0 failed.")
        return True
    except Exception as e:
        log_fn(f"    -> FAILED: {e}")
        log_fn(traceback.format_exc())
        log_fn("Done. 0 succeeded, 1 failed.")
        return False


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convert Placeholder")
        self.geometry("760x560")
        self.minsize(640, 480)

        self.mode = tk.StringVar(value="folder")
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.suffix = tk.StringVar(value="")
        self.quiet = tk.BooleanVar(value=True)
        self.slides_spec = tk.StringVar(value="")
        self.skip_empty_text = tk.BooleanVar(value=False)

        self.msg_queue = queue.Queue()
        self.worker_thread = None

        self._build_ui()
        self.after(100, self._poll_queue)

    def _build_ui(self):
        pad = {'padx': 10, 'pady': 6}

        mode_frame = ttk.LabelFrame(self, text="What do you want to convert?")
        mode_frame.pack(fill='x', **pad)
        ttk.Radiobutton(mode_frame, text="A single .pptx file",
                         variable=self.mode, value="file",
                         command=self._on_mode_change).pack(side='left', padx=10, pady=8)
        ttk.Radiobutton(mode_frame, text="A whole folder (incl. subfolders)",
                         variable=self.mode, value="folder",
                         command=self._on_mode_change).pack(side='left', padx=10, pady=8)

        in_frame = ttk.LabelFrame(self, text="Input")
        in_frame.pack(fill='x', **pad)
        ttk.Entry(in_frame, textvariable=self.input_path).pack(side='left', fill='x', expand=True, padx=(10, 6), pady=8)
        self.browse_in_btn = ttk.Button(in_frame, text="Browse...", command=self._browse_input)
        self.browse_in_btn.pack(side='left', padx=(0, 10), pady=8)

        out_frame = ttk.LabelFrame(self, text="Output")
        out_frame.pack(fill='x', **pad)
        ttk.Entry(out_frame, textvariable=self.output_path).pack(side='left', fill='x', expand=True, padx=(10, 6), pady=8)
        self.browse_out_btn = ttk.Button(out_frame, text="Browse...", command=self._browse_output)
        self.browse_out_btn.pack(side='left', padx=(0, 10), pady=8)

        opt_frame = ttk.Frame(self)
        opt_frame.pack(fill='x', **pad)
        ttk.Label(opt_frame, text="Optional filename suffix:").pack(side='left')
        ttk.Entry(opt_frame, textvariable=self.suffix, width=20).pack(side='left', padx=(6, 20))
        ttk.Checkbutton(opt_frame, text="Quiet log (summary only)", variable=self.quiet).pack(side='left')

        # Slide-range selector - scopes the WHOLE conversion (placeholder
        # conversion, renaming, everything) to only the listed slides.
        self.slides_frame = ttk.Frame(self)
        self.slides_frame.pack(fill='x', **pad)
        ttk.Label(self.slides_frame,
                  text="Convert only these slides (applies to the whole conversion; blank = all slides):").pack(side='left')
        self.slides_entry = ttk.Entry(self.slides_frame, textvariable=self.slides_spec, width=24)
        self.slides_entry.pack(side='left', padx=(6, 10))
        ttk.Label(self.slides_frame, text="e.g. 1  or  2-4  or  2,7,8",
                  foreground="#666666").pack(side='left')

        empty_frame = ttk.Frame(self)
        empty_frame.pack(fill='x', **pad)
        ttk.Checkbutton(empty_frame, text="Don't convert empty text boxes (no text)",
                         variable=self.skip_empty_text).pack(side='left')

        run_frame = ttk.Frame(self)
        run_frame.pack(fill='x', **pad)
        self.run_btn = ttk.Button(run_frame, text="Convert", command=self._on_run)
        self.run_btn.pack(side='left')
        self.progress = ttk.Progressbar(run_frame, mode='indeterminate')
        self.progress.pack(side='left', fill='x', expand=True, padx=10)

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill='both', expand=True, **pad)
        self.log_text = tk.Text(log_frame, wrap='word', state='disabled')
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        self.log_text.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        scroll.pack(side='right', fill='y', pady=6)

        self._on_mode_change()

    def _on_mode_change(self):
        self.input_path.set("")
        self.output_path.set("")
        if self.mode.get() == "file":
            self.slides_entry.configure(state='normal')
        else:
            self.slides_spec.set("")
            self.slides_entry.configure(state='disabled')

    def _browse_input(self):
        if self.mode.get() == "file":
            path = filedialog.askopenfilename(
                title="Select a PowerPoint file",
                filetypes=[("PowerPoint files", "*.pptx")]
            )
        else:
            path = filedialog.askdirectory(title="Select input folder")
        if path:
            self.input_path.set(path)
            if not self.output_path.get():
                if self.mode.get() == "file":
                    base, ext = os.path.splitext(path)
                    self.output_path.set(f"{base}_converted{ext}")
                else:
                    self.output_path.set(path.rstrip("/\\") + "_converted")

    def _browse_output(self):
        if self.mode.get() == "file":
            path = filedialog.asksaveasfilename(
                title="Save converted file as",
                defaultextension=".pptx",
                filetypes=[("PowerPoint files", "*.pptx")]
            )
        else:
            path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_path.set(path)

    def _log(self, text):
        self.msg_queue.put(text)

    def _poll_queue(self):
        try:
            while True:
                line = self.msg_queue.get_nowait()
                if line == "__DONE__":
                    self.run_btn.configure(state='normal')
                    self.browse_in_btn.configure(state='normal')
                    self.browse_out_btn.configure(state='normal')
                    self.progress.stop()
                    continue
                self.log_text.configure(state='normal')
                self.log_text.insert('end', line + "\n")
                self.log_text.see('end')
                self.log_text.configure(state='disabled')
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _on_run(self):
        in_path = self.input_path.get().strip()
        out_path = self.output_path.get().strip()

        if not in_path:
            messagebox.showerror("Missing input", "Please choose an input file or folder.")
            return
        if not out_path:
            messagebox.showerror("Missing output", "Please choose where to save the output.")
            return
        if self.mode.get() == "file" and not os.path.isfile(in_path):
            messagebox.showerror("Not found", f"Input file not found:\n{in_path}")
            return
        if self.mode.get() == "folder" and not os.path.isdir(in_path):
            messagebox.showerror("Not found", f"Input folder not found:\n{in_path}")
            return

        mode = self.mode.get()
        suffix = self.suffix.get().strip()
        quiet = self.quiet.get()
        skip_empty = self.skip_empty_text.get()

        slide_numbers = None
        if mode == "file":
            try:
                slide_numbers = parse_slide_spec(self.slides_spec.get())
            except ValueError:
                messagebox.showerror(
                    "Bad slide range",
                    "Couldn't understand the slide list.\n"
                    "Use things like: 1   or   2-4   or   2,7,8"
                )
                return

        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

        self.run_btn.configure(state='disabled')
        self.browse_in_btn.configure(state='disabled')
        self.browse_out_btn.configure(state='disabled')
        self.progress.start(12)

        def worker():
            try:
                if mode == "file":
                    run_single(in_path, out_path, quiet, self._log,
                               slide_numbers=slide_numbers, skip_empty_text=skip_empty)
                else:
                    run_batch(in_path, out_path, suffix, quiet, self._log,
                              skip_empty_text=skip_empty)
            except Exception:
                self._log(traceback.format_exc())
            finally:
                self.msg_queue.put("__DONE__")

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()


def main():
    # Still usable from the command line exactly like before, if someone
    # wants that: ConvertPlaceholder.py INPUT_ROOT OUTPUT_ROOT [--suffix ..] [--quiet]
    if len(sys.argv) > 1:
        import argparse
        ap = argparse.ArgumentParser(
            description="Auto-classify and convert every .pptx under INPUT_ROOT into a mirrored OUTPUT_ROOT folder tree."
        )
        ap.add_argument('input_root')
        ap.add_argument('output_root')
        ap.add_argument('--suffix', default='')
        ap.add_argument('--quiet', action='store_true')
        args = ap.parse_args()
        ok = run_batch(args.input_root, args.output_root, args.suffix, args.quiet, print)
        sys.exit(0 if ok else 1)
    else:
        app = App()
        app.mainloop()


if __name__ == '__main__':
    main()
