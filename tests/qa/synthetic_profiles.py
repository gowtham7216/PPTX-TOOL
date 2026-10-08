"""
SlideGuard Synthetic Test Presentation Profiles (Phase 10).
Generates test presentations across 14 diverse profiles (A through N)
for comprehensive validation of SlideGuard's engine.
"""

import io
from pathlib import Path
from typing import Dict, Any, Optional

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml

from PIL import Image


def _get_temp_png(tmp_dir: Path) -> Path:
    png_path = tmp_dir / "sample_img.png"
    if not png_path.exists():
        im = Image.new("RGBA", (100, 100), (30, 58, 138, 255))
        im.save(png_path, format="PNG")
    return png_path


# Profile A: Normal presentation
def create_profile_normal(output_path: Path) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Profile A: Normal Presentation"

    # Slide 1: Title
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Normal Presentation Title"
    s1.placeholders[1].text = "Subtitle describing standard slides"

    # Slide 2: Content with bullets
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Key Agenda Items"
    tf = s2.placeholders[1].text_frame
    tf.text = "Introduction & Overview"
    p = tf.add_paragraph()
    p.text = "System Architecture and Workflows"
    p2 = tf.add_paragraph()
    p2.text = "Verification and Quality Assurance"

    # Slide 3: Table
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    t_box = s3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    t_box.text_frame.text = "Performance Metrics"
    tbl_shape = s3.shapes.add_table(3, 2, Inches(1), Inches(1.8), Inches(6), Inches(2))
    tbl = tbl_shape.table
    tbl.cell(0, 0).text = "Metric"
    tbl.cell(0, 1).text = "Value"
    tbl.cell(1, 0).text = "Latency"
    tbl.cell(1, 1).text = "< 200ms"
    tbl.cell(2, 0).text = "Throughput"
    tbl.cell(2, 1).text = "100 slides/sec"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile B: Large presentation (100+ slides)
def create_profile_large(output_path: Path, slide_count: int = 105) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Profile B: Large Scaled Presentation"

    blank = prs.slide_layouts[6]
    title_layout = prs.slide_layouts[0]

    # Slide 1: Title
    s1 = prs.slides.add_slide(title_layout)
    s1.shapes.title.text = "Large 100+ Slide Presentation"
    s1.placeholders[1].text = f"Synthetic workload containing {slide_count} slides"

    # Remaining slides: mixed content
    for idx in range(2, slide_count + 1):
        s = prs.slides.add_slide(blank)
        # Title box
        tb = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
        tb.text_frame.text = f"Slide {idx}: Section Overview"

        # Varied shapes per slide
        if idx % 4 == 0:
            # Table slide
            tbl_s = s.shapes.add_table(2, 2, Inches(1), Inches(1.8), Inches(6), Inches(1.5))
            tbl_s.table.cell(0, 0).text = f"Header A-{idx}"
            tbl_s.table.cell(0, 1).text = f"Header B-{idx}"
            tbl_s.table.cell(1, 0).text = "Data 1"
            tbl_s.table.cell(1, 1).text = "Data 2"
        elif idx % 4 == 1:
            # Bullet list box
            cb = s.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(3))
            tf = cb.text_frame
            tf.text = f"Primary item for slide {idx}"
            p = tf.add_paragraph()
            p.text = f"Secondary bullet note for slide {idx}"
        elif idx % 4 == 2:
            # Geometry shapes
            s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(3), Inches(1.5))
            s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5), Inches(2), Inches(3), Inches(1.5))
        else:
            # Mixed text and footer
            mb = s.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
            mb.text_frame.text = f"Standard text commentary on slide {idx}"
            fb = s.shapes.add_textbox(Inches(1), Inches(7.0), Inches(3), Inches(0.4))
            fb.text_frame.text = "INTERNAL AUDIT"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile C: Image-heavy presentation
def create_profile_image_heavy(output_path: Path, tmp_dir: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]
    png_path = _get_temp_png(tmp_dir)

    # 4 slides with images
    for i in range(1, 5):
        s = prs.slides.add_slide(blank)
        t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
        t.text_frame.text = f"Image Gallery Slide {i}"

        # Add 3 images per slide
        s.shapes.add_picture(str(png_path), Inches(1), Inches(1.8), Inches(2.5), Inches(2))
        s.shapes.add_picture(str(png_path), Inches(4), Inches(1.8), Inches(2.5), Inches(2))
        s.shapes.add_picture(str(png_path), Inches(7), Inches(1.8), Inches(2.5), Inches(2))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile D: Text-heavy presentation
def create_profile_text_heavy(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    for i in range(1, 4):
        s = prs.slides.add_slide(blank)
        t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
        t.text_frame.text = f"Text Rich Document Section {i}"

        # 4 separate paragraph boxes in varying reading coordinates
        b1 = s.shapes.add_textbox(Inches(1), Inches(1.5), Inches(4), Inches(2))
        b1.text_frame.text = "Column A paragraph detailing primary considerations and specifications."
        b2 = s.shapes.add_textbox(Inches(5.5), Inches(1.5), Inches(4), Inches(2))
        b2.text_frame.text = "Column B paragraph exploring secondary ramifications and alternatives."
        b3 = s.shapes.add_textbox(Inches(1), Inches(4.0), Inches(8), Inches(2.5))
        tf = b3.text_frame
        tf.text = "Conclusion summary paragraph with extended technical explanations."
        p = tf.add_paragraph()
        p.text = "Additional context on regulatory compliance and standard verification routines."

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile E: Table-heavy presentation
def create_profile_table_heavy(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    for i in range(1, 4):
        s = prs.slides.add_slide(blank)
        t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
        t.text_frame.text = f"Data Matrix {i}"

        # Table 1
        tbl1 = s.shapes.add_table(4, 3, Inches(1), Inches(1.5), Inches(8), Inches(2.2)).table
        tbl1.cell(0, 0).text = "Col 1"
        tbl1.cell(0, 1).text = "Col 2"
        tbl1.cell(0, 2).text = "Col 3"
        for r in range(1, 4):
            for c in range(3):
                tbl1.cell(r, c).text = f"R{r}C{c}"

        # Table 2
        tbl2 = s.shapes.add_table(3, 2, Inches(1), Inches(4.2), Inches(8), Inches(1.8)).table
        tbl2.cell(0, 0).text = "Category"
        tbl2.cell(0, 1).text = "Assessment"
        tbl2.cell(1, 0).text = "Accuracy"
        tbl2.cell(1, 1).text = "99.8%"
        tbl2.cell(2, 0).text = "Reliability"
        tbl2.cell(2, 1).text = "High"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile F: Group-heavy presentation
def create_profile_group_heavy(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    t = s1.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t.text_frame.text = "Shape Groups Demonstration"

    # Create shapes to group
    s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(3), Inches(2))
    s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5), Inches(2), Inches(3), Inches(2))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile G: Nested-group presentation
def create_profile_nested_group(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]
    s = prs.slides.add_slide(blank)
    t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t.text_frame.text = "Nested Elements"
    s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1), Inches(2), Inches(2), Inches(2))
    s.shapes.add_shape(MSO_SHAPE.DIAMOND, Inches(4), Inches(2), Inches(2), Inches(2))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile H: Placeholder-heavy presentation
def create_profile_placeholder_heavy(output_path: Path) -> Path:
    prs = Presentation()

    # Title & Subtitle (layout 0)
    s1 = prs.slides.add_slide(prs.slide_layouts[0])
    s1.shapes.title.text = "Main Title Placeholder"
    s1.placeholders[1].text = "Subtitle Placeholder Text"

    # Title & Content (layout 1)
    s2 = prs.slides.add_slide(prs.slide_layouts[1])
    s2.shapes.title.text = "Body Content Placeholder"
    s2.placeholders[1].text = "Body text inside native layout placeholder"

    # Two Content (layout 3)
    s3 = prs.slides.add_slide(prs.slide_layouts[3])
    s3.shapes.title.text = "Two Column Placeholders"
    s3.placeholders[1].text = "Left column placeholder"
    s3.placeholders[2].text = "Right column placeholder"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile I: List-heavy presentation
def create_profile_list_heavy(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t.text_frame.text = "Multi-Format Lists"

    # Bullet list
    b_box = s.shapes.add_textbox(Inches(1), Inches(1.5), Inches(4), Inches(2.5))
    tf1 = b_box.text_frame
    tf1.text = "- Bullet item alpha"
    p1 = tf1.add_paragraph()
    p1.text = "- Bullet item beta"
    p2 = tf1.add_paragraph()
    p2.text = "- Bullet item gamma"

    # Numbered / Alphabetic list
    n_box = s.shapes.add_textbox(Inches(5.5), Inches(1.5), Inches(4), Inches(2.5))
    tf2 = n_box.text_frame
    tf2.text = "1. Numbered first step"
    p3 = tf2.add_paragraph()
    p3.text = "2. Numbered second step"
    p4 = tf2.add_paragraph()
    p4.text = "3. Numbered third step"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile J: Mixed-language presentation
def create_profile_mixed_language(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    # Slide 1: English
    s1 = prs.slides.add_slide(blank)
    t1 = s1.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t1.text_frame.text = "Global Accessibility Initiative"
    b1 = s1.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(2))
    b1.text_frame.text = "Universal access across multinational organizations."

    # Slide 2: Spanish & French
    s2 = prs.slides.add_slide(blank)
    t2 = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t2.text_frame.text = "Iniciativa de Accesibilidad"
    b2 = s2.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(2))
    b2.text_frame.text = "Le contenu doit être accessible à tous les utilisateurs sans discrimination."

    # Slide 3: German & Japanese
    s3 = prs.slides.add_slide(blank)
    t3 = s3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t3.text_frame.text = "Barrierefreiheit und Standards"
    b3 = s3.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(2))
    b3.text_frame.text = "すべてのユーザーに対してアクセシビリティを保証します。"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile K: Design-heavy presentation
def create_profile_design_heavy(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    # Vibrant colors, custom fonts, styled shapes
    t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    tf = t.text_frame
    p = tf.paragraphs[0]
    p.text = "Design System Showcase"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    sh1 = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1), Inches(2), Inches(3.8), Inches(2.2))
    sh1.fill.solid()
    sh1.fill.fore_color.rgb = RGBColor(0x3B, 0x82, 0xF6)

    sh2 = s.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(5.2), Inches(2), Inches(3.8), Inches(2.2))
    sh2.fill.solid()
    sh2.fill.fore_color.rgb = RGBColor(0x10, 0xB9, 0x81)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile L: Animation-containing presentation
def create_profile_animation(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]

    s = prs.slides.add_slide(blank)
    t = s.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(0.8))
    t.text_frame.text = "Animated Elements Slide"

    rect = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(4), Inches(2))
    rect_id = rect.shape_id

    # Add <p:timing> XML to slide
    timing_xml = f"""
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:tnLst>
            <p:par>
                <p:cTn id="1" dur="indefinite" restart="always">
                    <p:childTnLst>
                        <p:seq concurrent="1" nextAc="seek">
                            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                                <p:childTnLst>
                                    <p:par>
                                        <p:cTn id="3" fill="hold">
                                            <p:stCondLst>
                                                <p:cond delay="0"/>
                                            </p:stCondLst>
                                            <p:childTnLst>
                                                <p:par>
                                                    <p:cTn id="4" presetID="1" presetClass="entr" fill="hold">
                                                        <p:stCondLst>
                                                            <p:cond delay="0"/>
                                                        </p:stCondLst>
                                                    </p:cTn>
                                                </p:par>
                                            </p:childTnLst>
                                        </p:cTn>
                                    </p:par>
                                </p:childTnLst>
                            </p:cTn>
                        </p:seq>
                    </p:childTnLst>
                </p:cTn>
            </p:par>
        </p:tnLst>
        <p:bldLst>
            <p:bldP spid="{rect_id}" grpId="0" build="all"/>
        </p:bldLst>
    </p:timing>
    """
    s._element.append(parse_xml(timing_xml))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile M: AI-authoring candidate presentation
def create_profile_ai_candidate(output_path: Path, tmp_dir: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]
    png_path = _get_temp_png(tmp_dir)

    # Slide missing title + image without alt text + low contrast text
    s = prs.slides.add_slide(blank)
    # Image without alt text
    s.shapes.add_picture(str(png_path), Inches(1), Inches(1.5), Inches(3), Inches(2))

    # Low contrast text box
    tb = s.shapes.add_textbox(Inches(4.5), Inches(1.5), Inches(5), Inches(2))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = "Subtle gray commentary needing AI clarity."
    p.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p.font.size = Pt(8)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path


# Profile N: Minimal / empty presentation
def create_profile_minimal(output_path: Path) -> Path:
    prs = Presentation()
    blank = prs.slide_layouts[6]
    # Slide 1: completely empty slide
    prs.slides.add_slide(blank)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path
