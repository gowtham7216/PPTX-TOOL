"""
SlideGuard Test PPTX Fixture Generator.
Programmatically generates test PowerPoint presentations with controlled accessibility attributes.
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE


def create_comprehensive_test_pptx(output_path: Path) -> Path:
    """
    Creates a 6-slide test presentation containing diverse shapes,
    groups, footers, tables, and accessibility variations.
    """
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Core Properties
    prs.core_properties.title = "SlideGuard Test Presentation"
    prs.core_properties.author = "SlideGuard Test Suite"

    blank_layout = prs.slide_layouts[6]
    title_layout = prs.slide_layouts[0]

    # --- SLIDE 1: Title Slide ---
    s1 = prs.slides.add_slide(title_layout)
    s1.shapes.title.text = "Introduction to SlideGuard"
    s1.placeholders[1].text = "Accessibility and Integrity Verification Engine"

    # --- SLIDE 2: Simple Text + Content Slide ---
    s2 = prs.slides.add_slide(blank_layout)
    # Manual Title Box
    t_box = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    t_tf = t_box.text_frame
    t_tf.text = "Key Objectives"
    # Content Box with simulated bullets
    c_box = s2.shapes.add_textbox(Inches(1), Inches(1.8), Inches(8), Inches(3))
    c_tf = c_box.text_frame
    p1 = c_tf.paragraphs[0]
    p1.text = "- First objective: analyze reading order"
    p2 = c_tf.add_paragraph()
    p2.text = "- Second objective: protect visual design"

    # --- SLIDE 3: Footer slide with "WILEY" text ---
    s3 = prs.slides.add_slide(blank_layout)
    s3_title = s3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    s3_title.text_frame.text = "Slide with Footer Content"
    # Main body
    s3_body = s3.shapes.add_textbox(Inches(1), Inches(2.0), Inches(8), Inches(2))
    s3_body.text_frame.text = "This slide contains corporate footer branding."
    # Footer placed in bottom margin (7.0 inches down on 7.5 inch slide)
    footer_box = s3.shapes.add_textbox(Inches(1), Inches(7.0), Inches(3), Inches(0.4))
    footer_box.text_frame.text = "WILEY"

    # --- SLIDE 4: Slide with Table ---
    s4 = prs.slides.add_slide(blank_layout)
    s4_title = s4.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    s4_title.text_frame.text = "Compliance Metrics Table"
    # 3 rows, 2 columns table
    tbl_shape = s4.shapes.add_table(3, 2, Inches(1), Inches(2.0), Inches(6), Inches(2))
    tbl = tbl_shape.table
    tbl.cell(0, 0).text = "Criterion"
    tbl.cell(0, 1).text = "Status"
    tbl.cell(1, 0).text = "WCAG 1.3.1"
    tbl.cell(1, 1).text = "Compliant"
    tbl.cell(2, 0).text = "WCAG 2.4.2"
    tbl.cell(2, 1).text = "Compliant"

    # --- SLIDE 5: Slide missing title + empty text box + low contrast text ---
    s5 = prs.slides.add_slide(blank_layout)
    # No title shape!
    # Empty text box
    empty_box = s5.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    empty_box.text_frame.text = ""  # empty
    # Low contrast text box
    lc_box = s5.shapes.add_textbox(Inches(1), Inches(3), Inches(6), Inches(1))
    lc_tf = lc_box.text_frame
    lc_p = lc_tf.paragraphs[0]
    lc_p.text = "Hard to read low contrast text"
    lc_run = lc_p.runs[0]
    lc_run.font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    lc_run.font.size = Pt(8)  # very small font

    # --- SLIDE 6: Group of shapes ---
    s6 = prs.slides.add_slide(blank_layout)
    s6_title = s6.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    s6_title.text_frame.text = "Grouped Elements Slide"
    # Auto-shapes
    s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(3), Inches(2))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path
