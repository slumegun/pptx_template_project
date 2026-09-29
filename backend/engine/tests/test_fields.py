"""Field Marker and Field Checker: free areas measured on the slide without text, letters checked in the PDF."""
import shutil
from io import BytesIO

import pytest
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from engine.export import export_pdf
from engine.fields import largest_free_rectangle, mark_fields
from engine.ingest import inspect_template
from engine.models import SlideContent
from engine.rendered_audit import audit_pdf
from engine.renderer import render_variant, strict_regions


def libreoffice():
    from engine.export import _find_soffice
    try:
        _find_soffice()
    except RuntimeError:
        if not (shutil.which("soffice") or shutil.which("soffice.com")):
            pytest.skip("LibreOffice is not installed")


def picture(slide, left, top, width, height):
    stream = BytesIO()
    Image.new("RGB", (60, 60), (200, 30, 60)).save(stream, format="PNG")
    return slide.shapes.add_picture(stream, Inches(left), Inches(top), Inches(width), Inches(height))


def textbox(slide, left, top, width, height, text, size=18):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.text_frame.word_wrap = True
    box.text_frame.text = text
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(size)
    return box


def test_largest_free_rectangle_skips_the_drawn_cells():
    free, ink = True, False
    grid = [[free, free, free, ink],
            [free, free, free, ink],
            [ink, free, free, free]]
    assert largest_free_rectangle(grid) == (0, 0, 3, 2)


def test_field_marker_measures_the_free_part_of_each_field(tmp_path):
    libreoffice()
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    heading = textbox(slide, 0.5, 0.4, 9, 0.8, "Заголовок — ключевая мысль слайда", 28)
    half = textbox(slide, 0.5, 1.6, 8, 2, "Описание, наполовину закрытое картинкой справа", 16)
    covered = textbox(slide, 0.5, 4.3, 4, 1.4, "Подпись под картинкой целиком", 16)
    photo = picture(slide, 4.5, 1.6, 4, 2)
    picture(slide, 0.4, 4.2, 4.2, 1.6)
    source = tmp_path / "template.pptx"
    deck.save(source)
    template = inspect_template(source)
    mark_fields(source, template, tmp_path / "work")
    slots = {slot.shape_id: slot for slot in template.compositions[0].slots}
    assert slots[heading.shape_id].clear_share > 0.95
    partly = slots[half.shape_id]
    assert 0.35 < partly.clear_share < 0.65
    # The free area ends where the picture starts (one grid cell of tolerance).
    assert partly.clear_box[0] + partly.clear_box[2] <= photo.left + Inches(8) / 48 * 1.5
    assert slots[covered.shape_id].clear_share < 0.5
    _, regions = strict_regions(template.compositions[0], template.width, template.height)
    assert covered.shape_id not in {slot.shape_id for slot in regions}
    assert half.shape_id in {slot.shape_id for slot in regions}
    # The second call reads the cache instead of rendering again.
    assert mark_fields(source, template, tmp_path / "work")


def test_strict_text_stays_in_the_free_area_by_inner_margins_only(tmp_path):
    libreoffice()
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    textbox(slide, 0.5, 0.4, 9, 0.8, "Заголовок — ключевая мысль слайда", 28)
    half = textbox(slide, 0.5, 1.6, 8, 2, "Описание, наполовину закрытое картинкой справа", 16)
    picture(slide, 4.5, 1.6, 4, 2)
    source = tmp_path / "template.pptx"
    deck.save(source)
    template = inspect_template(source)
    mark_fields(source, template, tmp_path / "work")
    content = [SlideContent("slide_1", "Новый заголовок", ["Короткий тезис слева"], ["f1"])]
    target = tmp_path / "deck.pptx"
    render_variant(source, template, content, 0, target, strict=True, plan=[0])
    written = next(shape for shape in Presentation(target).slides[0].shapes if shape.shape_id == half.shape_id)
    assert (written.left, written.top, written.width, written.height) == (half.left, half.top, half.width, half.height)
    assert written.text_frame.margin_right > Inches(3)


def test_field_checker_finds_text_covered_from_above_and_text_outside_its_free_area(tmp_path):
    libreoffice()
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    under = textbox(slide, 0.5, 0.5, 8, 1, "Текст, который закрыт фигурой сверху", 24)
    cover = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3), Inches(0.4), Inches(3), Inches(1.2))
    cover.fill.solid()
    cover.fill.fore_color.rgb = RGBColor(0, 119, 255)
    free = textbox(slide, 0.5, 3, 8, 1, "Этот текст длиннее своей свободной области поля", 24)
    source = tmp_path / "deck.pptx"
    deck.save(source)
    pdf = export_pdf(source, tmp_path / "pdf")
    fields = {0: {free.shape_id: [free.left, free.top, int(Inches(3)), free.height]}}
    issues = audit_pdf(source, pdf, None, fields)
    found = {(issue["rule_id"], issue["object_id"]) for issue in issues}
    assert ("field_text_covered", str(under.shape_id)) in found
    assert ("field_text_outside_clear_area", str(free.shape_id)) in found
    assert ("field_text_covered", str(free.shape_id)) not in found
