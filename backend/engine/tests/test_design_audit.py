"""Appendix 1 design checks: palette, contrast, guides, margins, pictures, brand, fill and chart axes."""
import io

from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

from engine.design_audit import content_area, contrast, design_issues
from engine.ingest import inspect_template
from engine.native_data import _axis_title
from engine.models import SlideContent


def rules(issues, rule=None):
    found = [(issue["slide_id"], issue["rule_id"], issue["severity"]) for issue in issues]
    return [item for item in found if rule is None or item[1] == rule]


def blank_template(tmp_path):
    """A template whose own text sits in the default title and content placeholders."""
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[1])
    slide.shapes.title.text = "Заголовок шаблона"
    slide.placeholders[1].text = "Текст шаблона"
    path = tmp_path / "template.pptx"
    deck.save(path)
    return path


def text_box(slide, text, left, top, width, height, color=None, size=18):
    box = slide.shapes.add_textbox(left, top, width, height)
    box.text_frame.text = text
    run = box.text_frame.paragraphs[0].runs[0]
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return box


def picture(width_px=200, height_px=100):
    stream = io.BytesIO()
    Image.new("RGB", (width_px, height_px), (40, 90, 160)).save(stream, "PNG")
    stream.seek(0)
    return stream


def test_foreign_colours_and_low_contrast_are_found_and_template_colours_pass(tmp_path):
    template = blank_template(tmp_path)
    deck = Presentation(template)
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    text_box(slide, "Главный тезис", Inches(0.5), Inches(1.5), Inches(9), Inches(1), color="FF00FF")
    text_box(slide, "Бледный тезис", Inches(0.5), Inches(3), Inches(9), Inches(1), color="C8C8C8", size=14)
    text_box(slide, "Читаемый тезис", Inches(0.5), Inches(4.5), Inches(9), Inches(1), color="1F1F1F")
    path = tmp_path / "deck.pptx"
    deck.save(path)
    slides = [SlideContent("slide_1", "Заголовок шаблона", ["Текст шаблона"], []),
              SlideContent("slide_2", "Главный тезис", ["Бледный тезис", "Читаемый тезис"], [])]
    issues = design_issues(path, template, slides)
    palette = [issue for issue in issues if issue["rule_id"] == "off_palette_color"]
    assert [issue["slide_id"] for issue in palette] == ["slide_2"] and "#FF00FF" in palette[0]["evidence"]
    low = [issue for issue in issues if issue["rule_id"] == "low_contrast"]
    # #C8C8C8 on white is 1.7:1; magenta at 18 pt passes the large-text 3:1 rule; near-black passes.
    assert [(issue["slide_id"], issue["severity"]) for issue in low] == [("slide_2", "blocking")]
    assert "на #FFFFFF" in low[0]["evidence"]
    # The dark palette of variant B is part of the approved design.
    approved = design_issues(path, template, slides, extra_colors={"accent": "#FF00FF"})
    assert rules(approved, "off_palette_color") == []
    assert round(contrast((0, 0, 0), (255, 255, 255)), 1) == 21.0


def test_blocks_off_the_grid_and_in_the_margins_are_reported(tmp_path):
    template = blank_template(tmp_path)
    left, top, right, bottom = content_area(Presentation(template))
    assert left == Inches(0.5) and right == Inches(9.5)
    deck = Presentation(template)
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    text_box(slide, "Ровный блок", Inches(0.5), Inches(1.5), Inches(9), Inches(1))
    text_box(slide, "Сдвинутый блок", Inches(2.1), Inches(3), Inches(3.3), Inches(1))
    # Its right edge is on the grid; its left edge enters the margin.
    text_box(slide, "Блок у края", Inches(0.05), Inches(4.5), Inches(9.45), Inches(1))
    path = tmp_path / "deck.pptx"
    deck.save(path)
    slides = [SlideContent("slide_1", "Заголовок шаблона", ["Текст шаблона"], []),
              SlideContent("slide_2", "Ровный блок", ["Сдвинутый блок", "Блок у края"], [])]
    issues = design_issues(path, template, slides)
    misaligned = [issue["evidence"] for issue in issues if issue["rule_id"] == "misaligned_block"]
    margins = [issue["evidence"] for issue in issues if issue["rule_id"] == "margin_intrusion"]
    assert len(misaligned) == 1 and "TextBox 2" in misaligned[0]
    assert len(margins) == 1 and "у левого края на 1,1 см" in margins[0]


def test_stretched_picture_and_moved_brand_element(tmp_path):
    deck = Presentation()
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    logo = slide.shapes.add_picture(picture(), Inches(8), Inches(0.3), Inches(1.6), Inches(0.8))
    template = tmp_path / "template.pptx"
    deck.save(template)
    prepared = inspect_template(template)
    prepared.compositions[0].object_roles = {str(logo.shape_id): "fixed_brand"}

    generated = Presentation(template)
    moved = generated.slides[0].shapes[0]
    moved.left += Inches(1)
    generated.slides[0].shapes.add_picture(picture(), Inches(1), Inches(2), Inches(6), Inches(1))
    generated.slides[0].shapes.add_picture(picture(), Inches(1), Inches(4), Inches(2), Inches(1))
    path = tmp_path / "deck.pptx"
    generated.save(path)
    issues = design_issues(path, template, [SlideContent("slide_1", "", [], [])], prepared, source_slides=[1])
    assert rules(issues, "brand_moved") == [("slide_1", "brand_moved", "warning")]
    assert "2,5 см" in next(issue["evidence"] for issue in issues if issue["rule_id"] == "brand_moved")
    distorted = [issue for issue in issues if issue["rule_id"] == "image_distorted"]
    # 6x1 in shows a 2:1 image three times wider; 2x1 in keeps it; the template's logo is its own design.
    assert len(distorted) == 1 and "растянуто по ширине на 200%" in distorted[0]["evidence"]
    assert rules(design_issues(template, template, [SlideContent("slide_1", "", [], [])], prepared, [1])) == []


def test_slide_fill_outside_a_quarter_to_three_quarters(tmp_path):
    template = blank_template(tmp_path)
    deck = Presentation(template)
    sparse = deck.slides.add_slide(deck.slide_layouts[6])
    text_box(sparse, "Короткий тезис", Inches(0.5), Inches(0.5), Inches(3), Inches(0.6))
    dense = deck.slides.add_slide(deck.slide_layouts[6])
    dense.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.4), Inches(9), Inches(6.4))
    text_box(dense, "Плотный тезис", Inches(0.6), Inches(0.5), Inches(8), Inches(1))
    balanced = deck.slides.add_slide(deck.slide_layouts[6])
    balanced.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(9), Inches(3.5))
    text_box(balanced, "Тезис в карточке", Inches(0.8), Inches(1.8), Inches(8), Inches(1))
    path = tmp_path / "deck.pptx"
    deck.save(path)
    slides = [SlideContent("slide_1", "Заголовок шаблона", [], []),
              SlideContent("slide_2", "", ["Короткий тезис"], []),
              SlideContent("slide_3", "", ["Плотный тезис"], []),
              SlideContent("slide_4", "", ["Тезис в карточке"], [])]
    issues = design_issues(path, template, slides)
    assert rules(issues, "slide_underfilled") == [("slide_2", "slide_underfilled", "warning")]
    assert rules(issues, "slide_overfilled") == [("slide_3", "slide_overfilled", "warning")]


def test_chart_needs_axis_labels_and_units(tmp_path):
    template = blank_template(tmp_path)
    deck = Presentation(template)
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    data = CategoryChartData()
    data.categories = ["Январь", "Февраль"]
    data.add_series("Проекты", (12, 18))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(1), Inches(4.5), Inches(4), data)
    labelled = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(5), Inches(1), Inches(4.5), Inches(4), data)
    labelled.name = "Labelled chart"
    _axis_title(labelled.chart.category_axis, "Месяц")
    _axis_title(labelled.chart.value_axis, "Проекты, шт.")
    path = tmp_path / "deck.pptx"
    deck.save(path)
    slides = [SlideContent("slide_1", "Заголовок шаблона", ["Текст шаблона"], []),
              SlideContent("slide_2", "", [], [])]
    issues = [issue for issue in design_issues(path, template, slides) if issue["rule_id"] == "chart_axis_unlabeled"]
    assert len(issues) == 1 and "нет подписи с единицами" in issues[0]["evidence"]
    labelled = next(shape for shape in Presentation(path).slides[1].shapes if shape.name == "Labelled chart")
    assert issues[0]["object_id"] != str(labelled.shape_id)
