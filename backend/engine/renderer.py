"""Deterministic, editable PPTX assembly from one uploaded template.

Only the current template contributes layouts and decorative objects. Existing
slide content is scrubbed before user facts are inserted.
"""

from __future__ import annotations

from collections import Counter
from contextvars import ContextVar
from dataclasses import dataclass
from copy import deepcopy
from types import SimpleNamespace
from io import BytesIO
import math
from pathlib import Path
import re
from typing import Iterable

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL, MSO_THEME_COLOR
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

from .fonts import effective_font, embed_fonts, theme_fonts, used_families
from .models import Composition, PreparedTemplate, SlideContent
from .native_data import NumericCsvData, add_native_data_visuals
from .typography import text_height, resolve_font
from .pagination import BODY_PT, body_fits, dense_geometry

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DARK_PALETTE = {
    "background": "#101927",
    "surface": "#23364B",
    "text": "#F5F8FC",
    "accent": "#70D3C5",
}


def _luminance(color: RGBColor) -> float:
    def linear(channel: int) -> float:
        value = channel / 255
        return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4

    return sum(weight * linear(channel) for weight, channel in zip((0.2126, 0.7152, 0.0722), color))


def _contrast(first: RGBColor, second: RGBColor) -> float:
    light, dark = sorted((_luminance(first), _luminance(second)), reverse=True)
    return (light + 0.05) / (dark + 0.05)


def _dark_palette(value: dict[str, str] | None) -> dict[str, RGBColor]:
    """Use a model-proposed palette only when it is safe for editable text."""
    def decode(source: dict[str, str]) -> dict[str, RGBColor]:
        result = {}
        for key in DARK_PALETTE:
            color = source[key]
            if not isinstance(color, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
                raise ValueError(f"Invalid dark palette color: {key}")
            result[key] = RGBColor.from_string(color[1:])
        if _luminance(result["background"]) > 0.12:
            raise ValueError("Dark palette background is too bright")
        if _contrast(result["text"], result["background"]) < 4.5:
            raise ValueError("Dark palette text has insufficient contrast on background")
        if _contrast(result["text"], result["surface"]) < 4.5:
            raise ValueError("Dark palette text has insufficient contrast on surface")
        if _contrast(result["accent"], result["background"]) < 4.5:
            raise ValueError("Dark palette accent has insufficient contrast on background")
        if _contrast(result["accent"], result["surface"]) < 4.5:
            raise ValueError("Dark palette accent has insufficient contrast on surface")
        return result

    if value is not None:
        try:
            return decode(value)
        except (KeyError, TypeError, ValueError):
            # A bad model suggestion must not make the exported deck unreadable.
            pass
    return decode(DARK_PALETTE)



def _is_template_accent(shape, slide_area: int) -> bool:
    """Keep small colored template details, but replace pale panels."""
    if shape.width * shape.height > slide_area * 0.20:
        return False
    try:
        if shape.fill.type != MSO_FILL.SOLID:
            return False
        color = shape.fill.fore_color
        if color.type == MSO_COLOR_TYPE.SCHEME:
            return color.theme_color in {
                MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2,
                MSO_THEME_COLOR.ACCENT_3, MSO_THEME_COLOR.ACCENT_4,
                MSO_THEME_COLOR.ACCENT_5, MSO_THEME_COLOR.ACCENT_6,
            }
        if color.type == MSO_COLOR_TYPE.RGB:
            channels = tuple(color.rgb)
            return max(channels) - min(channels) >= 55
    except (AttributeError, TypeError, ValueError):
        return False
    return False


def _set_chart_background(element, color: RGBColor) -> None:
    """Set an OOXML chart fill without flattening its data workbook."""
    previous = element.find(qn("c:spPr"))
    if previous is not None:
        element.remove(previous)
    properties = OxmlElement("c:spPr")
    fill = OxmlElement("a:solidFill")
    value = OxmlElement("a:srgbClr")
    value.set("val", str(color))
    fill.append(value)
    properties.append(fill)
    element.insert_element_before(properties, "c:txPr", "c:externalData", "c:extLst")


def _style_dark_native_data(slide, palette: dict[str, RGBColor]) -> None:
    """Keep the native chart and table editable while matching the dark slide."""
    for shape in slide.shapes:
        if shape.has_table:
            for row_number, row in enumerate(shape.table.rows):
                for column_number, _ in enumerate(shape.table.columns):
                    cell = shape.table.cell(row_number, column_number)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = (
                        palette["accent"] if row_number == 0 else
                        palette["surface"] if row_number % 2 else palette["background"]
                    )
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.font.color.rgb = palette["background"] if row_number == 0 else palette["text"]
                        for run in paragraph.runs:
                            run.font.color.rgb = palette["background"] if row_number == 0 else palette["text"]
        elif shape.has_chart:
            chart = shape.chart
            _set_chart_background(chart._chartSpace, palette["surface"])
            _set_chart_background(chart._chartSpace.chart.plotArea, palette["surface"])
            for axis in (chart.category_axis, chart.value_axis):
                axis.tick_labels.font.color.rgb = palette["text"]
                axis.format.line.color.rgb = palette["text"]
                if axis.has_major_gridlines:
                    axis.major_gridlines.format.line.color.rgb = palette["background"]
            if chart.has_legend:
                chart.legend.font.color.rgb = palette["text"]
            if chart.series:
                chart.series[0].format.fill.solid()
                chart.series[0].format.fill.fore_color.rgb = palette["accent"]
        elif (
            shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            and shape.width <= Inches(0.8)
            and shape.height <= Inches(0.1)
        ):
            # The native data slide's short rule uses the proposed accent.
            shape.fill.solid()
            shape.fill.fore_color.rgb = palette["accent"]



def _brand_needs_light_plate(shape, background: RGBColor, slide_area: int) -> bool:
    # Large transparent product artwork can contain dark line art that vanishes
    # on a dark canvas; give it a light field as well as compact brand marks.
    limit = 0.55 if shape.shape_type == MSO_SHAPE_TYPE.PICTURE else 0.10
    if shape.width * shape.height > slide_area * limit:
        return False
    if shape.has_text_frame and shape.text.strip():
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                color = run.font.color
                if color.type == MSO_COLOR_TYPE.RGB:
                    if _contrast(color.rgb, background) < 4.5:
                        return True
                elif (
                    color.type == MSO_COLOR_TYPE.SCHEME
                    and color.theme_color == MSO_THEME_COLOR.BACKGROUND_1
                ):
                    continue
                else:
                    # The default template text color is commonly dark.
                    return True
        return False
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        try:
            with Image.open(BytesIO(shape.image.blob)) as image:
                if image.width * image.height > 16_000_000:
                    return False
                image.thumbnail((128, 128))
                rgba = image.convert("RGBA")
                pixels = [
                    rgba.getpixel((x, y))
                    for y in range(rgba.height)
                    for x in range(rgba.width)
                ]
                if not pixels or all(alpha == 255 for _, _, _, alpha in pixels):
                    return False
                colored = [
                    RGBColor(red, green, blue)
                    for red, green, blue, alpha in pixels[::max(1, len(pixels) // 2048)]
                    if alpha >= 128
                ]
                return bool(colored) and sum(_luminance(color) for color in colored) / len(colored) < 0.25
        except (OSError, ValueError):
            return False
    return False


def _add_brand_plates(slide, protected: set[int], background: RGBColor, width: int, height: int) -> None:
    """Keep a dark text/logo legible without recoloring the brand object."""
    slide_area = width * height
    for brand in list(slide.shapes):
        if brand.shape_id not in protected or not _brand_needs_light_plate(brand, background, slide_area):
            continue
        pad = int(Pt(6))
        left = max(0, brand.left - pad)
        top = max(0, brand.top - pad)
        right = min(width, brand.left + brand.width + pad)
        bottom = min(height, brand.top + brand.height + pad)
        plate = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, right - left, bottom - top)
        plate.name = "Brand contrast plate"
        plate.fill.solid()
        plate.fill.fore_color.rgb = RGBColor(245, 248, 252)
        plate.line.fill.background()
        brand.element.addprevious(plate.element)
    # Layout artwork sits behind slide shapes. A slide-level plate would hide it,
    # so repeat the exact embedded image above the plate after covering the
    # low-contrast inherited copy.
    for brand in slide.slide_layout.shapes:
        if (brand.shape_type != MSO_SHAPE_TYPE.PICTURE
                or not _brand_needs_light_plate(brand, background, slide_area)):
            continue
        pad = int(Pt(6))
        left = max(0, brand.left - pad)
        top = max(0, brand.top - pad)
        right = min(width, brand.left + brand.width + pad)
        bottom = min(height, brand.top + brand.height + pad)
        plate = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, right - left, bottom - top)
        plate.name = "Layout brand contrast plate"
        plate.fill.solid()
        plate.fill.fore_color.rgb = RGBColor(245, 248, 252)
        plate.line.fill.background()
        _clone_shape(slide.slide_layout, slide, brand)

def _apply_dark_theme(
    slide,
    content: SlideContent,
    palette: dict[str, RGBColor],
    protected: set[int],
    width: int,
    height: int,
    *,
    native_data: bool = False,
) -> None:
    """Change only editable slide surfaces and text; source images stay intact."""
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = palette["background"]
    slide_area = width * height
    for shape in slide.shapes:
        if (
            shape.shape_id in protected
            or shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP}
            or shape.has_chart or shape.has_table
        ):
            continue
        try:
            if shape.fill.type == MSO_FILL.SOLID:
                if shape.has_text_frame and shape.text.strip():
                    shape.fill.fore_color.rgb = palette["surface"]
                elif shape.width * shape.height > slide_area * 0.35:
                    shape.fill.fore_color.rgb = palette["background"]
                elif not _is_template_accent(shape, slide_area):
                    shape.fill.fore_color.rgb = palette["surface"]
        except (AttributeError, TypeError, ValueError):
            pass
        if shape.has_text_frame and shape.text.strip():
            color = palette["accent"] if shape.text.strip() == content.title else palette["text"]
            for paragraph in shape.text_frame.paragraphs:
                paragraph.font.color.rgb = color
                for run in paragraph.runs:
                    run.font.color.rgb = color
    if native_data:
        _style_dark_native_data(slide, palette)
    else:
        _add_brand_plates(slide, protected, palette["background"], width, height)



def _resolved_color(color, slide) -> RGBColor | None:
    if color.type == MSO_COLOR_TYPE.RGB:
        return color.rgb
    if color.type == MSO_COLOR_TYPE.SCHEME:
        names = {MSO_THEME_COLOR.BACKGROUND_1: "lt1", MSO_THEME_COLOR.BACKGROUND_2: "lt2",
                 MSO_THEME_COLOR.TEXT_1: "dk1", MSO_THEME_COLOR.TEXT_2: "dk2",
                 MSO_THEME_COLOR.LIGHT_1: "lt1", MSO_THEME_COLOR.LIGHT_2: "lt2",
                 MSO_THEME_COLOR.DARK_1: "dk1", MSO_THEME_COLOR.DARK_2: "dk2"}
        names.update({getattr(MSO_THEME_COLOR, f"ACCENT_{i}"): f"accent{i}" for i in range(1, 7)})
        if color.theme_color in names:
            return _theme_color(slide, names[color.theme_color], "000000")
    return None


def _solid_fill_rgb(shape) -> RGBColor | None:
    try:
        if shape.fill.type != MSO_FILL.SOLID:
            return None
        color = shape.fill.fore_color
        if color.type == MSO_COLOR_TYPE.RGB:
            return color.rgb
        if color.type == MSO_COLOR_TYPE.SCHEME:
            # Any theme colour (lt1, accent2...) resolves through the slide's theme;
            # an unresolved card fill made light text land on a white card.
            try:
                resolved = _resolved_color(color, shape.part.slide)
            except (AttributeError, KeyError, ValueError):
                resolved = None
            if resolved is not None:
                return resolved
            if color.theme_color == MSO_THEME_COLOR.BACKGROUND_1:
                return RGBColor(255, 255, 255)
            if color.theme_color == MSO_THEME_COLOR.TEXT_1:
                return RGBColor(18, 24, 35)
    except (AttributeError, TypeError, ValueError):
        pass
    return None


def _overlap_fraction(cover, text_shape) -> float:
    intersection = (
        max(0, min(cover.left + cover.width, text_shape.left + text_shape.width)
            - max(cover.left, text_shape.left))
        * max(0, min(cover.top + cover.height, text_shape.top + text_shape.height)
              - max(cover.top, text_shape.top))
    )
    return intersection / max(1, text_shape.width * text_shape.height)


def _text_surface(slide, text_shape, fallback: RGBColor) -> tuple[RGBColor, bool]:
    """Return the visible solid surface and whether an image makes it uncertain."""
    background = fallback
    for owner in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        if owner.background.fill.type == MSO_FILL.SOLID:
            background = _resolved_color(owner.background.fill.fore_color, slide) or fallback
            break
    uncertain = False
    for shape in slide.shapes:
        if shape.shape_id == text_shape.shape_id:
            break
        fraction = _overlap_fraction(shape, text_shape)
        if fraction < 0.30:
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            uncertain = True
        elif fraction >= 0.90:
            fill = _solid_fill_rgb(shape)
            if fill is not None:
                background, uncertain = fill, False
    own_fill = _solid_fill_rgb(text_shape)
    if own_fill is not None:
        background, uncertain = own_fill, False
    return background, uncertain


def _ensure_text_contrast(slide, fallback: RGBColor, protected: set[int] | None = None) -> tuple[int, float, bool]:
    """Give every generated text run at least WCAG 4.5:1 on its local surface."""
    from .design_audit import _Theme, _text_color

    theme = _Theme(slide.slide_layout.slide_master)
    changed = 0
    minimum = float("inf")
    satisfied = True
    for shape in list(slide.shapes):
        if shape.shape_id in (protected or set()) or not shape.has_text_frame or not shape.text.strip():
            continue
        background, uncertain = _text_surface(slide, shape, fallback)
        if uncertain:
            # A solid text-box surface is needed over photography or raster art.
            background = fallback
            shape.fill.solid()
            shape.fill.fore_color.rgb = background
            changed += 1
        candidates = (RGBColor(20, 27, 39), RGBColor(248, 250, 252))
        safe = max(candidates, key=lambda color: _contrast(color, background))
        if _contrast(safe, background) < 4.5:
            safe = max((RGBColor(0, 0, 0), RGBColor(255, 255, 255)),
                       key=lambda color: _contrast(color, background))
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                current = _resolved_color(run.font.color, slide)
                if current is None:
                    # The colour the text inherits from its placeholder, layout or master.
                    inherited = _text_color(shape, paragraph, run, theme)
                    current = RGBColor(*inherited) if inherited is not None else None
                # The template's own colour stays while it keeps 3:1 (WCAG large-text level): a brand
                # white on the brand blue is the design, and the Design Inspector reports text under
                # 4.5:1. Unknown or unreadable colours take the safe one.
                if current is None or _contrast(current, background) < 3.0:
                    run.font.color.rgb = safe
                    changed += 1
                    current = safe
                ratio = _contrast(current, background)
                minimum = min(minimum, ratio)
                satisfied = satisfied and ratio >= 3.0
    return changed, (minimum if minimum != float("inf") else 0.0), satisfied

def _has_content_in_group(shape) -> bool:
    if shape.shape_type != MSO_SHAPE_TYPE.GROUP:
        return False
    for child in shape.shapes:
        if child.has_text_frame and child.text.strip():
            return True
        if child.has_table or child.has_chart or _has_content_in_group(child):
            return True
    return False


def _should_copy(shape, slide_area: int, role: str | None = None) -> bool:
    if role in {"remove", "unresolved"}:
        return False
    if shape.has_chart or shape.has_table:
        return False
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return role in {"fixed_brand", "reusable_asset"}
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP and _has_content_in_group(shape) and role != "fixed_brand":
        return False
    # Embedded objects and media can carry old content, so leave them out.
    if shape.shape_type in {MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.MEDIA}:
        return False
    return True


def _clone_shape(source_slide, target_slide, shape) -> None:
    element = deepcopy(shape.element)
    # Some imported PPTX files contain impossible shadow values (including
    # scientific notation in integer OOXML fields). PowerPoint repairs those
    # copied shapes even though python-pptx and LibreOffice accept the ZIP.
    for shadow in list(element.iter(qn("a:outerShdw"))):
        try:
            valid = (0 <= int(shadow.get("blurRad", "0")) <= 9223372036854775807
                     and 0 <= int(shadow.get("dist", "0")) <= 9223372036854775807
                     and 0 <= int(shadow.get("dir", "0")) <= 21600000
                     and all(0 <= int(alpha.get("val", "100000")) <= 100000
                             for alpha in shadow.iter(qn("a:alpha"))))
        except ValueError:
            valid = False
        if not valid:
            shadow.getparent().remove(shadow)
    for transform in element.iter(qn("a:xfrm")):
        extent = transform.find(qn("a:ext"))
        if extent is not None:
            for dimension in ("cx", "cy"):
                if int(extent.get(dimension, "0")) < 0:
                    extent.set(dimension, "0")
    # A reused visual object must not carry an old, possibly external action.
    for action in list(element.iter(qn("a:hlinkClick"))) + list(element.iter(qn("a:hlinkMouseOver"))):
        action.getparent().remove(action)
    for child in element.iter():
        for attribute, old_id in list(child.attrib.items()):
            if attribute.startswith("{" + REL_NS + "}"):
                try:
                    relation = source_slide.part.rels[old_id]
                except KeyError:
                    continue
                target = relation.target_ref if relation.is_external else relation.target_part
                child.set(attribute, target_slide.part.relate_to(target, relation.reltype, relation.is_external))
    target_slide.shapes._spTree.insert_element_before(element, "p:extLst")


def _template_typography(presentation) -> dict[str, str | int | bool | None]:
    """The template's typefaces and typical title/body sizes for regions the renderer adds.

    Text written into a template object keeps that object's own face and size; these
    values serve only new regions. They are read the way the template's text is really
    drawn (run, list styles, layout, master, theme), so a template that relies on theme
    fonts or placeholder styles gets its own typography, never a generic family.
    """
    body_names: Counter[str] = Counter()
    title_names: Counter[str] = Counter()
    titles: list[int] = []
    bodies: list[int] = []
    bold_titles = 0
    for slide in presentation.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if not run.text.strip():
                        continue
                    family, size, bold = effective_font(shape, paragraph, run)
                    size = round(size)
                    if 28 <= size <= 72:
                        titles.append(size)
                        bold_titles += bold
                        if family:
                            title_names[family] += len(run.text.strip())
                    elif 12 <= size <= 27:
                        bodies.append(size)
                        if family:
                            body_names[family] += len(run.text.strip())

    def median(values: list[int], default: int) -> int:
        if not values:
            return default
        values.sort()
        return values[len(values) // 2]

    major, minor = theme_fonts(presentation.slide_master)
    body_font = body_names.most_common(1)[0][0] if body_names else minor
    return {
        "font": body_font,
        "title_font": title_names.most_common(1)[0][0] if title_names else (major or body_font),
        "title_pt": median(titles, 36),
        "body_pt": median(bodies, 18),
        "title_bold": bold_titles * 2 >= len(titles) if titles else True,
    }


def _text_height(lines: list[str], width_points: float, font_size: float,
                 family: str = "Arial", bold: bool = False, italic: bool = False) -> float:
    return text_height(lines, width_points, font_size, family, bold, italic)


# Template typography each generated region keeps, collected per render and
# checked by the Font Inspector against the saved PPTX.
_FONT_EXPECTATIONS: ContextVar[dict | None] = ContextVar("font_expectations", default=None)


def _role_font(shape, role: str) -> str | None:
    """Theme typeface a heading or body text of this slide inherits."""
    try:
        major, minor = theme_fonts(shape.part.slide.slide_layout.slide_master)
    except AttributeError:
        return None
    return major if role == "title" else minor


def _clear_text_preserving_style(
    shape, lines: list[str], *, role: str = "body",
    fallback_font: str | None = None, fallback_size: int | float | None = None,
) -> None:
    """Replace the text; the typeface and size stay exactly the template's.

    Text in a template object is drawn in the face and size that object resolves to
    (run, list styles, layout and master placeholders, theme). A region the renderer
    adds uses the template's own typography (fallback_*). The size never shrinks to fit:
    long text is paginated, shortened by the repair loop or reported by the audit.
    """
    text_frame = shape.text_frame
    source_paragraph = next(
        (paragraph for paragraph in text_frame.paragraphs if paragraph.runs),
        text_frame.paragraphs[0] if text_frame.paragraphs else None,
    )
    source_ppr = deepcopy(source_paragraph._p.pPr) if source_paragraph is not None and source_paragraph._p.pPr is not None else None
    source_run = source_paragraph.runs[0] if source_paragraph is not None and source_paragraph.runs else None
    source_rpr = deepcopy(source_run._r.rPr) if source_run is not None and source_run._r.rPr is not None else None
    source_alignment = source_paragraph.alignment if source_paragraph is not None else None
    if source_run is not None or shape.is_placeholder:
        family, size, _ = effective_font(shape, source_paragraph, source_run)
    else:
        family = fallback_font or _role_font(shape, role)
        size = float(fallback_size or (36 if role == "title" else 18))
    # Inherited card padding can consume a resized region entirely.
    horizontal_margin = min(int(shape.width * 0.04), int(Pt(6)))
    vertical_margin = min(int(shape.height * 0.05), int(Pt(4)))
    text_frame.margin_left = min(text_frame.margin_left, horizontal_margin)
    text_frame.margin_right = min(text_frame.margin_right, horizontal_margin)
    text_frame.margin_top = min(text_frame.margin_top, vertical_margin)
    text_frame.margin_bottom = min(text_frame.margin_bottom, vertical_margin)
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.NONE
    text_frame.vertical_anchor = MSO_ANCHOR.TOP
    for index, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        if source_ppr is not None:
            if paragraph._p.pPr is not None:
                paragraph._p.remove(paragraph._p.pPr)
            paragraph._p.insert(0, deepcopy(source_ppr))
        if source_alignment is not None:
            paragraph.alignment = source_alignment
        # Inherited hanging indents change wrapping even after replacing text.
        ppr = paragraph._p.get_or_add_pPr()
        ppr.set("marL", "0")
        ppr.set("marR", "0")
        ppr.set("indent", "0")
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        paragraph.line_spacing = 1.15
        run = paragraph.add_run()
        if source_rpr is not None:
            run._r.insert(0, deepcopy(source_rpr))
        run.text = line
        if family:
            run.font.name = family
        run.font.size = Pt(size)
    expectations = _FONT_EXPECTATIONS.get()
    if expectations is not None and any(line.strip() for line in lines):
        expectations[shape._element] = (shape, family, float(size))


def _clear_all_text(slide) -> list:
    text_shapes = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            _clear_text_preserving_style(shape, [])
            text_shapes.append(shape)
    return text_shapes


def _title_and_body(text_shapes: list, width: int, height: int):
    candidates = [shape for shape in text_shapes if shape.width > width * 0.15 and shape.height > height * 0.025]
    if not candidates:
        return None, None
    # A small eyebrow above the heading is not the title. Prefer the actual
    # typographic hierarchy, then position among similarly prominent regions.
    def size(shape):
        return max((r.font.size.pt for p in shape.text_frame.paragraphs for r in p.runs
                    if r.text.strip() and r.font.size), default=18)
    upper = [shape for shape in candidates if shape.top < height * 0.60] or candidates
    largest = max(map(size, upper), default=0)
    headings = [shape for shape in upper if size(shape) >= largest * 0.85] if largest else upper
    title = min(headings, key=lambda shape: shape.top / max(height, 1) - shape.width / max(width, 1) * 0.06)
    remaining = [shape for shape in candidates if shape is not title]
    body = max(remaining, key=lambda shape: shape.width * shape.height, default=None)
    return title, body


def _add_text_region(slide, x: float, y: float, w: float, h: float):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    shape.text_frame.word_wrap = True
    return shape


def _remove_shape(shape) -> None:
    element = shape.element
    if element.getparent() is not None:
        element.getparent().remove(element)


def _card_regions(slide, title, width: int, height: int) -> list:
    """Find rectangular infographic panels that can each carry one point."""
    def rectangular(shape) -> bool:
        geometry = shape.element.find(".//" + qn("a:prstGeom"))
        return geometry is not None and geometry.get("prst") in {"rect", "roundRect"}

    area = width * height
    candidates = [
        shape for shape in slide.shapes
        if shape.shape_id != title.shape_id
        and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
        and rectangular(shape)
        and shape.left >= 0 and shape.top >= 0
        and shape.left + shape.width <= width and shape.top + shape.height <= height
        and shape.has_text_frame
        and shape.top >= title.top + title.height // 2
        and shape.width >= width * 0.17
        and shape.height >= height * 0.10
        and 0.025 <= shape.width * shape.height / area <= 0.30
    ]
    groups = [
        [other for other in candidates
         if 0.70 <= other.width / shape.width <= 1.30
         and 0.70 <= other.height / shape.height <= 1.30]
        for shape in candidates
    ]
    cards = max(groups, key=len, default=[])
    distinct = []
    for shape in sorted(cards, key=lambda item: (item.top, item.left)):
        overlap = any(
            max(0, min(shape.left + shape.width, other.left + other.width) - max(shape.left, other.left))
            * max(0, min(shape.top + shape.height, other.top + other.height) - max(shape.top, other.top))
            > min(shape.width * shape.height, other.width * other.height) * 0.5
            for other in distinct
        )
        if not overlap:
            distinct.append(shape)
    return distinct if len(distinct) >= 2 or (
        len(distinct) == 1 and distinct[0].width >= width * 0.25
        and distinct[0].height >= height * 0.20
    ) else []


def _card_text_color(card) -> RGBColor:
    try:
        if card.fill.type == MSO_FILL.SOLID and card.fill.fore_color.type == MSO_COLOR_TYPE.RGB:
            return RGBColor(255, 255, 255) if _luminance(card.fill.fore_color.rgb) < 0.23 else RGBColor(20, 25, 35)
    except (AttributeError, TypeError, ValueError):
        pass
    return RGBColor(20, 25, 35)


def _fill_cards(slide, cards: list, bullets: list[str], title, variant_number: int,
                typography: dict[str, str | int], protected: set[int]) -> None:
    used = cards[:min(len(cards), len(bullets))]
    for card in cards[len(used):]:
        # A card owns the markers and labels enclosed by its rectangle.
        contained = [shape for shape in list(slide.shapes)
                     if shape.shape_id not in {title.shape_id, card.shape_id, *protected}
                     and shape.left >= card.left and shape.top >= card.top
                     and shape.left + shape.width <= card.left + card.width
                     and shape.top + shape.height <= card.top + card.height]
        for shape in contained:
            if shape.element.getparent() is not None:
                _remove_shape(shape)
        _remove_shape(card)
    for index, card in enumerate(used):
        # The card's inherited labels, pills and icons describe the old slide.
        # Remove them before placing the new point on the panel itself.
        for shape in list(slide.shapes):
            if shape.shape_id in {card.shape_id, title.shape_id, *protected}:
                continue
            if (shape.left >= card.left and shape.top >= card.top
                    and shape.left + shape.width <= card.left + card.width
                    and shape.top + shape.height <= card.top + card.height):
                _remove_shape(shape)
        _clear_text_preserving_style(card, [])
        inset_x = min(int(card.width * (0.10 if variant_number == 2 and index % 2 == 0 else 0.06)),
                      int(Inches(0.42 if variant_number == 2 else 0.28)))
        inset_top = min(int(card.height * (0.26 if variant_number == 2 else 0.18)),
                        int(Inches(0.38 if variant_number == 2 else 0.25)))
        region = slide.shapes.add_textbox(
            card.left + inset_x, card.top + inset_top,
            card.width - 2 * inset_x,
            card.height - inset_top - min(int(card.height * 0.08), int(Inches(0.10))),
        )
        region.name = "Generated infographic text"
        lines = bullets[index:] if index == len(used) - 1 else [bullets[index]]
        _clear_text_preserving_style(
            region, lines, fallback_font=typography["font"],
            fallback_size=int(typography["body_pt"]),
        )
        for paragraph in region.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = _card_text_color(card)



def _theme_color(slide, name: str, fallback: str) -> RGBColor:
    from lxml import etree
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    theme = slide.slide_layout.slide_master.part.part_related_by(RT.THEME)
    scheme = etree.fromstring(theme.blob).find(".//" + qn("a:clrScheme"))
    item = scheme.find(qn("a:" + name)) if scheme is not None else None
    value = item[0] if item is not None and len(item) else None
    return RGBColor.from_string((value.get("lastClr") or value.get("val")) if value is not None else fallback)


def _clear_region(slide, title, region: tuple[int, int, int, int], protected: set[int], width: int, height: int) -> None:
    """Old text and the objects in a diagram's area leave; the heading, brand and backdrop stay."""
    for shape in list(slide.shapes):
        if shape.shape_id == title.shape_id or shape.shape_id in protected:
            continue
        if not shape.has_text_frame and shape.width * shape.height > width * height * 0.55:
            continue
        if shape.has_text_frame or (
            max(0, min(shape.left + shape.width, region[0] + region[2]) - max(shape.left, region[0]))
            * max(0, min(shape.top + shape.height, region[1] + region[3]) - max(shape.top, region[1]))
            > shape.width * shape.height * 0.20
        ):
            _remove_shape(shape)


def _add_chart_infographic(slide, spec, bullets: list[str], region: tuple[int, int, int, int], gap: int,
                           typography: dict, palette: dict[str, RGBColor] | None) -> None:
    """An editable column (or bar) chart of the points' figures, with the points as its key takeaways."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
    left, top, region_width, region_height = region
    chart_width = int((region_width - gap) * 0.58)
    data = CategoryChartData()
    data.categories = spec.categories
    data.add_series(spec.unit or "Значение", spec.values)
    long_labels = max(len(label) for label in spec.categories) > 14
    kind = XL_CHART_TYPE.BAR_CLUSTERED if long_labels else XL_CHART_TYPE.COLUMN_CLUSTERED
    graphic = slide.shapes.add_chart(kind, left, top, chart_width, region_height, data)
    graphic.name = "Infographic node chart"
    chart = graphic.chart
    accent = palette["accent"] if palette is not None else _theme_color(slide, "accent1", "4472C4")
    chart.has_legend = False
    chart.font.size = Pt(11)
    if typography.get("font"):
        chart.font.name = typography["font"]
    series = chart.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = accent
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.font.size, labels.font.bold = Pt(12), True
    labels.number_format, labels.number_format_is_linked = ('0.##"%"' if spec.unit == "%" else "0.##"), False
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.tick_labels.font.size = Pt(10)
    chart.category_axis.tick_labels.font.size = Pt(11)
    # The value axis names its unit, as the design audit expects of every chart.
    chart.value_axis.has_title = True
    chart.value_axis.axis_title.text_frame.text = (spec.unit or "Значение")[:60]
    for paragraph in chart.value_axis.axis_title.text_frame.paragraphs:
        paragraph.font.size, paragraph.font.bold = Pt(10), False
    caption = slide.shapes.add_textbox(left + chart_width + gap, top, region_width - chart_width - gap, region_height)
    caption.name = "Infographic node caption"
    caption.text_frame.word_wrap = True
    _clear_text_preserving_style(caption, bullets, fallback_font=typography["font"],
                                 fallback_size=int(typography["body_pt"]))
    caption.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    for paragraph in caption.text_frame.paragraphs[1:]:
        paragraph.space_before = Pt(10)


def _render_infographic(
    slide, title, bullets: list[str], width: int, height: int,
    layout: str, typography: dict[str, str | int],
    protected: set[int], palette: dict[str, RGBColor] | None,
) -> bool:
    """Build a factual, editable diagram from the already approved bullet text.

    modules: a tile per point; flow: tiles with arrows; stats: a tile per point under its big
    figure; chart: a native chart of the figures the points carry, the points beside it.
    """
    from .visuals import chart_series, stat_parts
    stats = stat_parts(bullets) if layout == "stats" else None
    spec = chart_series(bullets) if layout == "chart" else None
    if (layout not in {"modules", "flow", "stats", "chart"} or (layout == "stats" and stats is None)
            or (layout == "chart" and spec is None)
            or not (1 if layout == "chart" else 2) <= len(bullets) <= (8 if layout == "chart" else 6)):
        return False
    top = max(title.top + title.height + Inches(0.22), Inches(1.9))
    bottom = height - Inches(0.75)
    left = Inches(0.85)
    gap = Inches(0.24)
    if layout == "chart":
        if bottom - top < Inches(2.4):
            return False
        region = (left, top, width - 2 * left, bottom - top)
        _clear_region(slide, title, region, protected, width, height)
        _add_chart_infographic(slide, spec, bullets, region, gap, typography, palette)
        return True
    columns = len(bullets) if layout == "flow" else (3 if len(bullets) in {3, 5, 6} else 2)
    rows = math.ceil(len(bullets) / columns)
    node_width = (width - 2 * left - gap * (columns - 1)) // columns
    node_height = (bottom - top - gap * (rows - 1)) // rows
    if node_height < Inches(0.75) or node_width < Inches(1):
        return False
    figure_pt = min(44.0, float(typography["body_pt"]) * 2.4)
    # A node is as tall as its text needs (with air), not the whole free height; the block
    # of nodes sits in the middle of the free area instead of stretching empty boxes.
    need = max(_text_height([bullet], max(20, node_width / 12700 - 24), float(typography["body_pt"]),
                            typography["font"]) for bullet in bullets) * 12700 + Inches(0.6)
    if stats is not None:
        need += int(figure_pt * 1.25 * 12700)
    node_height = int(min(node_height, max(need, Inches(1.2))))
    top += (bottom - top - node_height * rows - gap * (rows - 1)) // 2
    region = (left, top, width - 2 * left, bottom - top)
    _clear_region(slide, title, region, protected, width, height)
    surface = palette["surface"] if palette is not None else _theme_color(slide, "lt2", "FFFFFF")
    accent = palette["accent"] if palette is not None else _theme_color(slide, "accent1", "000000")
    ink = palette["text"] if palette is not None else _theme_color(slide, "dk1", "000000")
    if _contrast(ink, surface) < 4.5:
        ink = RGBColor(0x14, 0x19, 0x23) if _luminance(surface) > 0.4 else RGBColor(0xFF, 0xFF, 0xFF)
    for index, bullet in enumerate(bullets):
        x = left + (index % columns) * (node_width + gap)
        y = top + (index // columns) * (node_height + gap)
        node = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, node_width, node_height)
        node.name = f"Infographic node {index + 1}"
        node.fill.solid()
        node.fill.fore_color.rgb = surface
        node.line.color.rgb = accent
        node.line.width = Pt(1.2)
        _clear_text_preserving_style(node, [bullet], fallback_font=typography["font"],
                                     fallback_size=int(typography["body_pt"]))
        # A tile's text is the theme's text colour: the shape's default (white) vanishes on a light tile.
        for paragraph in node.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = ink
        node.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        if stats is not None:
            # The figure is the tile's headline: large, bold, in the accent colour, in a box of its
            # own above the point, so the point keeps the template's size.
            band = int(figure_pt * 1.35 * 12700)
            node.text_frame.vertical_anchor = MSO_ANCHOR.TOP
            node.text_frame.margin_top = band + int(Inches(0.12))
            figure = slide.shapes.add_textbox(x, y + int(Inches(0.1)), node_width, band)
            figure.name = f"Infographic figure {index + 1}"
            figure.text_frame.word_wrap = True
            paragraph = figure.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.CENTER
            run = paragraph.add_run()
            run.text = stats[index][0]
            run.font.size, run.font.bold, run.font.color.rgb = Pt(figure_pt), True, accent
            if typography.get("font"):
                run.font.name = typography["font"]
            for item in node.text_frame.paragraphs:
                item.alignment = PP_ALIGN.CENTER
        if layout == "flow" and index < len(bullets) - 1:
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + node_width + Inches(.02),
                                          y + node_height // 2 - Inches(.10), gap - Inches(.04), Inches(.20))
            arrow.name = f"Infographic connector {index + 1}"
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = accent
            arrow.line.fill.background()
    return True

def _prune_unassigned_source_text(slide, original_text_shapes: list, title, cards: list, width: int, height: int, protected: set[int]) -> None:
    area = width * height
    protected_ids = {title.shape_id, *(card.shape_id for card in cards), *protected}
    for label in original_text_shapes:
        if label.shape_id in protected_ids or label.element.getparent() is None:
            continue
        for panel in list(slide.shapes):
            if panel.shape_id in protected_ids or panel.shape_id == label.shape_id or (panel.has_text_frame and panel.text.strip()):
                continue
            panel_area = panel.width * panel.height
            contains_active = any(
                panel.left <= active.left and panel.top <= active.top
                and panel.left + panel.width >= active.left + active.width
                and panel.top + panel.height >= active.top + active.height
                for active in [title, *cards]
                if active.element.getparent() is not None
            )
            if (not contains_active and panel_area <= area * 0.35
                    and panel_area >= label.width * label.height * 0.5
                    and panel.left <= label.left and panel.top <= label.top
                    and panel.left + panel.width >= label.left + label.width
                    and panel.top + panel.height >= label.top + label.height):
                _remove_shape(panel)
        _remove_shape(label)


def _fit_body_region(slide, body, title, width: int, height: int, protected: set[int], bullets: list[str], allow_reposition: bool = True) -> None:
    """Find space for editable text without covering retained artwork or a heading."""
    gap = Inches(0.18)
    left, top = body.left, body.top
    right = min(width - Inches(0.35), body.left + body.width)
    bottom = height - Inches(0.35)
    required = _text_height(bullets, max(20, (right - left) / 12700 - 12), 18) * 12700
    if bottom - top < required + Inches(0.2):
        top = Inches(0.7)

    if left < title.left + title.width and right > title.left:
        top = max(top, title.top + title.height + gap)
    inherited = [shape for owner in (slide.slide_layout, slide.slide_layout.slide_master)
                 for shape in owner.shapes if not shape.is_placeholder]
    obstacles = [shape for shape in slide.shapes
                 if shape.shape_id not in {body.shape_id, title.shape_id}
                 and (shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP}
                      or shape.shape_id in protected)
                 and 0 < shape.width * shape.height < width * height * 0.55]
    obstacles += [shape for shape in inherited if 0 < shape.width * shape.height < width * height * 0.55]
    regions = [(left, top, right, bottom)]
    for shape in obstacles:
        ox, oy = shape.left - gap, shape.top - gap
        ex, ey = shape.left + shape.width + gap, shape.top + shape.height + gap
        next_regions = []
        for x, y, r, b in regions:
            if r <= ox or x >= ex or b <= oy or y >= ey:
                next_regions.append((x, y, r, b))
            else:
                next_regions.extend([(x, y, min(r, ox), b), (max(x, ex), y, r, b),
                                     (x, y, r, min(b, oy)), (x, max(y, ey), r, b)])
        regions = [box for box in next_regions if box[2] - box[0] >= Inches(1.5)
                   and box[3] - box[1] >= Inches(0.6)]
    if regions:
        x, y, r, b = max(regions, key=lambda box: (box[2] - box[0]) * (box[3] - box[1]))
        body.left, body.top, body.width, body.height = x, y, r - x, b - y
        needed = _text_height(bullets, max(20, (r - x) * 0.88 / 12700 - 12), 16) * 12700 + Inches(0.35)
        if allow_reposition and needed > b - y and title.top > height * 0.30:
            # Cover/footer captions cannot hold body paragraphs. Reclaim the free
            # upper part of the same column while preserving all brand artwork.
            title.top = Inches(1.45)
            body.top = title.top + title.height + gap
            _fit_body_region(slide, body, title, width, height, protected, bullets, False)


def _numbering(shape) -> bool:
    return shape.has_text_frame and bool(_NUMBERING.fullmatch(shape.text.strip()))


def _point_decoration(slide, point, keep: set[int], slide_area: int) -> list:
    """The number, icon or marker above a point region, in the same column."""
    left, right = point.left, point.left + point.width
    found = []
    for shape in slide.shapes:
        if shape.shape_id in keep or shape.shape_id == point.shape_id or shape.width * shape.height > slide_area * 0.03:
            continue
        if shape.has_text_frame and shape.text.strip() and not _numbering(shape):
            continue
        center = shape.left + shape.width // 2
        near = (point.top - (shape.top + shape.height) < point.height * 2.5
                and shape.top < point.top + point.height)
        if left <= center <= right and near:
            found.append(shape)
    return found


def _remove_point_decoration(slide, point, keep: set[int], slide_area: int) -> None:
    """Drop the number, icon or marker that belongs to an unused point region."""
    for shape in _point_decoration(slide, point, keep, slide_area):
        _remove_shape(shape)


def _fill_points(slide, title, points: list, bullets: list[str], typography: dict,
                 protected: set[int], width: int, height: int) -> list:
    """One point per parallel region of the template, at the region's own typeface and size."""
    count = min(len(points), len(bullets))
    per = math.ceil(len(bullets) / count)
    groups = [group for group in (bullets[index * per:(index + 1) * per] for index in range(count)) if group]
    single_row = all(abs(shape.top - points[0].top) <= points[0].height * 0.5 for shape in points)
    if single_row and 1 < len(groups) < len(points):
        # A row keeps its ends: two points of three take the outer regions, the row stays symmetric.
        chosen = sorted({round(index * (len(points) - 1) / (len(groups) - 1)) for index in range(len(groups))})
        used = [points[index] for index in chosen]
    else:
        used = points[:len(groups)]
    unused = [shape for shape in points if all(shape is not item for item in used)]
    # An unused region leaves with its number or icon, even one the analysis marked as brand:
    # it belongs to that point, not to the slide's frame.
    keep = {title.shape_id, *(shape.shape_id for shape in used)}
    for shape in unused:
        _remove_point_decoration(slide, shape, keep, width * height)
        _remove_shape(shape)
    bottom = height - Inches(0.4)
    for shape, lines in zip(used, groups):
        _clear_text_preserving_style(shape, lines, fallback_font=typography["font"],
                                     fallback_size=typography["body_pt"])
        frame = shape.text_frame
        run = frame.paragraphs[0].runs[0]
        inner = max(20, (shape.width - frame.margin_left - frame.margin_right) / 12700)
        needed = (int(_text_height(lines, inner, run.font.size.pt, run.font.name or typography["font"]) * 12700)
                  + frame.margin_top + frame.margin_bottom + int(Pt(4)))
        # A longer point grows down to the next row or the bottom margin, never shrinks.
        below = [other.top for other in used if other.top >= shape.top + shape.height // 2
                 and other.left < shape.left + shape.width and shape.left < other.left + other.width]
        limit = min(below + [bottom]) - int(Inches(0.1))
        if needed > shape.height:
            left, top, box_width = shape.left, shape.top, shape.width
            shape.left, shape.top, shape.width, shape.height = left, top, box_width, max(shape.height, min(needed, limit - top))
    return used


def _fill_slide(slide, content: SlideContent, width: int, height: int, variant_number: int = 0,
                protected: set[int] | None = None, typography: dict[str, str | int] | None = None,
                infographic_layout: str | None = None,
                dark_palette: dict[str, RGBColor] | None = None, points: list[int] | None = None,
                heading: int | None = None, strict_heading: bool = False, heading_grow: int = 0) -> bool:
    typography = typography or {"font": None, "title_font": None, "title_pt": 36, "body_pt": 18}
    text_shapes = [shape for shape in slide.shapes if shape.has_text_frame and shape.shape_id not in (protected or set())]
    original_text_shapes = [shape for shape in text_shapes if shape.text.strip()]
    content_shapes = [shape for shape in text_shapes if shape.text.strip() or shape.is_placeholder]
    title, body = _title_and_body(content_shapes, width, height)
    point_bullets = [item for item in content.bullets if item.strip()]
    if points or strict_heading:
        # A point layout (and every slide of the strict template mode) takes its heading from
        # the template analysis: a centred heading (points around a circle) or a narrow
        # auto-size one ("Кейс") is smaller than a guess by type size would expect.
        title = next((shape for shape in text_shapes if shape.shape_id == heading), title)
        if body is title:
            body = max((shape for shape in content_shapes if shape is not title
                        and shape.width > width * 0.15 and shape.height > height * 0.025),
                       key=lambda shape: shape.width * shape.height, default=None)
    if title is not None and heading_grow > title.width and title.text_frame.word_wrap is False:
        # A heading that does not wrap takes the room its line grows into, and wraps there.
        left, top = title.left, title.top
        title.left, title.top, title.width, title.height = left, top, min(heading_grow, width - left), title.height
        title.text_frame.word_wrap = True
    point_shapes = [shape for shape_id in (points or []) for shape in text_shapes
                    if shape.shape_id == shape_id and shape is not title]
    if (len(point_shapes) >= 2 and len(point_bullets) >= 2 and not infographic_layout and title is not None
            and not _card_regions(slide, title, width, height)):
        # The layout has a region per point (numbered steps, icon points, columns): each point
        # takes its own region, numbering stays, unused regions leave with their decoration.
        kept = {shape.shape_id for shape in point_shapes}
        for shape in text_shapes:
            if shape is not title and shape.shape_id not in kept and not _numbering(shape):
                _clear_text_preserving_style(shape, [])
        # All four values at once: a layout placeholder may have no own transform.
        left, top = min(max(0, title.left), width - 91440), min(max(0, title.top), height - 91440)
        title.left, title.top, title.width, title.height = (left, top, max(91440, min(title.width, width - left)),
                                                            max(91440, min(title.height, height - top)))
        _clear_text_preserving_style(title, [content.title], role="title", fallback_font=typography.get("title_font"),
                                     fallback_size=int(typography["title_pt"]))
        run = title.text_frame.paragraphs[0].runs[0]
        needed = int(_text_height([content.title], max(20, title.width / 12700 - 12), run.font.size.pt,
                                  run.font.name or typography.get("title_font")) * 12700 + Inches(0.12))
        title.height = min(height - title.top - Inches(0.5), max(title.height, needed))
        used = _fill_points(slide, title, point_shapes, point_bullets, typography, protected or set(), width, height)
        numbering = [shape for shape in slide.shapes if _numbering(shape)]
        _prune_unassigned_source_text(slide, original_text_shapes, title, [*used, *numbering], width, height,
                                      protected or set())
        return False
    for shape in text_shapes:
        if shape is not title and shape is not body:
            _clear_text_preserving_style(shape, [])
    width_in = width / 914400
    height_in = height / 914400
    if title is None:
        title = _add_text_region(slide, 0.7, 0.45, max(3.0, width_in - 1.4), 1.0)
    cards = [shape for shape in _card_regions(slide, title, width, height)
             if shape.shape_id not in (protected or set())]
    if body is not None and body.height < Inches(1.4):
        # Expand short template captions instead of shrinking a whole paragraph to 11 pt.
        available = height - body.top - Inches(0.35)
        if available >= Inches(1.4):
            body.height = min(available, Inches(2.6))
    if body is None and (not cards or not content.bullets):
        top = max(Inches(1.75), title.top + title.height + Inches(0.2))
        body = _add_text_region(
            slide, 0.9, top / 914400, max(3.0, width_in - 1.8),
            max(1.4, (height - top - Inches(0.4)) / 914400),
        )
    # Some legitimate template text boxes intentionally bleed past the canvas.
    # Generated text must remain entirely on-slide for editable exports.
    for shape in (title, body) if body is not None else (title,):
        shape.left = min(max(0, shape.left), width - 91440)
        shape.top = min(max(0, shape.top), height - 91440)
        shape.width = max(91440, min(shape.width, width - shape.left))
        shape.height = max(91440, min(shape.height, height - shape.top))
    _clear_text_preserving_style(
        title, [content.title], role="title", fallback_font=typography.get("title_font"),
        fallback_size=int(typography["title_pt"]),
    )
    # Increase title frame height when wrapping requires more than the source caption.
    title_size = max((run.font.size.pt for paragraph in title.text_frame.paragraphs
                      for run in paragraph.runs if run.font.size), default=28)
    needed = int(_text_height([content.title], max(20, title.width / 12700 - 12), title_size) * 12700 + Inches(0.12))
    title.height = min(height - title.top - Inches(0.5), max(title.height, needed))
    bullets = [item for item in content.bullets if item.strip()]
    if infographic_layout and _render_infographic(
        slide, title, bullets, width, height, infographic_layout, typography,
        protected or set(), dark_palette,
    ):
        return True
    if cards and bullets:
        _fill_cards(slide, cards, bullets, title, variant_number, typography, protected or set())
        _prune_unassigned_source_text(slide, original_text_shapes, title, cards, width, height, protected or set())
        return False
    _fit_body_region(slide, body, title, width, height, protected or set(), bullets)
    _clear_text_preserving_style(
        body, bullets or [" "], fallback_font=typography["font"],
        fallback_size=int(typography["body_pt"]),
    )
    # Layout variants preserve template colors and the exact approved content.
    if bullets and variant_number in {1, 2} and dark_palette is None:
        x, y, w, h = body.left, body.top, body.width, body.height
        gap = min(int(w * 0.04), 228600)
        groups = [[bullet] for bullet in bullets]
        box_w = int(w * 0.88) if variant_number == 2 else w
        first = next((run for paragraph in body.text_frame.paragraphs for run in paragraph.runs), None)
        size = first.font.size.pt if first is not None and first.font.size else float(typography["body_pt"])
        family = first.font.name if first is not None and first.font.name else typography["font"]
        weights = [_text_height(group, max(20, box_w / 12700 - 12), size, family) for group in groups]
        available = max(1, h - gap * (len(groups) - 1))
        boxes = []
        cursor = y
        for i, weight in enumerate(weights):
            row_h = int(available * weight / sum(weights))
            boxes.append((x + (int(w * 0.12) if variant_number == 2 and i % 2 == 0 else 0), cursor, box_w, row_h))
            cursor += row_h + gap
        # The points share the body's height at the template's size; when they
        # cannot, they stay together in one region instead of shrinking.
        if not all(weight * 12700 + Pt(8) <= box[3] for weight, box in zip(weights, boxes)):
            boxes = []
        for number, (group, box) in enumerate(zip(groups, boxes)):
            if not group:
                continue
            if number == 0:
                region = body
            else:
                element = deepcopy(body.element)
                node = element.find(".//" + qn("p:cNvPr"))
                node.set("id", str(slide.shapes._next_shape_id))
                node.set("name", f"Content {number + 1}")
                slide.shapes._spTree.insert_element_before(element, "p:extLst")
                region = slide.shapes[-1]
            region.left, region.top, region.width, region.height = box
            _clear_text_preserving_style(
                region, group, fallback_font=typography["font"],
                fallback_size=int(typography["body_pt"]),
            )
    _prune_unassigned_source_text(slide, original_text_shapes, title, [body], width, height, protected or set())
    if title is body:
        raise ValueError("Title and body reference the same shape")
    return False


def _prune_orphan_decoration(slide, protected: set[int]) -> None:
    active = [shape for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
    for shape in list(slide.shapes):
        if shape.shape_id in protected or (shape.has_text_frame and shape.text.strip()):
            continue
        if shape.shape_type not in {MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.TEXT_BOX}:
            continue
        geom = shape.element.find(".//" + qn("a:prstGeom"))
        # Large background accents and decorative freeform artwork survive.
        if geom is not None and geom.get("prst") not in {"rect", "roundRect", "line"}:
            continue
        # Retain only panels that actually contain an active text object.
        if any(shape.left <= text.left and shape.top <= text.top
               and shape.left+shape.width >= text.left+text.width
               and shape.top+shape.height >= text.top+text.height for text in active):
            continue
        if shape.height > Inches(6):
            continue
        _remove_shape(shape)


def _render_dense_text(slide, content: SlideContent, width: int, height: int, typography: dict) -> None:
    """Give dense copy an unobstructed surface; never paint text over artwork."""
    from .audit import _estimated_overflow
    family, title_family = typography["font"], typography.get("title_font")
    title_pt, body_pt = float(typography["title_pt"]), float(typography["body_pt"])
    left, top, body_w, title_h, body_top, body_h = dense_geometry(width, height, content.title, family,
                                                                 title_pt, title_family)
    # Dense pages deliberately replace decorative objects with a clean theme surface.
    for shape in list(slide.shapes):
        _remove_shape(shape)
    slide._element.set("showMasterSp", "0")
    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, width, height)
    panel.name = "Dense content surface"
    panel.fill.solid()
    panel.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
    panel.line.fill.background()
    title = slide.shapes.add_textbox(Pt(left), Pt(top), Pt(body_w), Pt(title_h))
    title.name = "Dense content title"
    _clear_text_preserving_style(title, [content.title], role="title", fallback_font=title_family,
                                 fallback_size=title_pt)
    for run in title.text_frame.paragraphs[0].runs:
        run.font.bold = bool(typography.get("title_bold", True))
        run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    body = slide.shapes.add_textbox(Pt(left), Pt(body_top), Pt(body_w), Pt(body_h))
    body.name = "Dense content body"
    size = body_pt
    # Pagination sizes pages at the template's body size. Text split for another size steps
    # down to the reading minimum instead of failing; the Font Inspector still expects the
    # template's size and reports the page.
    while size > BODY_PT and not body_fits(content.bullets, body_w, body_h, family, size):
        size -= 1
    _clear_text_preserving_style(body, content.bullets, fallback_font=family, fallback_size=size)
    expectations = _FONT_EXPECTATIONS.get()
    if expectations is not None and body._element in expectations:
        shape, written_family, _ = expectations[body._element]
        expectations[body._element] = (shape, written_family, body_pt)
    for paragraph in body.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    if _estimated_overflow(title) or _estimated_overflow(body):
        raise ValueError("Плотный слайд требует разбиения на продолжения перед рендерингом.")


def _dense_page_fits(content: SlideContent, width: int, height: int, typography: dict) -> bool:
    """Whether a clean page holds the slide at the smallest size a dense page may use."""
    family = typography["font"]
    _, _, body_w, _, _, body_h = dense_geometry(width, height, content.title, family,
                                                float(typography["title_pt"]), typography.get("title_font"))
    size = min(float(typography["body_pt"]), BODY_PT)
    return body_h > 20 and body_fits(content.bullets, body_w, body_h, family, size)


def _repair_text_layout(slide, content: SlideContent, width: int, height: int,
                        protected: set[int], typography: dict, force: bool = False) -> bool:
    """Reflow layouts that cannot hold their text at a readable size."""
    from .audit import _estimated_overflow
    active = [shape for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
    editable = [shape for shape in active if shape.shape_id not in protected]
    title = next((shape for shape in editable if shape.text == content.title), None)
    if title is None:
        return False
    def intersects(a, b):
        return (min(a.left+a.width, b.left+b.width)-max(a.left,b.left) > Pt(2)
                and min(a.top+a.height, b.top+b.height)-max(a.top,b.top) > Pt(2))
    broken = any(_estimated_overflow(shape) for shape in editable)
    broken |= any(intersects(a,b) for i,a in enumerate(active) for b in active[i+1:]
                  if a.shape_id not in protected or b.shape_id not in protected)
    # force: the exported PDF showed letters of this slide overlapping or leaving a frame.
    if not broken and not force:
        return False
    region = _add_text_region(slide, .7, (title.top+title.height+Inches(.2))/914400,
                              width/914400-1.4, max(.1, (height-title.top-title.height-Inches(.65))/914400))
    _fit_body_region(slide, region, title, width, height, protected, content.bullets)
    for shape in list(slide.shapes):
        if shape.shape_id in {title.shape_id, region.shape_id, *protected}:
            continue
        if shape.shape_type in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP}:
            continue
        if shape.width * shape.height >= width * height * .55:
            continue
        _remove_shape(shape)
    region.name = "Reflowed content"
    _clear_text_preserving_style(region, content.bullets, fallback_font=typography["font"],
                                fallback_size=typography["body_pt"])
    # A clean page is used only when it can hold the text; otherwise the audit reports
    # the overflow and the repair loop shortens the text.
    if _estimated_overflow(region) and _dense_page_fits(content, width, height, typography):
        _render_dense_text(slide, content, width, height, typography)
        protected.clear()
    return True


def _prepare_native_area(slide, content: SlideContent, width: int, height: int,
                         typography: dict | None = None) -> tuple[int, int, int, int]:
    """Use an opaque, template-sized surface so inherited artwork cannot hide data."""
    for shape in list(slide.shapes):
        element = shape.element
        element.getparent().remove(element)
    panel = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, width, height)
    panel.fill.solid()
    panel.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_1
    panel.line.fill.background()
    heading = _add_text_region(slide, 0.55, 0.29, width / 914400 - 1.1, 0.73)
    _clear_text_preserving_style(heading, [content.title], role="title",
                                 fallback_font=(typography or {}).get("title_font"),
                                 fallback_size=(typography or {}).get("title_pt"))
    for paragraph in heading.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.color.theme_color = MSO_THEME_COLOR.TEXT_1
    accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.55), Inches(1.08), Inches(0.62), Inches(0.055)
    )
    accent.fill.solid()
    accent.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
    accent.line.fill.background()
    margin = Inches(0.55)
    top = Inches(1.25)
    bottom = Inches(0.4)
    return (margin, top, width - 2 * margin, height - top - bottom)


def _normalize_presentation_xml(presentation) -> None:
    # ECMA-376 requires notesMasterIdLst before sldIdLst. A few uploaded decks
    # place it after the slide list; python-pptx otherwise preserves that order.
    root = presentation._element
    notes = root.find(qn("p:notesMasterIdLst"))
    slides = root.find(qn("p:sldIdLst"))
    if notes is not None and slides is not None and root.index(notes) > root.index(slides):
        root.remove(notes)
        root.insert(root.index(slides), notes)


def _remove_original_slides(presentation, original_count: int) -> None:
    slide_ids = presentation.slides._sldIdLst
    for _ in range(original_count):
        slide_id = slide_ids[0]
        presentation.part.drop_rel(slide_id.rId)
        slide_ids.remove(slide_id)


LAYOUT_ARCHETYPES = {"title", "content", "comparison", "data", "closing"}
MAX_LAYOUT_SHAPES = 60
# A heading sample such as "Заголовок — ключевая мысль слайда" marks a designed layout; a heading
# about fonts, icons, screenshots, tables or charts marks a page of the template's style guide.
_LAYOUT_LABEL = re.compile(r"заголов|тезис|описани|title|heading|headline", re.IGNORECASE)
_GUIDE_LABEL = re.compile(r"пример|оформлени|шрифт|иконк|скриншот|диаграмм|гистограмм|график|таблиц|нумерац|"
                          r"таймлайн|гант|example|sample|guideline|font|icon|screenshot|chart|table", re.IGNORECASE)
_NUMBERING = re.compile(r"0?\d{1,2}\.?")


def _slot_role(composition: Composition, slot) -> str:
    return composition.object_roles.get(str(slot.shape_id), "replaceable")


def _slot_pt(slot, default: float = 14.0) -> float:
    """A slot's type size in points (the analysis stores it in EMU)."""
    size = slot.font_size or 0
    return size / 12700 if size > 1000 else float(size or default)


def _slot_width(slot) -> int:
    """The width a region's text may take: a frame of words that does not wrap grows to its neighbour."""
    if sum(character.isalpha() for character in slot.text) < 2 and slot.text.strip():
        return slot.width
    return max(slot.width, getattr(slot, "grow_width", 0) or 0)


def _heading_slot(composition: Composition, width: int, height: int):
    """The layout's heading: the largest wide text region in the upper part of the slide.

    A slot without its own size inherits the placeholder's heading style, so it counts as large.
    A narrow auto-size heading ("Кейс") is wide by the room its line grows into.
    """
    wide = [slot for slot in composition.slots if _slot_width(slot) >= width * 0.25 and slot.y < height * 0.45
            and _slot_role(composition, slot) == "replaceable"]
    return max(wide, key=lambda slot: (_slot_pt(slot, 36), -slot.y), default=None)


def _point_slots(composition: Composition, width: int, height: int) -> list:
    """Parallel text regions of a layout (numbered steps, icon points, columns) in reading order."""
    heading = _heading_slot(composition, width, height)
    candidates = [slot for slot in composition.slots
                  if slot is not heading and _slot_role(composition, slot) == "replaceable"
                  and sum(character.isalpha() for character in slot.text) >= 3
                  and slot.width * slot.height >= width * height * 0.004
                  # Under a top heading an eyebrow above it is not a point; a centred heading
                  # (points around a circle) has points above it too.
                  and (heading is None or heading.y > height * 0.25 or slot.y >= heading.y + heading.height * 0.5)]
    best: list = []
    for slot in candidates:
        similar = [other for other in candidates
                   if 0.7 <= other.width / max(1, slot.width) <= 1.3 and 0.7 <= other.height / max(1, slot.height) <= 1.3]
        distinct: list = []
        for other in sorted(similar, key=lambda item: (item.y, item.x)):
            if not any(other.x < kept.x + kept.width and kept.x < other.x + other.width
                       and other.y < kept.y + kept.height and kept.y < other.y + other.height for kept in distinct):
                distinct.append(other)
        if len(distinct) > len(best):
            best = distinct
    if len(best) < 2:
        return []
    rows: list[list] = []
    for slot in sorted(best, key=lambda item: item.y):
        row = next((row for row in rows if abs(row[0].y - slot.y) <= slot.height * 0.5), None)
        if row is None:
            rows.append([slot])
        else:
            row.append(slot)
    return [slot for row in rows for slot in sorted(row, key=lambda item: item.x)]


# --- Strict template mode ---------------------------------------------------------------------
# Every text object of the chosen template slide stays where it is and keeps its own paragraph
# and run properties; only its text changes, and never beyond the lines the template's own text
# took there. A list (one text object, several paragraphs) is a place per paragraph.

@dataclass
class Region:
    """One place for one text: a whole text object, one paragraph of a list or one marked item."""
    slot: object
    part: int | None
    max_chars: int
    parts: int = 1
    # (top, height) in EMU of an item level with its marker; None for a paragraph or a whole object.
    band: tuple[int, int] | None = None

    @property
    def shape_id(self) -> int:
        return self.slot.shape_id

    @property
    def x(self) -> int:
        return self.slot.x

    @property
    def width(self) -> int:
        return self.slot.width

    @property
    def height(self) -> int:
        return self.band[1] if self.band else self.slot.height // self.parts

    @property
    def y(self) -> int:
        return self.band[0] if self.band else self.slot.y + (self.part or 0) * self.height

    @property
    def paragraphs(self) -> int:
        return 1 if self.part is not None else self.slot.paragraphs

    @property
    def kind(self) -> str:
        return "marked_item" if self.band else "list_line" if self.part is not None else "text"


def _reading_order(slots: list) -> list:
    rows: list[list] = []
    for slot in sorted(slots, key=lambda item: item.y):
        row = next((row for row in rows if abs(row[0].y - slot.y) <= max(slot.height, row[0].height) * 0.5), None)
        if row is None:
            rows.append([slot])
        else:
            row.append(slot)
    return [slot for row in rows for slot in sorted(row, key=lambda item: item.x)]


def _regions_of(slot) -> list[Region]:
    items = list(getattr(slot, "list_items", None) or [])
    if len(items) > 1:
        # A column of checkmarks, dots or numbers beside one text box: a place per marker.
        return [Region(slot, number, item["max_chars"], len(items), (item["top"], item["height"]))
                for number, item in enumerate(items)]
    limits = list(getattr(slot, "paragraph_chars", None) or [])
    if len(limits) > 1:
        return [Region(slot, number, limit, len(limits)) for number, limit in enumerate(limits)]
    return [Region(slot, None, slot.max_chars if slot.max_chars > 0 else 450)]


def strict_regions(composition: Composition, width: int, height: int):
    """(heading slot, places in reading order) generated text fills in strict mode.

    Numbering ("1", "02") stays as the template's; brand text, objects to remove and fields the
    Field Marker found mostly covered are not places. Without a heading the first place takes
    the title.
    """
    heading = _heading_slot(composition, width, height)
    slots = [slot for slot in composition.slots
             if slot is not heading and _slot_role(composition, slot) == "replaceable" and slot.clear_share >= 0.5
             and not _NUMBERING.fullmatch(slot.text.strip())
             and (sum(character.isalpha() for character in slot.text) >= 3 or not slot.text.strip())]
    regions = [region for slot in _reading_order(slots) for region in _regions_of(slot)]
    if heading is None and regions:
        heading, regions = regions[0].slot, [region for region in regions if region.slot is not regions[0].slot]
    return heading, regions


def _limit(slot, default: int) -> int:
    return slot.max_chars if slot is not None and slot.max_chars > 0 else default


def slot_budget(compositions: list[Composition], width: int, height: int) -> dict:
    """Characters a slide's title and each place may take in all the given layouts.

    kinds tells the writer what each place is ("marked_item": a short point beside a checkmark,
    icon or number; "list_line": a line of a list; "text": a text block); min_texts is how many
    places must be filled so that no marker of a list is left without its point.
    """
    regions = [strict_regions(item, width, height) for item in compositions]
    count = min(len(texts) for _, texts in regions)
    kinds = [regions[0][1][index].kind for index in range(count)]
    marked = [index for index, kind in enumerate(kinds) if kind == "marked_item"]
    return {
        "title": min(_limit(heading, 110) for heading, _ in regions),
        "texts": [min(texts[index].max_chars for _, texts in regions) for index in range(count)],
        "paragraphs": [min(texts[index].paragraphs for _, texts in regions) for index in range(count)],
        "kinds": kinds,
        "min_texts": min(6, marked[-1] + 1) if marked else 0,
    }


def _text_length(text: str) -> int:
    return len("\n".join(" ".join(line.split()) for line in text.split("\n") if line.strip()))


# Words that must not end a cut text: prepositions, conjunctions and particles.
_DANGLING = {"и", "а", "но", "или", "в", "во", "на", "по", "с", "со", "к", "ко", "о", "об", "от", "до", "за",
             "из", "у", "для", "при", "про", "без", "над", "под", "что", "как", "не", "the", "a", "an", "of",
             "to", "in", "on", "for", "and", "or", "with"}


def fit_chars(text: str, limit: int) -> str:
    """The text within limit characters (whitespace collapsed), cut only between whole words.

    A cut prefers a sentence end; it never splits a word and drops a trailing preposition or
    conjunction, so the rest reads as a finished phrase.
    """
    text = "\n".join(" ".join(line.split()) for line in text.split("\n") if line.strip())
    if limit <= 0 or len(text) <= limit:
        return text
    head = text[:limit]
    ends = [index + 1 for index, character in enumerate(head) if character in ".!?;" and index + 1 >= limit * 0.5]
    if ends:
        return head[:ends[-1]].strip()
    # Whole words only: a word the limit falls into goes entirely.
    cut = head if text[limit] in " \n" else head[:max(head.rfind(" "), head.rfind("\n"), 0)]
    words = cut.replace("\n", " \n").split(" ")
    # A figure never stays without the unit after it ("на 15" of "на 15 млн руб."): it goes too,
    # and so does a preposition left dangling before it.
    while len(words) > 1 and (words[-1].strip(",;:—–-\n").casefold() in _DANGLING
                              or any(character.isdigit() for character in words[-1])):
        words.pop()
    result = " ".join(words).replace(" \n", "\n").rstrip(" ,;:—–-\n")
    # A single word longer than the limit is kept whole rather than broken.
    return result or text.split()[0]


def fit_slide(content: SlideContent, budget: dict) -> SlideContent:
    """A slide cut to its budget: title and one text per place, none longer than the template's."""
    from dataclasses import replace
    texts = [item for item in content.bullets if item.strip()]
    limits = budget.get("texts", [])
    return replace(content, title=fit_chars(content.title, budget.get("title", 0)),
                   bullets=[fit_chars(text, limit) for text, limit in zip(texts, limits)] if limits else texts)


def _replace_text_exact(shape, text: str) -> None:
    """New text in the object's own paragraph and run properties: nothing else changes.

    Line i of the text takes the style of the template's paragraph i (a list keeps its markers).
    """
    frame = shape.text_frame
    paragraphs = list(frame.paragraphs)
    filled = [paragraph for paragraph in paragraphs if paragraph.runs and paragraph.text.strip()] \
        or [paragraph for paragraph in paragraphs if paragraph.runs] or paragraphs[:1]
    first = filled[0] if filled else None
    first_run = next((run for run in first.runs if run.text.strip()), first.runs[0] if first is not None and first.runs else None) \
        if first is not None else None
    family, size, _ = effective_font(shape, first, first_run)
    styles = []
    for paragraph in filled:
        run = next((item for item in paragraph.runs if item.text.strip()), paragraph.runs[0] if paragraph.runs else None)
        styles.append((deepcopy(paragraph._p.pPr) if paragraph._p.pPr is not None else None,
                       deepcopy(run._r.rPr) if run is not None and run._r.rPr is not None else None))
    body = frame._txBody
    for paragraph in list(body.findall(qn("a:p"))):
        body.remove(paragraph)
    lines = [line for line in text.split("\n") if line.strip()] or [""]
    for index, line in enumerate(lines):
        ppr, rpr = styles[min(index, len(styles) - 1)] if styles else (None, None)
        paragraph = body.add_p()
        if ppr is not None:
            paragraph.insert(0, deepcopy(ppr))
        run = paragraph.add_r()
        if rpr is not None:
            run.insert(0, deepcopy(rpr))
        run.text = line
    expectations = _FONT_EXPECTATIONS.get()
    if expectations is not None and text.strip():
        expectations[shape._element] = (shape, family, float(size))


def _keep_in_clear_area(shape, slot) -> None:
    """Text stays inside the field's free area: only the frame's inner margins change, not the frame."""
    if slot is None or not slot.clear_box or slot.clear_share >= 0.97:
        return
    x, y, w, h = slot.clear_box
    frame = shape.text_frame
    left, top = x - shape.left, y - shape.top
    right, bottom = shape.left + shape.width - (x + w), shape.top + shape.height - (y + h)
    frame.margin_left = max(frame.margin_left or 0, left)
    frame.margin_top = max(frame.margin_top or 0, top)
    frame.margin_right = max(frame.margin_right or 0, right)
    frame.margin_bottom = max(frame.margin_bottom or 0, bottom)


def _fit_in_field(shape, text: str) -> str:
    """Shorten the written text by words until it fits its field at the template's size.

    Words wrap differently from the template's sample, so equal length does not always fit;
    at most 40 % of the text goes, the rest is left to the Field Checker and the repair loop.
    """
    from .audit import _estimated_overflow
    floor = len(text) * 0.6
    while text and len(text) > floor and _estimated_overflow(shape):
        shorter = fit_chars(text, max(1, len(text) - max(2, len(text) // 12)))
        if shorter == text or len(shorter) < floor:
            break
        text = shorter
        _replace_text_exact(shape, text)
    return text


def _paragraph_bands(shape) -> list[tuple[int, int]]:
    """Vertical bands (EMU) of the template paragraphs of a text object, before its text changes."""
    from .fonts import text_block_height
    frame = shape.text_frame
    paragraphs = [paragraph for paragraph in frame.paragraphs if paragraph.text.strip()]
    heights = [max(1.0, text_block_height(shape, [paragraph.text])) for paragraph in paragraphs]
    inner = shape.height - (frame.margin_top or 0) - (frame.margin_bottom or 0)
    block = min(inner, int(sum(heights) * 12700))
    anchor = frame.vertical_anchor
    offset = (inner - block) // 2 if anchor == MSO_ANCHOR.MIDDLE else inner - block if anchor == MSO_ANCHOR.BOTTOM else 0
    cursor, bands = shape.top + (frame.margin_top or 0) + offset, []
    for value in heights:
        band = int(block * value / sum(heights))
        bands.append((cursor, cursor + band))
        cursor += band
    return bands


def _remove_line_decoration(slide, shape, bands: list[tuple[int, int]], keep: set[int], slide_area: int) -> None:
    """Markers and icons beside the lines of removed list paragraphs leave with them."""
    left, right = shape.left - int(Inches(1.2)), shape.left + shape.width
    for item in list(slide.shapes):
        if item.shape_id in keep or item.shape_id == shape.shape_id or item.width * item.height > slide_area * 0.03:
            continue
        if item.has_text_frame and item.text.strip() and not _numbering(item):
            continue
        center_x, center_y = item.left + item.width // 2, item.top + item.height // 2
        if left <= center_x <= right and any(top <= center_y <= bottom for top, bottom in bands):
            _remove_shape(item)


def _grow_heading(shape, slot, text: str, slide_width: int) -> None:
    """A heading frame that does not wrap takes the room its line grows into, and wraps there.

    Without this a narrow auto-size heading ("Кейс") either runs off the slide or holds only a
    word cut to its sample's length.
    """
    grow = getattr(slot, "grow_width", 0) or 0
    if grow <= shape.width or len(text) <= len(slot.text):
        return
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if slot.align == "center":
        left = left + width // 2 - grow // 2
    elif slot.align == "right":
        left = left + width - grow
    # All four values at once: a placeholder may have no transform of its own.
    shape.left, shape.top, shape.width, shape.height = max(0, left), top, min(grow, slide_width - max(0, left)), height
    shape.text_frame.word_wrap = True


def _clone_text_shape(slide, shape, number: int):
    """A copy of a text object with its own id, right above it in the drawing order."""
    element = deepcopy(shape.element)
    node = element.find(".//" + qn("p:cNvPr"))
    node.set("id", str(slide.shapes._next_shape_id))
    node.set("name", f"{shape.name} — item {number + 1}")
    shape.element.addnext(element)
    return next(item for item in slide.shapes if item.element is element)


def _fill_marked_items(slide, shape, items: list) -> list:
    """Each point of a marked list in its own copy of the text object, level with its marker.

    The copies keep the object's paragraph and run properties, left edge and width; only their
    top and height follow the markers, so a point that wraps never pushes the next one away from
    its checkmark.
    """
    boxes = [shape] + [_clone_text_shape(slide, shape, number) for number in range(1, len(items))]
    written = []
    for box, (region, text) in zip(boxes, items):
        top, band = region.band
        left, width = box.left, box.width
        box.left, box.top, box.width, box.height = left, top, width, band
        box.text_frame.vertical_anchor = MSO_ANCHOR.TOP
        box.text_frame.word_wrap = True
        _replace_text_exact(box, text)
        written.append((box, _fit_in_field(box, text)))
    return written


def _remove_markers(slide, slot, parts: set[int]) -> None:
    """The checkmarks, dots or numbers of a marked list's unused items leave with them."""
    ids = {item["marker_id"] for number, item in enumerate(getattr(slot, "list_items", None) or []) if number in parts}
    for shape in list(slide.shapes):
        if shape.shape_id in ids:
            _remove_shape(shape)


def _fill_strict(slide, composition: Composition, content: SlideContent, width: int, height: int) -> SlideContent:
    """Strict template mode for one slide; returns the text as written (cut to the template)."""
    from dataclasses import replace
    heading, regions = strict_regions(composition, width, height)
    shapes = {shape.shape_id: shape for shape in slide.shapes}
    written: set[int] = set()
    title = fit_chars(content.title, _limit(heading, 110))
    if heading is not None and heading.shape_id in shapes:
        _replace_text_exact(shapes[heading.shape_id], title)
        _grow_heading(shapes[heading.shape_id], heading, title, width)
        _keep_in_clear_area(shapes[heading.shape_id], heading)
        title = _fit_in_field(shapes[heading.shape_id], title)
        written.add(heading.shape_id)
    texts = [item for item in content.bullets if item.strip()][:len(regions)]
    used = regions[:len(texts)]
    row = [region for region in regions if region.part is None
           and abs(region.y - regions[0].y) <= regions[0].height * 0.5] if regions else []
    if 1 < len(texts) < len(row) and all(region.part is None for region in used):
        # The first row keeps its ends: two points of three take the outer places, the row stays
        # symmetric; a caption under the row (a footnote) stays unused.
        used = [row[round(index * (len(row) - 1) / (len(texts) - 1))] for index in range(len(texts))]
    texts = [fit_chars(text, region.max_chars) for region, text in zip(used, texts)]
    # Paragraph places of one list are written together, in order, each in its own style.
    by_shape: dict[int, list[tuple[Region, str]]] = {}
    for region, text in zip(used, texts):
        by_shape.setdefault(region.shape_id, []).append((region, text))
    listed = {region.shape_id: [item for item in regions if item.shape_id == region.shape_id] for region in regions}
    lines: list[str] = []
    for shape_id, items in by_shape.items():
        shape = shapes.get(shape_id)
        if shape is None:
            continue
        parts = listed[shape_id]
        if items[0][0].band is not None:
            # A marked list: a point per checkmark, the unused checkmarks leave.
            for box, text in _fill_marked_items(slide, shape, items):
                written.add(box.shape_id)
                lines.append(text)
            _remove_markers(slide, items[0][0].slot, set(range(len(items), len(parts))))
            continue
        if len(parts) > 1 and len(items) < len(parts):
            bands = _paragraph_bands(shape)
            keep = set(written) | set(by_shape) | ({heading.shape_id} if heading is not None else set())
            _remove_line_decoration(slide, shape, bands[len(items):], keep, width * height)
        joined = "\n".join(text for _, text in items)
        _replace_text_exact(shape, joined)
        _keep_in_clear_area(shape, items[0][0].slot)
        lines += [line for line in _fit_in_field(shape, joined).split("\n") if line.strip()]
        written.add(shape_id)
    # A place without text leaves with its number, icon or checkmarks; the rest of the slide stays.
    for shape_id in {region.shape_id for region in regions} - set(by_shape):
        shape = shapes.get(shape_id)
        if shape is not None and shape.element.getparent() is not None:
            _remove_point_decoration(slide, shape, written, width * height)
            _remove_markers(slide, listed[shape_id][0].slot, set(range(len(listed[shape_id]))))
            _remove_shape(shape)
    kept_roles = {"fixed_brand", "reusable_asset"}
    for shape in list(slide.shapes):
        # Old sample text outside the places (a caption, a stray label) must not survive.
        if (shape.has_text_frame and shape.text.strip() and shape.shape_id not in written and not _numbering(shape)
                and composition.object_roles.get(str(shape.shape_id)) not in kept_roles):
            _replace_text_exact(shape, "")
    return replace(content, title=title, bullets=lines)


# --- Variant goals: base, more text, more visuals ------------------------------------------------
GOALS = ("base", "more_text", "more_visual")


def _text_capacity(composition: Composition, width: int, height: int) -> int:
    _, regions = strict_regions(composition, width, height)
    return sum(region.max_chars for region in regions)


def _visual_weight(composition: Composition, width: int, height: int) -> float:
    """Icons, numbers, markers and shapes the layout keeps besides its text, and its parallel points.

    A logo repeats on every slide of a template, so brand objects do not make a layout visual.
    """
    roles = composition.object_roles
    decor = sum(1 for role in roles.values() if role == "reusable_asset")
    decor += max(0, composition.shape_count - len(composition.slots) - len(roles))
    markers = 3 if any(slot.list_items for slot in composition.slots) else 0
    return decor + markers + (4 if len(_point_slots(composition, width, height)) >= 2 else 0)


def _layout_score(composition: Composition, width: int, height: int, *, content: bool) -> float | None:
    """How well a template slide carries generated text; None marks a slide not to reuse.

    Real templates mix layouts with service pages: empty backdrops, style-guide pages,
    screenshot frames, tables and charts whose data would be removed.
    """
    if composition.archetype not in LAYOUT_ARCHETYPES:
        return None
    roles = list(composition.object_roles.values())
    if roles and sum(role in {"remove", "unresolved"} for role in roles) * 2 > len(roles):
        return None
    if composition.shape_count > MAX_LAYOUT_SHAPES or composition.chart_count or composition.table_count:
        return None
    if composition.picture_area_ratio >= (0.3 if content else 0.5):
        return None
    heading = _heading_slot(composition, width, height)
    if heading is None:
        return None
    # Figures other than a numbering (43 %, 2010) are old data that new text would not replace.
    if any(len(slot.text.strip()) <= 12 and sum(character.isdigit() for character in slot.text) >= 2
           and not _NUMBERING.fullmatch(slot.text.strip()) for slot in composition.slots):
        return None
    if content:
        body = [slot for slot in composition.slots if slot is not heading and slot.width * slot.height >= width * height * 0.05]
        if not (_point_slots(composition, width, height) or body or heading.y + heading.height < height * 0.5):
            return None
    # A designed example with sample text beats a backdrop with empty placeholders, whose
    # artwork is not described by the slots; empty ones serve templates without examples.
    sample = 0.0 if heading.text.strip() else -1.0
    return (sample + (2.0 if _LAYOUT_LABEL.search(heading.text) else 0.0)
            - (3.0 if _GUIDE_LABEL.search(heading.text) else 0.0))


def _slot_capacity(slot, size: float) -> float:
    """Characters a region holds at the given size (about 0.6 em a character, 1.25 lines)."""
    per_line = int((slot.width / 12700) / (size * 0.6))
    lines = int((slot.height / 12700) / (size * 1.25))
    return max(0, per_line) * max(1, lines)


def _points_penalty(composition: Composition, regions: list, content: SlideContent | None,
                    width: int, height: int) -> int | None:
    """0 for a region per point, 1 when part of a grid stays empty, 2 for one plain text region."""
    if content is None:
        return 0 if not regions else 1
    heading = _heading_slot(composition, width, height)
    # A heading of a fixed small region (the centre of a circle) must hold the slide's title
    # at the template's size; a wide heading grows down freely.
    if (heading is not None and heading.width < width * 0.5
            and len(content.title) > _slot_capacity(heading, _slot_pt(heading, 36)) * 0.9):
        return None
    texts = [item for item in content.bullets if item.strip()]
    count = len(regions)
    if count < 2:
        return 2
    # A grid of several rows may leave part of its last row empty; a single row loses one region at most.
    rows = len({round(slot.y / (height * 0.05)) for slot in regions})
    spare = max(1, count // 4) if rows > 1 else 1
    if len(texts) < 2 or len(texts) > count * (2 if count == 2 else 1) or count - len(texts) > spare:
        return None
    per = math.ceil(len(texts) / count)
    for index, slot in enumerate(regions[:math.ceil(len(texts) / per)]):
        # A point may grow down to the next row or the bottom margin, at the template's size.
        text = " ".join(texts[index * per:(index + 1) * per])
        size = _slot_pt(slot)
        below = [other.y for other in regions if other.y >= slot.y + slot.height / 2
                 and other.x < slot.x + slot.width and slot.x < other.x + other.width]
        room = min(below + [height - Inches(0.4)]) - slot.y
        capacity = (slot.width / 12700) / (size * 0.52) * (room / 12700) / (size * 1.25)
        if len(text) > capacity * 0.85:
            return None
    return 0 if count == len(texts) else 1


def _closing_fits(composition: Composition, content: SlideContent, width: int, height: int) -> bool:
    heading = _heading_slot(composition, width, height)
    if heading is None or len(content.title) > _slot_capacity(heading, _slot_pt(heading, 36)) * 1.5:
        return False
    text = " ".join(item for item in content.bullets if item.strip())
    others = [slot for slot in composition.slots
              if slot is not heading and _slot_role(composition, slot) == "replaceable"
              and sum(character.isalpha() for character in slot.text) >= 3]
    if others:
        room = sum(_slot_capacity(slot, _slot_pt(slot)) for slot in others)
    else:
        # Without a caption slot the text goes to the free area under the heading.
        below = SimpleNamespace(width=heading.width,
                                height=max(0, height - heading.y - heading.height - Inches(0.4)))
        room = _slot_capacity(below, 18)
    return len(text) <= room * 0.9


def _strict_penalty(composition: Composition, content: SlideContent | None, width: int, height: int) -> int | None:
    """Strict mode: a layout needs text regions; the closer their count to the points, the better."""
    _, regions = strict_regions(composition, width, height)
    if not regions:
        return None
    if content is None:
        return 0
    return min(3, abs(len(regions) - len([item for item in content.bullets if item.strip()])))


def _layouts_by_analysis(source: list[Composition], width: int, height: int, slide_count: int,
                         slides: list[SlideContent] | None, avoid: list[list[Composition]] | None = None,
                         strict: bool = False, goal: str = "base") -> list[Composition] | None:
    scores = {id(item): _layout_score(item, width, height, content=item.archetype not in {"title", "closing"})
              for item in source}
    pool = [item for item in source if item.archetype in {"content", "comparison", "data"}
            and scores[id(item)] is not None]
    if not pool:
        return None
    preferred = [item for item in pool if scores[id(item)] >= 0] or pool
    covers = [item for item in source if item.archetype == "title" and scores[id(item)] is not None]
    closings = [item for item in source if item.archetype == "closing" and scores[id(item)] is not None]
    def avoided(item, position: int) -> bool:
        # Variants B and C take another layout wherever the template offers one.
        return any(position < len(plan) and plan[position] is item for plan in (avoid or []))

    # The best cover and closing come first; another one only when it is as good.
    cover = (max(covers, key=lambda item: (scores[id(item)], not avoided(item, 0), -item.source_slide_index))
             if covers else preferred[0])
    if slide_count == 1:
        return [cover]
    closing = (max(closings, key=lambda item: (scores[id(item)], not avoided(item, slide_count - 1), item.source_slide_index))
               if closings else preferred[-1])
    if slide_count == 2:
        return [cover, closing]
    regions = {id(item): _point_slots(item, width, height) for item in preferred}
    # "more_text" favours the layouts that hold the most characters, "more_visual" those with
    # icons, checkmarks, numbers and parallel points; "base" takes the closest fit. The bonus is
    # relative to the template's best layout, so it works for sparse and rich templates alike.
    if goal == "more_text":
        capacity = {id(item): _text_capacity(item, width, height) for item in preferred}
        best = max(capacity.values(), default=0) or 1
        bonus = {key: 6.0 * value / best for key, value in capacity.items()}
    elif goal == "more_visual":
        weight = {id(item): _visual_weight(item, width, height) for item in preferred}
        best = max(weight.values(), default=0) or 1
        bonus = {key: 5.0 * value / best for key, value in weight.items()}
    else:
        bonus = {id(item): 0.0 for item in preferred}
    # The text theme may repeat its roomiest layout; the others alternate columns, icons and points.
    repeat, reuse, differ = (1.5, 0.5, 1.0) if goal == "more_text" else (3.0, 1.0, 2.0)
    uses: Counter[int] = Counter()
    chosen: list[Composition] = []
    previous = None
    last = slides[slide_count - 1] if slides is not None and len(slides) >= slide_count else None
    # A final slide with a real message (not "Спасибо") that the closing layout cannot hold
    # takes a content layout instead of spilling into an empty fallback page.
    closing_fits = last is None or _closing_fits(closing, last, width, height)
    for position in range(1, slide_count - (1 if closing_fits else 0)):
        content = slides[position] if slides is not None and position < len(slides) else None
        options = []
        for item in preferred:
            penalty = (_strict_penalty(item, content, width, height) if strict
                       else _points_penalty(item, regions[id(item)], content, width, height))
            if penalty is not None:
                # A fitting layout wins, but repeating one layout slide after slide costs more
                # than a near fit: the deck alternates columns, icons and numbered points.
                cost = (penalty * (1 if goal == "more_text" else 2) + (repeat if item is previous else 0)
                        + reuse * uses[id(item)] + (differ if avoided(item, position) else 0) - bonus[id(item)])
                options.append((cost, -scores[id(item)], item.source_slide_index, item))
        if not options:
            options = [(9 + uses[id(item)] + (3 if item is previous else 0), -scores[id(item)],
                        item.source_slide_index, item) for item in preferred]
        item = min(options, key=lambda row: row[:3])[3]
        uses[id(item)] += 1
        previous = item
        chosen.append(item)
    return [cover, *chosen, closing] if closing_fits else [cover, *chosen]


def choose_compositions(template: PreparedTemplate, slide_count: int,
                        slides: list[SlideContent] | None = None, *, avoid: list[list[Composition]] | None = None,
                        strict: bool = False, goal: str = "base") -> list[Composition]:
    """Pick a source layout for every slide, keeping a cover first and a closing last.

    With the Template Analyst's result the choice follows it: service pages, empty backdrops,
    screenshots, tables and charts are not reused, and a content slide takes a layout with a
    region per point when its text fits there. Without an analysis the source slides are
    sampled in their original order.
    """
    source = sorted(template.compositions, key=lambda item: item.source_slide_index)
    if not source or slide_count < 1:
        raise ValueError("No usable composition in template")
    if any(item.analysis_complete for item in source):
        chosen = _layouts_by_analysis(source, template.width, template.height, slide_count, slides, avoid, strict, goal)
        if chosen is not None:
            return chosen
    if slide_count == 1:
        return [source[0]]
    if len(source) == 1:
        return source * slide_count
    if len(source) == 2 or slide_count == 2:
        positions = [round(index * (len(source) - 1) / max(slide_count - 1, 1))
                     for index in range(slide_count)]
    else:
        middle_count = slide_count - 2
        last_middle = len(source) - 2
        if middle_count == 1:
            middle = [round((1 + last_middle) / 2)]
        else:
            middle = [1 + round(index * (last_middle - 1) / (middle_count - 1))
                      for index in range(middle_count)]
        positions = [0, *middle, len(source) - 1]
    return [source[index] for index in positions]

def render_variant(template_path: Path, template: PreparedTemplate, slides: list[SlideContent],
                   variant_number: int, output_path: Path, **options) -> dict:
    """Render one variant; every region's template typography is recorded for the Font Inspector."""
    token = _FONT_EXPECTATIONS.set({})
    try:
        return _render_variant(template_path, template, slides, variant_number, output_path, **options)
    finally:
        _FONT_EXPECTATIONS.reset(token)


def _render_variant(
    template_path: Path,
    template: PreparedTemplate,
    slides: list[SlideContent],
    variant_number: int,
    output_path: Path,
    *,
    native_data: NumericCsvData | None = None,
    native_slide_index: int | None = None,
    dark_palette: dict[str, str] | None = None,
    infographic_layouts: dict[str, str] | None = None,
    relayout: set[int] | None = None,
    strict: bool = False,
    plan: list[int] | None = None,
) -> dict:
    if (native_data is None) != (native_slide_index is None):
        raise ValueError("Native data and slide index must be supplied together")
    if native_slide_index is not None and not 0 <= native_slide_index < len(slides):
        raise ValueError("Native data slide index is outside the deck")
    presentation = Presentation(str(template_path))
    source_slides = list(presentation.slides)
    typography = _template_typography(presentation)
    original_count = len(source_slides)
    # All variants reuse the same source slide for each slide index; C only rearranges content within it.
    source_plan = choose_compositions(template, len(slides), slides)
    colors = _dark_palette(dark_palette) if variant_number == 1 and dark_palette is not None else None
    selected: list[int] = []
    dropped_objects = 0
    contrast_repairs = 0
    contrast_ratios: list[float] = []
    infographic_slides: list[int] = []
    layout_repairs: list[int] = []
    infographic_fallbacks: list[int] = []
    composition_by_source = {c.source_slide_index: c for c in template.compositions}
    slide_area = template.width * template.height
    written_slides = list(slides)
    for index, content in enumerate(slides):
        composition = (composition_by_source[plan[index]] if plan is not None
                       else composition_by_source[content.source_slide_index]
                       if content.source_slide_index is not None else source_plan[index])
        source = source_slides[composition.source_slide_index]
        target = presentation.slides.add_slide(source.slide_layout)
        if source._element.cSld.find(qn("p:bg")) is not None:
            background = deepcopy(source._element.cSld.find(qn("p:bg")))
            # Raster backgrounds may contain immutable source facts. They need a
            # separate classification; do not silently reintroduce them here.
            if background.find(".//" + qn("a:blip")) is None:
                target._element.cSld.insert(0, background)
        for shape in list(target.shapes):
            element = shape.element
            element.getparent().remove(element)
        # The checkmarks, dots and icons of a marked list belong to its layout, like its text box;
        # a slide drawn as a diagram does not keep them.
        planned = bool((infographic_layouts or {}).get(content.slide_id))
        markers = ({item["marker_id"] for slot in composition.slots for item in (slot.list_items or [])}
                   if strict and not planned else set())
        dropped_markers = ({item["marker_id"] for slot in composition.slots for item in (slot.list_items or [])}
                           if strict and planned else set())
        for shape in source.shapes:
            role = composition.object_roles.get(str(shape.shape_id))
            if shape.shape_id in dropped_markers:
                dropped_objects += 1
            elif _should_copy(shape, slide_area, role) or (shape.shape_id in markers and role not in {"remove", "unresolved"}):
                _clone_shape(source, target, shape)
            else:
                dropped_objects += 1
        source_by_id = {str(shape.shape_id): shape for shape in source.shapes}
        protected = {int(key) for key, role in composition.object_roles.items()
                     if role == "fixed_brand" and key in source_by_id}
        has_infographic = False
        if (strict and not (native_data is not None and index == native_slide_index)
                and not (infographic_layouts or {}).get(content.slide_id)):
            written_slides[index] = _fill_strict(target, composition, content, template.width, template.height)
        elif not content.dense_layout:
            heading_slot = _heading_slot(composition, template.width, template.height)
            has_infographic = _fill_slide(
                target, content, template.width, template.height, variant_number, protected, typography,
                (infographic_layouts or {}).get(content.slide_id), colors,
                points=[slot.shape_id for slot in _point_slots(composition, template.width, template.height)],
                heading=getattr(heading_slot, "shape_id", None), strict_heading=strict,
                heading_grow=getattr(heading_slot, "grow_width", 0) or 0,
            )

        has_native_data = native_data is not None and index == native_slide_index
        if has_native_data:
            bounds = _prepare_native_area(target, content, template.width, template.height, typography)
            add_native_data_visuals(
                target, native_data,
                slide_width=template.width,
                slide_height=template.height,
                bounds=bounds,
                variant_number=variant_number,
            )
        if not has_native_data and not strict:
            if content.dense_layout:
                _render_dense_text(target, content, template.width, template.height, typography)
                protected.clear()
                layout_repairs.append(index + 1)
            elif _repair_text_layout(target, content, template.width, template.height, protected, typography,
                                     force=index in (relayout or set())):
                layout_repairs.append(index + 1)
        if has_infographic:
            # A text fallback must not be reported as a successfully drawn diagram.
            destination = (infographic_slides if any(s.name.startswith("Infographic node")
                           for s in target.shapes) else infographic_fallbacks)
            destination.append(index + 1)
        elif (infographic_layouts or {}).get(content.slide_id):
            # The heading keeps the template's size, so a long one can leave no room
            # for the planned diagram: the points stay as text on this slide.
            infographic_fallbacks.append(index + 1)
        if colors is not None:
            _apply_dark_theme(
                target, content, colors, set() if has_native_data else protected,
                template.width, template.height, native_data=has_native_data,
            )
        if not strict:
            _prune_orphan_decoration(target, protected)
        fallback_surface = colors["background"] if colors is not None else RGBColor(255, 255, 255)
        repaired, minimum, satisfied = _ensure_text_contrast(target, fallback_surface, protected)
        contrast_repairs += repaired
        if minimum:
            if not satisfied:
                raise ValueError(f"Контраст текста ниже 3:1 на слайде {index + 1}")
            contrast_ratios.append(minimum)
        # Reflow and inherited source geometry can leave editable text beyond the canvas.
        for shape in target.shapes:
            if strict or shape.shape_id in protected or not shape.has_text_frame or not shape.text.strip():
                continue
            shape.left = max(0, min(shape.left, template.width - 91440))
            shape.top = max(0, min(shape.top, template.height - 91440))
            shape.width = max(91440, min(shape.width, template.width - shape.left))
            shape.height = max(91440, min(shape.height, template.height - shape.top))
        selected.append(composition.source_slide_index + 1)
    _remove_original_slides(presentation, original_count)
    _normalize_presentation_xml(presentation)
    # The deck carries its fonts: PowerPoint and LibreOffice draw the template's
    # typeface on computers where it is not installed.
    font_report = embed_fonts(presentation, used_families(presentation))
    slide_numbers = {slide.part: number for number, slide in enumerate(presentation.slides, 1)}
    font_expectations = [
        {"slide": slide_numbers[shape.part], "shape_id": shape.shape_id, "family": family, "size": size}
        for element, (shape, family, size) in (_FONT_EXPECTATIONS.get() or {}).items()
        if element.getparent() is not None and shape.part in slide_numbers
    ]
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(str(output_path))
    # Re-open to verify OOXML relationships, native text and slide count.
    verified = Presentation(str(output_path))
    if len(verified.slides) != len(slides):
        raise ValueError("Rendered PPTX failed slide-count verification")
    for index, content in enumerate(written_slides):
        actual_text = "\n".join(shape.text for shape in verified.slides[index].shapes if shape.has_text_frame)
        expected_bullets = [] if native_data is not None and index == native_slide_index else content.bullets
        if content.title not in actual_text or any(bullet not in actual_text for bullet in expected_bullets):
            raise ValueError(f"Rendered PPTX lost generated text on slide {index + 1}")
    if native_data is not None:
        data_slide = verified.slides[native_slide_index]
        tables = [shape for shape in data_slide.shapes if shape.has_table]
        charts = [shape for shape in data_slide.shapes if shape.has_chart]
        if len(tables) != 1 or len(charts) != 1:
            raise ValueError("Rendered PPTX lost its native chart or table")
        table = tables[0].table
        for row_index, row in enumerate(native_data.display_rows, 1):
            for column_index, value in enumerate(row):
                if table.cell(row_index, column_index).text != value:
                    raise ValueError("Rendered PPTX changed a native table value")
        chart_series = list(charts[0].chart.series)
        if len(chart_series) != len(native_data.series):
            raise ValueError("Rendered PPTX changed native chart series count")
        for actual_series, expected_series in zip(chart_series, native_data.series):
            values = list(actual_series.values)
            if len(values) != len(expected_series.values) or any(
                not math.isclose(float(actual), float(expected), rel_tol=1e-12, abs_tol=1e-12)
                for actual, expected in zip(values, expected_series.values)
            ):
                raise ValueError("Rendered PPTX changed a native chart value")
        for shape in (tables[0], charts[0]):
            if (
                shape.left < 0 or shape.top < 0
                or shape.left + shape.width > template.width
                or shape.top + shape.height > template.height
            ):
                raise ValueError("Rendered native data lies outside the slide")
    return {
        "source_slides": selected,
        "slide_count": len(slides),
        "continuation_slides": [i + 1 for i, item in enumerate(slides) if item.title.endswith(" (продолжение)")],
        "dense_text_slides": [i + 1 for i, item in enumerate(verified.slides) if any(s.name == "Dense content body" for s in item.shapes)],
        "infographic_fallback_slides": infographic_fallbacks,
        "composition_strategy": ("strict_template" if strict else "dark_theme" if colors is not None
                                 else ("template", "stacked", "reflow")[variant_number]),
        "written_slides": [item.to_dict() for item in written_slides] if strict else None,
        "dark_palette": (
            {key: "#" + str(value) for key, value in colors.items()} if colors is not None else None
        ),
        "excluded_source_objects": dropped_objects,
        "editable_text": True,
        "editable_chart": native_data is not None,
        "editable_table": native_data is not None,
        "native_data_slide": native_slide_index + 1 if native_slide_index is not None else None,
        "contrast_repairs": contrast_repairs,
        "minimum_text_contrast": round(min(contrast_ratios), 2) if contrast_ratios else None,
        "template_font": typography["font"] or "theme",
        "template_title_pt": typography["title_pt"],
        "template_body_pt": typography["body_pt"],
        "infographic_slides": infographic_slides,
        "layout_repairs": layout_repairs,
        "relayout_slides": sorted(index + 1 for index in (relayout or set())),
        "fonts": font_report,
        "font_expectations": font_expectations,
    }

