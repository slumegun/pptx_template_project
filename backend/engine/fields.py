"""Field Marker: the exact free area of every text field, measured on the slide without text.

The template is copied with all text removed and rendered by LibreOffice. Inside each text
field the rendered pixels show what the field really has under and over it: a picture, an
icon, a card edge or a line. The largest rectangle free of them is where generated text may
go; its main colour is the surface the text sits on. The measurement is deterministic and
pixel-exact, so it is done once per template and cached.
"""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .export import export_pdf, export_previews
from .models import PreparedTemplate

MARKER_VERSION = 2  # 2: a frame that does not wrap is measured over the room its line grows into
DPI = 60
GRID = 48            # cells along the longer side of a field
INK_DISTANCE = 60    # RGB distance from the field's surface that counts as a drawn object
INK_SHARE = 0.06     # a cell with this much ink is not free
BLOCKED_SHARE = 0.5  # a field less free than this is not used for text


def _clear_text(shapes) -> None:
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            _clear_text(shape.shapes)
        elif shape.has_text_frame:
            frame = shape.text_frame
            for paragraph in list(frame.paragraphs)[1:]:
                paragraph._p.getparent().remove(paragraph._p)
            paragraph = frame.paragraphs[0]
            for run in list(paragraph.runs):
                run._r.getparent().remove(run._r)
            # A no-break space keeps an empty placeholder from showing its prompt text.
            paragraph.add_run().text = " "


def blank_copy(template_path: Path, target: Path) -> Path:
    """The template with every text removed: shapes, pictures and backgrounds stay."""
    deck = Presentation(str(template_path))
    for slide in deck.slides:
        _clear_text(slide.shapes)
    deck.save(str(target))
    return Path(target)


def largest_free_rectangle(grid: list[list[bool]]) -> tuple[int, int, int, int]:
    """(left, top, width, height) in cells of the largest all-free rectangle (maximal histogram)."""
    best = (0, 0, 0, 0)
    heights = [0] * (len(grid[0]) if grid else 0)
    for row_index, row in enumerate(grid):
        heights = [height + 1 if free else 0 for height, free in zip(heights, row)]
        stack: list[int] = []
        for column in range(len(heights) + 1):
            current = heights[column] if column < len(heights) else 0
            while stack and heights[stack[-1]] >= current:
                height = heights[stack.pop()]
                left = stack[-1] + 1 if stack else 0
                if height * (column - left) > best[2] * best[3]:
                    best = (left, row_index - height + 1, column - left, height)
            stack.append(column)
    return best


def _pixels(image: Image.Image) -> list:
    # Pillow 12 renamed getdata; older versions only have getdata.
    return list(image.get_flattened_data() if hasattr(image, "get_flattened_data") else image.getdata())


def measure_field(image: Image.Image, box: tuple[float, float, float, float],
                  obstacles: list[tuple[float, float, float, float]] = ()) -> dict | None:
    """Free rectangle (pixels), surface colour and free share of one field on a blank render.

    obstacles are pictures of the slide, its layout and master (in pixels): they block the
    field wherever they lie, whatever their colours.
    """
    left, top, right, bottom = (max(0, round(box[0])), max(0, round(box[1])),
                                min(image.width, round(box[2])), min(image.height, round(box[3])))
    if right - left < 4 or bottom - top < 4:
        return None
    crop = image.crop((left, top, right, bottom)).convert("RGB")
    # A small working copy keeps the measurement fast; the grid is coarser anyway.
    factor = min(1.0, 192 / max(crop.width, crop.height))
    small = crop.resize((max(1, round(crop.width * factor)), max(1, round(crop.height * factor))), Image.Resampling.BOX)
    blocked = Image.new("L", small.size, 0)
    for obstacle in obstacles:
        x0, y0 = (obstacle[0] - left) * factor, (obstacle[1] - top) * factor
        x1, y1 = (obstacle[2] - left) * factor, (obstacle[3] - top) * factor
        if x1 > 0 and y1 > 0 and x0 < small.width and y0 < small.height:
            blocked.paste(255, (max(0, round(x0)), max(0, round(y0)), min(small.width, round(x1)), min(small.height, round(y1))))
    pixels, covered = _pixels(small), _pixels(blocked)
    # The surface is the field's dominant colour outside pictures (a card, the slide background).
    counts: dict[tuple, int] = {}
    for pixel, hidden in zip(pixels, covered):
        if not hidden:
            key = tuple(channel // 8 for channel in pixel)
            counts[key] = counts.get(key, 0) + 1
    surface = tuple(channel * 8 + 4 for channel in max(counts, key=counts.get)) if counts else pixels[0]
    ink = Image.new("L", small.size)
    ink.putdata([255 if hidden or sum(abs(a - b) for a, b in zip(pixel, surface)) > INK_DISTANCE else 0
                 for pixel, hidden in zip(pixels, covered)])
    crop = small
    scale = GRID / max(crop.width, crop.height)
    columns, rows = max(1, min(crop.width, round(crop.width * scale))), max(1, min(crop.height, round(crop.height * scale)))
    cells = ink.resize((columns, rows), Image.Resampling.BOX)
    grid = [[cells.getpixel((column, row)) <= 255 * INK_SHARE for column in range(columns)] for row in range(rows)]
    cell_left, cell_top, cell_width, cell_height = largest_free_rectangle(grid)
    step_x, step_y = crop.width / columns / factor, crop.height / rows / factor
    free = (left + cell_left * step_x, top + cell_top * step_y,
            left + (cell_left + cell_width) * step_x, top + (cell_top + cell_height) * step_y)
    share = (cell_width * cell_height) / (columns * rows)
    return {"free": free, "surface": "#%02X%02X%02X" % surface, "share": round(share, 3)}


def _cache_path(template: PreparedTemplate, output_dir: Path) -> Path:
    root = os.getenv("MODEL_TEMPLATE_CACHE_DIR")
    name = f"{template.template_sha256}-v{MARKER_VERSION}.json"
    return (Path(root) / "fields" / name) if root else Path(output_dir) / "fields" / name


def field_box(slot) -> tuple[int, int, int, int]:
    """(x, y, width, height) a field's text may take: a frame that does not wrap grows to the margin."""
    grow = getattr(slot, "grow_width", 0) or 0
    if grow <= slot.width:
        return slot.x, slot.y, slot.width, slot.height
    left = (slot.x + slot.width // 2 - grow // 2 if slot.align == "center"
            else slot.x + slot.width - grow if slot.align == "right" else slot.x)
    return max(0, left), slot.y, grow, slot.height


def mark_fields(template_path: Path, template: PreparedTemplate, output_dir: Path,
                timeout: float = 240) -> dict[str, dict]:
    """Mark every text field of the template; the result is cached by the template's SHA-256."""
    cache = _cache_path(template, output_dir)
    if cache.is_file():
        marks = json.loads(cache.read_text(encoding="utf-8"))
    else:
        marks = {}
        with tempfile.TemporaryDirectory(prefix="fields-") as scratch:
            folder = Path(scratch)
            pdf = export_pdf(blank_copy(template_path, folder / "blank.pptx"), folder, timeout=timeout)
            pages = export_previews(pdf, folder / "pages", dpi=DPI, timeout=timeout)
            deck = Presentation(str(template_path))
            slides = list(deck.slides)
            area = template.width * template.height
            for composition in template.compositions:
                if composition.source_slide_index >= len(pages):
                    continue
                with Image.open(pages[composition.source_slide_index]) as page:
                    image = page.convert("RGB")
                sx, sy = image.width / template.width, image.height / template.height
                slide = slides[composition.source_slide_index]
                # Pictures block a field; a full-slide background picture is the design's canvas.
                pictures = [item for owner in (slide.slide_layout.slide_master, slide.slide_layout, slide)
                            for item in owner.shapes if item.shape_type == MSO_SHAPE_TYPE.PICTURE
                            and not item.is_placeholder and item.width and item.height
                            and item.width * item.height < area * 0.55]
                obstacles = [(item.left * sx, item.top * sy, (item.left + item.width) * sx, (item.top + item.height) * sy)
                             for item in pictures]
                for slot in composition.slots:
                    x, y, w, h = field_box(slot)
                    measured = measure_field(image, (x * sx, y * sy, (x + w) * sx, (y + h) * sy), obstacles)
                    if measured is None:
                        continue
                    free = measured["free"]
                    marks[f"{composition.source_slide_index}:{slot.shape_id}"] = {
                        "clear_box": [round(free[0] / sx), round(free[1] / sy),
                                      round((free[2] - free[0]) / sx), round((free[3] - free[1]) / sy)],
                        "surface": measured["surface"], "clear_share": measured["share"]}
        cache.parent.mkdir(parents=True, exist_ok=True)
        temporary = cache.with_suffix(".tmp")
        temporary.write_text(json.dumps(marks), encoding="utf-8")
        temporary.replace(cache)
    apply_marks(template, marks)
    return marks


def apply_marks(template: PreparedTemplate, marks: dict[str, dict]) -> PreparedTemplate:
    """Put the measured free area on the slots; a partly covered field holds less text."""
    for composition in template.compositions:
        for slot in composition.slots:
            mark = marks.get(f"{composition.source_slide_index}:{slot.shape_id}")
            if mark is None:
                continue
            slot.clear_box = list(mark["clear_box"])
            slot.surface = mark["surface"]
            slot.clear_share = float(mark["clear_share"])
            if slot.clear_share < 0.97 and slot.font_pt:
                # The text must fit the free part at the template's size, not the whole frame.
                width = slot.clear_box[2] / 12700
                height = slot.clear_box[3] / 12700
                capacity = max(0, int(width / (slot.font_pt * 0.55)) * max(1, int(height / (slot.font_pt * 1.2))))
                limited = min(slot.max_chars, int(capacity * 0.9)) if slot.max_chars else int(capacity * 0.9)
                if slot.max_chars and slot.paragraph_chars:
                    ratio = limited / slot.max_chars
                    slot.paragraph_chars = [max(1, int(value * ratio)) for value in slot.paragraph_chars]
                slot.max_chars = limited
    return template

