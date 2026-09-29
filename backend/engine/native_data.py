"""Validated CSV data as editable PowerPoint charts and tables.

The first column contains category labels; the following one to three columns
contain numeric series. Oversized inputs are rejected rather than truncated.
"""

from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.util import Inches, Pt

MAX_CSV_BYTES = 256 * 1024
MAX_DATA_ROWS = 12
MAX_SERIES = 3
MAX_LABEL_LENGTH = 48
MAX_HEADER_LENGTH = 36
MAX_ABSOLUTE_VALUE = Decimal("1000000000000")
_NUMBER = re.compile(r"[+-]?(?:\d+(?:[.,]\d*)?|[.,]\d+)(?:[eE][+-]?\d+)?\Z")


@dataclass(frozen=True)
class NumericSeries:
    name: str
    values: tuple[Decimal, ...]


@dataclass(frozen=True)
class NumericCsvData:
    category_header: str
    categories: tuple[str, ...]
    series: tuple[NumericSeries, ...]
    display_rows: tuple[tuple[str, ...], ...]


@dataclass(frozen=True)
class NativeDataShapes:
    table: Any
    chart: Any


def _safe_label(value: str, *, maximum: int, location: str) -> str:
    label = value.strip()
    if not label:
        raise ValueError(f"Empty {location}")
    if len(label) > maximum:
        raise ValueError(f"{location} is too long (maximum {maximum} characters)")
    if any(ord(char) < 32 or char == "\x7f" for char in label):
        raise ValueError(f"{location} contains control characters")
    # Chart data includes an editable workbook. Never pass formula-looking text.
    if label[0] in "=+-@":
        raise ValueError(f"{location} starts with a spreadsheet formula character")
    return label


def _number(value: str, *, location: str) -> Decimal:
    raw = value.strip()
    if not _NUMBER.fullmatch(raw):
        raise ValueError(f"{location} must be a plain number")
    try:
        number = Decimal(raw.replace(",", "."))
    except InvalidOperation as exc:
        raise ValueError(f"{location} must be a plain number") from exc
    if not number.is_finite() or abs(number) > MAX_ABSOLUTE_VALUE:
        raise ValueError(f"{location} is outside the supported numeric range")
    if number and float(number) == 0.0:
        raise ValueError(f"{location} is too small for a PowerPoint chart")
    return number


def parse_numeric_csv(path: Path) -> NumericCsvData:
    """Parse UTF-8 CSV with comma, semicolon, or tab delimiters.

    Decimal commas are accepted in numeric fields. Up to 12 data rows can be
    rendered on a single slide; larger inputs must first be summarized.
    """
    path = Path(path)
    if path.suffix.lower() != ".csv" or not path.is_file():
        raise ValueError("Numeric data source must be an existing .csv file")
    if path.stat().st_size > MAX_CSV_BYTES:
        raise ValueError(f"CSV exceeds the {MAX_CSV_BYTES} byte limit")
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as stream:
            sample = stream.read(8192)
            if not sample.strip():
                raise ValueError("CSV is empty")
            first_record = sample.splitlines()[0]
            widths = {}
            for candidate in (",", ";", "\t"):
                try:
                    widths[candidate] = len(next(csv.reader([first_record], delimiter=candidate, strict=True)))
                except csv.Error:
                    widths[candidate] = 0
            highest = max(widths.values())
            matches = [candidate for candidate, width in widths.items() if width == highest]
            if highest < 2 or len(matches) != 1:
                raise ValueError("CSV delimiter could not be determined")
            delimiter = matches[0]
            stream.seek(0)
            rows = list(csv.reader(stream, delimiter=delimiter, strict=True))
    except UnicodeError as exc:
        raise ValueError("CSV must be UTF-8 encoded") from exc
    except csv.Error as exc:
        raise ValueError("CSV could not be parsed") from exc

    if len(rows) < 2:
        raise ValueError("CSV needs a header and at least one data row")
    if len(rows) - 1 > MAX_DATA_ROWS:
        raise ValueError(f"CSV has too many rows for one slide (maximum {MAX_DATA_ROWS})")
    header = rows[0]
    if not 2 <= len(header) <= MAX_SERIES + 1:
        raise ValueError(f"CSV needs one category column and 1-{MAX_SERIES} numeric columns")
    headers = tuple(
        _safe_label(value, maximum=MAX_HEADER_LENGTH, location=f"header column {index}")
        for index, value in enumerate(header, 1)
    )
    if len({value.casefold() for value in headers}) != len(headers):
        raise ValueError("CSV column headers must be distinct")

    categories: list[str] = []
    columns: list[list[Decimal]] = [[] for _ in headers[1:]]
    display_rows: list[tuple[str, ...]] = []
    for row_number, row in enumerate(rows[1:], 2):
        if len(row) != len(headers):
            raise ValueError(f"CSV row {row_number} has {len(row)} columns; expected {len(headers)}")
        category = _safe_label(row[0], maximum=MAX_LABEL_LENGTH, location=f"row {row_number} category")
        categories.append(category)
        numeric_text = tuple(value.strip() for value in row[1:])
        for index, value in enumerate(numeric_text):
            columns[index].append(_number(value, location=f"row {row_number}, {headers[index + 1]}"))
        display_rows.append((category, *numeric_text))
    return NumericCsvData(
        category_header=headers[0],
        categories=tuple(categories),
        series=tuple(NumericSeries(name, tuple(values)) for name, values in zip(headers[1:], columns)),
        display_rows=tuple(display_rows),
    )


def _validate_box(box: tuple[int, int, int, int], slide_width: int, slide_height: int) -> None:
    x, y, width, height = box
    if x < 0 or y < 0 or width <= 0 or height <= 0:
        raise ValueError("Data visual bounds must be positive and start on the slide")
    if x + width > slide_width or y + height > slide_height:
        raise ValueError("Data visual bounds exceed the slide")
    if width < Inches(6.5) or height < Inches(3.8):
        raise ValueError("Data visual area needs at least 6.5 x 3.8 inches")


def _axis_title(axis, text: str, size: int = 9) -> None:
    axis.has_title = True
    axis.axis_title.text_frame.text = text[:60]
    for paragraph in axis.axis_title.text_frame.paragraphs:
        paragraph.font.size = Pt(size)
        paragraph.font.bold = False


def _style_table(table, data: NumericCsvData, width: int, height: int) -> None:
    row_count = len(data.categories) + 1
    column_count = len(data.series) + 1
    first_width = int(width * 0.34)
    table.columns[0].width = first_width
    remaining = width - first_width
    series_width = remaining // (column_count - 1)
    for index in range(1, column_count):
        table.columns[index].width = series_width if index < column_count - 1 else remaining - series_width * (column_count - 2)
    header_height = min(Inches(0.58), max(Inches(0.4), height // row_count))
    table.rows[0].height = header_height
    body_height = (height - header_height) // (row_count - 1)
    for index in range(1, row_count):
        table.rows[index].height = body_height if index < row_count - 1 else height - header_height - body_height * (row_count - 2)

    headers = (data.category_header, *(item.name for item in data.series))
    for row_index, row in enumerate((headers, *data.display_rows)):
        for column_index, value in enumerate(row):
            cell = table.cell(row_index, column_index)
            cell.text = value
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.05)
            cell.margin_right = Inches(0.05)
            cell.margin_top = 0
            cell.margin_bottom = 0
            if row_index == 0:
                cell.fill.solid()
                cell.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1
            elif row_index % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.theme_color = MSO_THEME_COLOR.BACKGROUND_2
            for paragraph in cell.text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.LEFT if column_index == 0 else PP_ALIGN.RIGHT
                paragraph.font.size = Pt(9 if row_count >= 10 else 10)
                paragraph.font.bold = row_index == 0
                paragraph.font.color.theme_color = MSO_THEME_COLOR.BACKGROUND_1 if row_index == 0 else MSO_THEME_COLOR.TEXT_1


def add_native_data_visuals(
    slide,
    data: NumericCsvData,
    *,
    slide_width: int,
    slide_height: int,
    bounds: tuple[int, int, int, int] | None = None,
    variant_number: int = 0,
) -> NativeDataShapes:
    """Add a native chart and table in an empty area of an existing slide.

    All dimensions are EMUs. The default area leaves room for a title. Existing
    shapes are not removed; callers must supply an unobstructed content area.
    """
    if not data.categories or not data.series:
        raise ValueError("Numeric CSV data must contain categories and series")
    if len(data.categories) > MAX_DATA_ROWS or len(data.series) > MAX_SERIES:
        raise ValueError("Numeric CSV data exceeds the one-slide limit")
    if any(len(item.values) != len(data.categories) for item in data.series):
        raise ValueError("Numeric CSV series lengths differ")
    if len(data.display_rows) != len(data.categories):
        raise ValueError("Numeric CSV table rows differ from chart categories")

    if bounds is None:
        margin = Inches(0.55)
        top = Inches(1.25)
        bottom = Inches(0.4)
        bounds = (margin, top, slide_width - 2 * margin, slide_height - top - bottom)
    _validate_box(bounds, slide_width, slide_height)
    x, y, width, height = bounds
    gap = Inches(0.22)
    chart_width = int((width - gap) * 0.57)
    table_width = width - gap - chart_width
    if table_width <= 0:
        raise ValueError("Data visual area is too narrow")

    chart_data = CategoryChartData()
    chart_data.categories = data.categories
    for item in data.series:
        chart_data.add_series(item.name, tuple(float(value) for value in item.values))
    chart_x = x + table_width + gap if variant_number == 1 else x
    table_x = x if variant_number == 1 else x + chart_width + gap
    chart_type = XL_CHART_TYPE.BAR_CLUSTERED if variant_number == 2 else XL_CHART_TYPE.COLUMN_CLUSTERED
    chart_shape = slide.shapes.add_chart(chart_type, chart_x, y, chart_width, height, chart_data)
    chart = chart_shape.chart
    palette = (MSO_THEME_COLOR.ACCENT_1, MSO_THEME_COLOR.ACCENT_2, MSO_THEME_COLOR.ACCENT_3)
    for index, series in enumerate(chart.series):
        series.format.fill.solid()
        series.format.fill.fore_color.theme_color = palette[index]
    chart.has_legend = len(data.series) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(9)
    chart.category_axis.tick_labels.font.size = Pt(9)
    chart.value_axis.tick_labels.font.size = Pt(9)
    chart.value_axis.has_major_gridlines = True
    # Both axes are named; the value axis carries the series name with its unit from the CSV header.
    _axis_title(chart.category_axis, data.category_header)
    _axis_title(chart.value_axis, data.series[0].name if len(data.series) == 1 else "Значение")

    table_shape = slide.shapes.add_table(
        len(data.categories) + 1, len(data.series) + 1,
        table_x, y, table_width, height,
    )
    _style_table(table_shape.table, data, table_width, height)
    return NativeDataShapes(table=table_shape, chart=chart_shape)

