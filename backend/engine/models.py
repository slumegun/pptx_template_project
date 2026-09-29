"""Small, serializable contracts shared by the engine stages."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


SCHEMA_VERSION = "1.0"


@dataclass
class Fact:
    fact_id: str
    text: str
    source: str
    location: str
    excerpt: str

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class SlideContent:
    slide_id: str
    title: str
    bullets: list[str]
    fact_ids: list[str] = field(default_factory=list)
    notes: str = ""
    source_slide_index: int | None = None
    dense_layout: bool = False
    continuation_of: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class SlideSlot:
    shape_id: int
    x: int
    y: int
    width: int
    height: int
    text: str
    font_size: int | None = None
    # Exact template data of the text object, resolved the way PowerPoint draws it:
    # own run, paragraph and list styles, layout and master placeholders, theme.
    font_family: str | None = None
    font_pt: float | None = None
    bold: bool = False
    italic: bool = False
    color: str | None = None          # "#RRGGBB" after inheritance, None when unknown
    align: str | None = None
    paragraphs: int = 1
    # Characters of the original text (whitespace collapsed): the most generated text may use.
    # A placeholder with a bare label or no sample gets its geometric capacity instead.
    max_chars: int = 0
    chars_source: str = "sample"
    # The same limit for every paragraph of a list (one point per paragraph, in order).
    paragraph_chars: list[int] = field(default_factory=list)
    # Paragraphs carry list bullets of their own (a list whose markers are in the text).
    bulleted: bool = False
    # Field Marker, measured on the template's PDF: vertical bands (EMU) of the paragraphs and the
    # list items a column of markers beside them defines ({"paragraphs", "markers", "max_chars", "top"}).
    paragraph_bands: list[list[int]] = field(default_factory=list)
    # A column of markers (checkmarks, dots, icons, numbers) beside one text object: one item per
    # marker ({"top", "height", "max_chars", "marker_id"}), each text written level with its marker.
    list_items: list[dict] = field(default_factory=list)
    # Width (EMU) a frame that does not wrap its lines may grow to before the slide's margin;
    # 0 when the frame wraps inside its own width.
    grow_width: int = 0
    # Field Marker, measured on the slide rendered without text: the largest rectangle of the
    # field free of pictures, icons and lines (EMU: x, y, width, height), the colour under it
    # and the free share of the frame.
    clear_box: list[int] | None = None
    surface: str | None = None
    clear_share: float = 1.0

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class Composition:
    source_slide_index: int
    layout_index: int
    fingerprint: str
    slots: list[SlideSlot]
    shape_count: int
    picture_area_ratio: float
    chart_count: int
    table_count: int
    score: float
    object_roles: dict[str, str] = field(default_factory=dict)
    archetype: str = "unknown"
    analysis_notes: list[str] = field(default_factory=list)
    analysis_complete: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PreparedTemplate:
    schema_version: str
    template_sha256: str
    source_name: str
    width: int
    height: int
    slide_count: int
    layout_count: int
    compositions: list[Composition]
    warnings: list[str] = field(default_factory=list)
    analysis_mode: str = "structural"
    analysis_model: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "PreparedTemplate":
        if value.get("schema_version") != SCHEMA_VERSION:
            raise ValueError("Unsupported TemplateIR schema version")
        compositions = []
        for item in value["compositions"]:
            slots = [SlideSlot(**slot) for slot in item["slots"]]
            compositions.append(Composition(**{**item, "slots": slots}))
        return cls(**{**value, "compositions": compositions})


@dataclass
class GeneratedVariant:
    variant_id: str
    pptx_path: Path
    pdf_path: Path | None
    html_path: Path | None
    preview_paths: list[Path]
    issues: list[dict[str, Any]]
    metrics: dict[str, Any]


def generation_budget_seconds(slide_count: int) -> float:
    """Generation time for all three variants.

    A usual deck keeps the five-minute SLA (290 s plus publication); a deck longer than
    fifteen slides gets 20 s more per extra slide for planning, critics and export.
    """
    return 290 + 20 * max(0, slide_count - 15)
