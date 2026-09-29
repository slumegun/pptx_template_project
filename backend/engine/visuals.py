"""Figures of a slide's approved points, found without a model, for the visual theme.

A point that carries one figure ("Выручка выросла на 45%") can become a big-number tile; three
or more points with figures in one unit, or one point listing "2023 — 12%, 2024 — 18%", can
become an editable chart. Only numbers already written in the approved text are used, so a
chart never shows a value the fact check has not seen.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

MAX_CHART_POINTS = 8
MAX_CATEGORY_CHARS = 32

# A figure with its unit: "45%", "1,2 млн руб.", "3 раза", "120 000 ₽", "$4.5 млрд".
_FIGURE = re.compile(
    r"(?<![\w.,])(?P<prefix>[$€₽])?\s?(?P<number>\d{1,3}(?:[\u00a0 ]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?)"
    r"(?:\s?(?P<unit>%|п\.\s?п\.|процент\w*|млрд|млн|тыс\.?|трлн|руб\.?|₽|\$|€|раз\w*|x|×|ч\b|час\w*|мин\w*|дн\w*|"
    r"сек\w*|человек|сотрудник\w*|клиент\w*|пользовател\w*|компани\w*|шт\.?|кг|т\b|км|м²|гб|тб))?"
    r"(?:\s?(?P<currency>руб\.?|₽|долл\w*|\$|€|евро))?",
    re.IGNORECASE,
)
# "2023 — 12%", "Москва: 40", "Q1 - 15 млн": a label and its figure inside one point; a comma
# inside a number ("1,2 млн") does not separate pairs.
_SEGMENTS = re.compile(r"[;,](?!\d)")
_LABEL_SEPARATOR = re.compile(r"\s*(?:\s[—–-]\s|[—–]|:)\s*")
_JOINERS = {"на", "до", "в", "во", "с", "со", "по", "от", "за", "около", "почти", "более", "менее", "свыше", "—", "–",
            "-", ":", "это", "составил", "составила", "составило", "составили", "достиг", "достигла", "достигло"}


@dataclass(frozen=True)
class Figure:
    number: float
    text: str          # as written: "45%", "1,2 млн руб."
    unit: str          # normalized unit: "%", "млн руб.", "" for a bare count


@dataclass(frozen=True)
class ChartSpec:
    categories: tuple[str, ...]
    values: tuple[float, ...]
    unit: str


def _value(raw: str) -> float | None:
    cleaned = raw.replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _unit(match: re.Match) -> str:
    unit = (match.group("unit") or "").casefold().rstrip(".")
    currency = (match.group("currency") or match.group("prefix") or "").casefold().rstrip(".")
    unit = {"процент": "%", "процента": "%", "процентов": "%", "₽": "руб", "×": "x"}.get(unit, unit)
    if unit.startswith("раз"):
        unit = "раз"
    currency = {"₽": "руб", "$": "долл", "€": "евро"}.get(currency, currency)
    if currency.startswith("долл"):
        currency = "долл"
    return " ".join(part for part in (unit, currency) if part)


def _is_year(match: re.Match, value: float) -> bool:
    return not _unit(match) and value.is_integer() and 1900 <= value <= 2100


def figures(text: str) -> list[Figure]:
    """Figures of a text with their units; years, dates and list numbers are not figures."""
    found = []
    for match in _FIGURE.finditer(text):
        value = _value(match.group("number"))
        if value is None or _is_year(match, value):
            continue
        # "1." of a numbered line and parts of a date ("12.05") are not figures.
        after = text[match.end():match.end() + 1]
        if not _unit(match) and (after == "." and match.start() == 0 or re.fullmatch(r"\d{1,2}[.]\d{2}", match.group("number"))):
            continue
        found.append(Figure(value, " ".join(match.group(0).split()), _unit(match)))
    return found


def stat_parts(bullets: list[str]) -> list[tuple[str, str]] | None:
    """(big figure, point) for every point when each of 2–6 points carries exactly one figure."""
    points = [item for item in bullets if item.strip()]
    if not 2 <= len(points) <= 6:
        return None
    parts = []
    for point in points:
        found = figures(point)
        if len(found) != 1:
            return None
        parts.append((found[0].text, point))
    return parts


def _category(text: str) -> str:
    words = [word for word in re.split(r"\s+", text.strip(" .,;:—–-")) if word]
    while words and words[0].casefold() in _JOINERS:
        words.pop(0)
    while words and words[-1].casefold() in _JOINERS:
        words.pop()
    label = " ".join(words)
    if len(label) > MAX_CATEGORY_CHARS:
        cut = label[:MAX_CATEGORY_CHARS]
        label = cut[:cut.rfind(" ")] if " " in cut else cut
    return label[:1].upper() + label[1:]


def chart_series(bullets: list[str]) -> ChartSpec | None:
    """An editable chart's data from the approved points, or None when they do not form a series.

    Either one point lists label–figure pairs ("2023 — 12%, 2024 — 18%, 2025 — 25%"), or three
    or more points carry one figure each; all figures must share one unit.
    """
    points = [item for item in bullets if item.strip()]
    for point in points:
        pairs = []
        for segment in _SEGMENTS.split(point):
            parts = [part for part in _LABEL_SEPARATOR.split(segment.strip()) if part.strip()]
            if len(parts) < 2:
                continue
            found = figures(parts[-1])
            label = _category(parts[-2])
            if len(found) == 1 and label:
                pairs.append((label, found[0]))
        if 3 <= len(pairs) <= MAX_CHART_POINTS and len({figure.unit for _, figure in pairs}) == 1:
            return ChartSpec(tuple(label for label, _ in pairs), tuple(figure.number for _, figure in pairs),
                             pairs[0][1].unit)
    if not 3 <= len(points) <= MAX_CHART_POINTS:
        return None
    rows = []
    for point in points:
        found = figures(point)
        if len(found) != 1:
            return None
        figure = found[0]
        start = point.find(figure.text.split()[0])
        label = _category((point[:start] + " " + point[start + len(figure.text):]) if start >= 0 else point)
        if not label:
            return None
        rows.append((label, figure))
    if len({figure.unit for _, figure in rows}) != 1 or len({label for label, _ in rows}) != len(rows):
        return None
    return ChartSpec(tuple(label for label, _ in rows), tuple(figure.number for _, figure in rows), rows[0][1].unit)


def visual_options(bullets: list[str]) -> list[str]:
    """Diagram kinds a slide's points allow, best first."""
    options = []
    if chart_series(bullets) is not None:
        options.append("chart")
    if stat_parts(bullets) is not None:
        options.append("stats")
    return options + ["modules", "flow"]
