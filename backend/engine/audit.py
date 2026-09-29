"""Deterministic audit with source and geometry evidence."""

from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .models import Fact, SlideContent


def _issue(
    slide_index: int,
    shape_id: int | None,
    rule_id: str,
    severity: str,
    evidence: str,
    repairability: str,
) -> dict[str, Any]:
    raw = f"{slide_index}:{shape_id}:{rule_id}:{evidence}"
    return {
        "issue_id": "issue_" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:12],
        "slide_id": f"slide_{slide_index + 1}",
        "object_id": str(shape_id) if shape_id is not None else None,
        "rule_id": rule_id,
        "severity": severity,
        "evidence": evidence,
        "repairability": repairability,
        "selected": False,
    }


def _estimated_overflow(shape) -> bool:
    if not shape.has_text_frame or not shape.text.strip() or shape.width <= 0 or shape.height <= 0:
        return False
    from .renderer import _text_height
    frame = shape.text_frame
    width = max(1, (shape.width - frame.margin_left - frame.margin_right) / 12700)
    height = max(1, (shape.height - frame.margin_top - frame.margin_bottom) / 12700)
    required = 0.0
    for paragraph in frame.paragraphs:
        sizes = [run.font.size.pt for run in paragraph.runs if run.font.size]
        size = max(sizes, default=paragraph.font.size.pt if paragraph.font.size else 18)
        run = next((run for run in paragraph.runs if run.text.strip()), None)
        family = run.font.name if run is not None and run.font.name else "Arial"
        required += _text_height([paragraph.text], width, size, family,
                                 run.font.bold if run else False, run.font.italic if run else False)
    return required > height + 1


def duplicate_content_issues(slides: list[SlideContent]) -> list[dict[str, Any]]:
    issues = []
    seen_content: dict[str, int] = {}
    for index, spec in enumerate(slides):
        signature = " ".join(" ".join(spec.bullets).casefold().split())
        if signature and signature in seen_content:
            issues.append(_issue(index, None, "duplicate_content", "warning",
                                 f"Тезисы дословно повторяют слайд {seen_content[signature] + 1}.", "manual"))
        elif signature:
            seen_content[signature] = index
    return issues


def audit_pptx(
    pptx_path: Path,
    slides: list[SlideContent],
    facts: list[Fact],
    source_texts: list[str] | None = None,
) -> list[dict[str, Any]]:
    presentation = Presentation(str(pptx_path))
    if len(presentation.slides) != len(slides):
        raise ValueError("Audit input slide count differs from SlideSpec")
    issues: list[dict[str, Any]] = []
    source_texts = [text.strip() for text in (source_texts or []) if len(text.strip()) >= 20]
    generated_text = "\n".join(
        text for item in slides for text in [item.title, *item.bullets]
    ).casefold()
    valid_fact_ids = {fact.fact_id for fact in facts}
    fact_by_id = {fact.fact_id: fact for fact in facts}
    issues.extend(duplicate_content_issues(slides))
    for index, (slide, spec) in enumerate(zip(presentation.slides, slides)):
        if not spec.fact_ids:
            issues.append(_issue(index, None, "missing_source", "warning", "Слайд не связан с реестром фактов.", "manual"))
        for fact_id in spec.fact_ids:
            if fact_id not in valid_fact_ids:
                issues.append(_issue(index, None, "unknown_fact_id", "blocking", fact_id, "manual"))
        allowed_excerpts = " ".join(fact_by_id[fact_id].excerpt for fact_id in spec.fact_ids if fact_id in fact_by_id)
        allowed_numbers = set(re.findall(r"\d+(?:[.,]\d+)?", allowed_excerpts))
        text_shapes = [shape for shape in slide.shapes if shape.has_text_frame and shape.text.strip()]
        for position, first in enumerate(text_shapes):
            for second in text_shapes[position + 1:]:
                dx = min(first.left + first.width, second.left + second.width) - max(first.left, second.left)
                dy = min(first.top + first.height, second.top + second.height) - max(first.top, second.top)
                if dx > 25400 and dy > 25400:
                    issues.append(_issue(index, first.shape_id, "text_text_overlap", "blocking",
                                         f"Текстовые рамки {first.shape_id} и {second.shape_id} пересекаются.", "manual"))
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text.strip():
                continue
            if shape.left < 0 or shape.top < 0 or shape.left + shape.width > presentation.slide_width + 5000 or shape.top + shape.height > presentation.slide_height + 5000:
                issues.append(_issue(index, shape.shape_id, "out_of_bounds", "blocking", "Текстовый объект выходит за пределы слайда.", "automatic"))
            if _estimated_overflow(shape):
                issues.append(_issue(index, shape.shape_id, "possible_text_overflow", "blocking", "Оценка указывает на возможную обрезку текста; проверьте превью.", "automatic"))
            text = shape.text.strip()
            explicit_sizes = [run.font.size.pt for paragraph in shape.text_frame.paragraphs
                              for run in paragraph.runs if run.text.strip() and run.font.size]
            if explicit_sizes and min(explicit_sizes) < 14 and len(text) > 50:
                issues.append(_issue(index, shape.shape_id, "small_body_text", "warning",
                                     f"Длинный текст набран кеглем {min(explicit_sizes):g} pt.", "manual"))
            if re.search(r"(?i)(?:lorem ipsum|\bTODO\b|\bXXX\b|вставьте текст)", text):
                issues.append(_issue(index, shape.shape_id, "placeholder_text", "blocking", text[:180], "manual"))
            artwork = list(slide.shapes) + [item for owner in (slide.slide_layout, slide.slide_layout.slide_master)
                                                   for item in owner.shapes if not item.is_placeholder]
            for picture in artwork:
                if picture.shape_type not in {MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.GROUP}:
                    continue
                # Full-slide artwork is a background; text overlay is intentional there.
                if picture.width * picture.height >= presentation.slide_width * presentation.slide_height * 0.55:
                    continue
                overlap = max(0, min(shape.left + shape.width, picture.left + picture.width) - max(shape.left, picture.left)) * max(0, min(shape.top + shape.height, picture.top + picture.height) - max(shape.top, picture.top))
                if overlap > shape.width * shape.height * 0.05:
                    issues.append(_issue(index, shape.shape_id, "text_image_overlap", "blocking",
                                         f"Текстовая рамка пересекает изображение/группу {picture.shape_id}.", "manual"))
            if text.casefold() not in generated_text and any(text.casefold() == old.casefold() for old in source_texts):
                issues.append(_issue(index, shape.shape_id, "old_template_content", "blocking", text[:180], "automatic"))
            # A step number of the layout ("1", "02.") is numbering, not a claim.
            found_numbers = set() if re.fullmatch(r"0?\d{1,2}\.?", text) else set(re.findall(r"\d+(?:[.,]\d+)?", text))
            unverified = found_numbers - allowed_numbers
            # Slide numbers and enumerated bullets are allowed; other numbers need a source.
            unverified = {value for value in unverified if not (value.isdigit() and int(value) == index + 1)}
            if unverified and spec.fact_ids:
                issues.append(_issue(index, shape.shape_id, "unverified_number", "warning", ", ".join(sorted(unverified)), "manual"))
        if not any(shape.has_text_frame and shape.text.strip() for shape in slide.shapes):
            issues.append(_issue(index, None, "empty_slide", "blocking", "Слайд не содержит редактируемого текста.", "manual"))
    return issues


def font_issues(pptx_path: Path, expectations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Font Inspector: generated text keeps the template's typeface and size; the deck carries its fonts."""
    from .fonts import CORE_FONTS, effective_font, embedded_families, used_families
    from .typography import _key, font_files

    deck = Presentation(str(pptx_path))
    issues: list[dict[str, Any]] = []
    for item in expectations:
        index = item["slide"] - 1
        if not 0 <= index < len(deck.slides):
            continue
        shape = next((shape for shape in deck.slides[index].shapes if shape.shape_id == item["shape_id"]), None)
        if shape is None or not shape.has_text_frame:
            continue
        families, sizes = set(), set()
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.text.strip():
                    family, size, _ = effective_font(shape, paragraph, run)
                    families.add(family or "")
                    sizes.add(round(size, 1))
        expected_family, expected_size = item["family"] or "", float(item["size"])
        wrong = sorted(family for family in families if _key(family) != _key(expected_family))
        if wrong:
            issues.append(_issue(index, shape.shape_id, "font_mismatch", "blocking",
                                 f"Шрифт «{', '.join(wrong)}» вместо «{expected_family}» из шаблона.", "manual"))
        wrong_sizes = sorted(size for size in sizes if abs(size - expected_size) > 0.05)
        if wrong_sizes:
            issues.append(_issue(index, shape.shape_id, "font_size_mismatch", "blocking",
                                 f"Кегль {', '.join(f'{size:g}' for size in wrong_sizes)} pt вместо {expected_size:g} pt из шаблона.",
                                 "manual"))
    embedded = embedded_families(deck)
    for family in sorted(used_families(deck)):
        if family in embedded or _key(family) in CORE_FONTS:
            continue
        where = "" if font_files(family) else " Он не найден ни в системе, ни в хранилище шрифтов."
        issues.append(_issue(0, None, "font_not_embedded", "warning",
                             f"Шрифт «{family}» не встроен в PPTX: на компьютере без него PowerPoint подставит другой.{where}",
                             "manual"))
    return issues


def source_texts_from_template(path: Path) -> list[str]:
    template = Presentation(str(path))
    return [
        shape.text.strip()
        for slide in template.slides
        for shape in slide.shapes
        if shape.has_text_frame and shape.text.strip()
    ]
