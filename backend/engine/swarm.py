"""Isolated model roles with checked outputs and one shared generation deadline."""

from __future__ import annotations

import base64
import io
import json
import logging
import os
import re
import time
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .audit import _issue
from .models import Fact, PreparedTemplate, SlideContent
from .provider import ModelClient, ModelProviderError
from .prompt_text import prompt_text


def remaining_seconds(deadline: float | None, maximum: float = 120) -> float:
    remaining = maximum if deadline is None else min(maximum, deadline - time.monotonic())
    if remaining <= 0:
        raise TimeoutError("Истёк общий бюджет генерации трёх вариантов (300 секунд).")
    return remaining


def numeric_tokens(text: str) -> set[str]:
    # Treat decimal commas and points equally without discarding signs or units.
    return {value.replace(",", ".").replace(" ", "") for value in re.findall(r"(?<!\w)[+-]?\d+(?:[.,]\d+)?(?:\s*%)?", text)}


# Concrete capabilities need a primary source even when a brief contains
# only a topic. These are especially easy to mistake for generic explanation.
_CAPABILITY_CLAIMS = {
    "режим реального времени": r"\bв\s+(?:режиме\s+)?реальн(?:ом|ого)\s+времени\b|\breal[ -]?time\b",
    "автоматическое выполнение": r"\bавтоматическ\w*\b|\bбез\s+участия\s+человека\b",
    "непрерывная работа": r"\bкруглосуточн\w*\b|\b24/7\b",
    "мгновенная работа": r"\bмгновенн\w*\b",
    "гарантированный результат": r"\bгарантированн\w*\b|\bбезошибочн\w*\b",
}


def _check_capability_claims(text: str, excerpts: str) -> None:
    for label, pattern in _CAPABILITY_CLAIMS.items():
        if re.search(pattern, text, re.IGNORECASE) and not re.search(pattern, excerpts, re.IGNORECASE):
            raise ModelProviderError(f"Слайд добавляет неподтверждённую возможность: {label}")


def validate_slide(
    value: dict[str, Any], slide_id: str, facts: dict[str, Fact], *, required_refs: set[str] | None = None
) -> SlideContent:
    if not isinstance(value, dict):
        raise ModelProviderError("Модель вернула слайд неверного формата")
    title, bullets, refs = value.get("title"), value.get("bullets"), value.get("fact_ids")
    if not isinstance(title, str) or not title.strip() or len(title) > 150:
        raise ModelProviderError("Заголовок слайда отсутствует или слишком длинный")
    if not isinstance(bullets, list) or not 1 <= len(bullets) <= 6:
        raise ModelProviderError("Слайду нужны от одного до шести тезисов")
    if any(not isinstance(item, str) or not item.strip() or len(item) > 450 for item in bullets):
        raise ModelProviderError("Некорректный текст тезиса")
    if not isinstance(refs, list) or not refs or any(not isinstance(ref, str) or ref not in facts for ref in refs):
        raise ModelProviderError("Слайд ссылается на отсутствующий источник")
    if required_refs is not None and set(refs) != required_refs:
        raise ModelProviderError("Slide Worker изменил набор фактов задания")
    excerpts = " ".join(facts[ref].excerpt for ref in refs)
    _check_capability_claims(" ".join([title, *bullets]), excerpts)
    missing = numeric_tokens(" ".join([title, *bullets])) - numeric_tokens(excerpts)
    if missing:
        raise ModelProviderError("Слайд добавляет числа вне первичных источников: " + ", ".join(sorted(missing)))
    return SlideContent(slide_id, title.strip(), [item.strip() for item in bullets], list(dict.fromkeys(refs)))


def contact_sheet(paths: list[Path]) -> str:
    from PIL import Image, ImageDraw

    width, height = 1280, 760
    canvas = Image.new("RGB", (width * 2, height * ((len(paths) + 1) // 2)), "white")
    draw = ImageDraw.Draw(canvas)
    for index, path in enumerate(paths):
        x, y = index % 2 * width, index // 2 * height
        with Image.open(path) as source:
            preview = source.convert("RGB")
            preview.thumbnail((width - 16, height - 34))
            canvas.paste(preview, (x + 8, y + 26))
        draw.text((x + 8, y + 6), f"Panel {index + 1}", fill="black")
    buffer = io.BytesIO()
    canvas.save(buffer, format="JPEG", quality=92)
    return "data:image/jpeg;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")


def analyze_template(
    client: ModelClient, template_path: Path, template: PreparedTemplate, previews: list[Path], checkpoint=None
) -> None:
    """Classify every source slide; geometry always comes from OOXML."""
    # Preparation is outside the generation timer; account limits remain shared.
    analysis_rpm = float(os.getenv("MODEL_ANALYSIS_RPM", "60"))
    if not 0 < analysis_rpm <= 6000:
        raise ValueError("MODEL_ANALYSIS_RPM must be greater than 0 and at most 6000")
    last_request: float | None = None
    deck = Presentation(str(template_path))
    if len(previews) != len(deck.slides):
        raise ValueError("Число превью анализа не совпадает с числом исходных слайдов")
    allowed_roles = {"fixed_brand", "replaceable", "reusable_asset", "remove", "unresolved"}

    class InvalidTemplateBatch(ModelProviderError):
        """A response-content failure that can improve with a smaller image batch."""

    batches = []
    batch = []
    shape_count = 0
    for composition in template.compositions:
        if composition.analysis_complete:
            continue
        if batch and (len(batch) >= 6 or shape_count + composition.shape_count > 65):
            batches.append(batch)
            batch, shape_count = [], 0
        batch.append(composition)
        shape_count += composition.shape_count
    if batch:
        batches.append(batch)

    unsafe_missing_types = {
        MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE,
        MSO_SHAPE_TYPE.GROUP, MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.WEB_VIDEO,
        MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT, MSO_SHAPE_TYPE.LINKED_OLE_OBJECT,
        MSO_SHAPE_TYPE.OLE_CONTROL_OBJECT, MSO_SHAPE_TYPE.CHART,
        MSO_SHAPE_TYPE.TABLE, MSO_SHAPE_TYPE.DIAGRAM, MSO_SHAPE_TYPE.IGX_GRAPHIC,
    }
    simple_types = {
        MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.CALLOUT,
        MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.LINE,
        MSO_SHAPE_TYPE.PLACEHOLDER, MSO_SHAPE_TYPE.TEXT_BOX,
        MSO_SHAPE_TYPE.TEXT_EFFECT,
    }

    def fallback_role(shape) -> str:
        if shape.shape_type in unsafe_missing_types or shape.has_chart or shape.has_table:
            return "unresolved"
        return "replaceable" if shape.has_text_frame or shape.shape_type in simple_types else "unresolved"

    def group_has_old_content(shape) -> bool:
        for child in shape.shapes:
            if (child.has_text_frame and child.text.strip()) or child.has_chart or child.has_table:
                return True
            if child.shape_type == MSO_SHAPE_TYPE.GROUP and group_has_old_content(child):
                return True
        return False

    def classify_batch(items: list) -> dict[int, tuple[dict[str, str], str, list[str], int, int]]:
        nonlocal last_request
        inventory = []
        for panel, composition in enumerate(items, 1):
            source = deck.slides[composition.source_slide_index]
            objects = [
                {"object_id": str(shape.shape_id), "type": str(shape.shape_type),
                 "box": [round(shape.left / template.width, 3), round(shape.top / template.height, 3), round(shape.width / template.width, 3), round(shape.height / template.height, 3)],
                 "text": shape.text[:120] if shape.has_text_frame else ""}
                for shape in source.shapes
            ]
            inventory.append({"panel": panel, "source_slide_index": composition.source_slide_index, "objects": objects})
        if last_request is not None:
            delay = last_request + 60 / analysis_rpm - time.monotonic()
            if delay > 0:
                time.sleep(delay)
        last_request = time.monotonic()
        try:
            response = client.complete_json(
                prompt_text("template_analyst"),
                json.dumps(inventory, ensure_ascii=False, separators=(",", ":")), vision=True,
                image_data_url=contact_sheet([previews[item.source_slide_index] for item in items]),
                max_tokens=min(4800, max(1200, sum(item.shape_count for item in items) * 22 + len(items) * 120)),
            )
        except ModelProviderError as error:
            # Only malformed/truncated model output is improved by smaller batches.
            # HTTP/auth/balance/rate errors must propagate without extra paid calls.
            if str(error) in {
                "Model response was truncated; reduce the task or increase max_tokens",
                "Model response is not a valid JSON object",
                "Model response is not a JSON object",
            }:
                raise InvalidTemplateBatch(str(error)) from error
            raise
        rows = response.get("slides") if isinstance(response, dict) else None
        if not isinstance(rows, list) or len(rows) != len(items):
            raise InvalidTemplateBatch("Template Analyst вернул неполную классификацию")
        by_index = {}
        for row in rows:
            if not isinstance(row, dict) or type(row.get("source_slide_index")) is not int:
                raise InvalidTemplateBatch("Template Analyst вернул некорректный слайд")
            by_index[row["source_slide_index"]] = row
        if set(by_index) != {item.source_slide_index for item in items}:
            raise InvalidTemplateBatch("Template Analyst изменил идентификаторы исходных слайдов")
        classified = {}
        for item in items:
            row = by_index[item.source_slide_index]
            shapes_by_id = {str(shape.shape_id): shape for shape in deck.slides[item.source_slide_index].shapes}
            expected = set(shapes_by_id)
            roles: dict[str, str] = {}
            conflicts: set[str] = set()
            if "roles" in row:
                groups = row["roles"]
                if not isinstance(groups, dict):
                    raise InvalidTemplateBatch("Template Analyst вернул неверную структуру ролей")
                for role, ids in groups.items():
                    if not isinstance(ids, list):
                        raise InvalidTemplateBatch("Template Analyst вернул неверный список объектов")
                    if role not in allowed_roles:
                        continue
                    for object_id in ids:
                        key = str(object_id)
                        if key not in expected:
                            continue
                        if key in roles and roles[key] != role:
                            conflicts.add(key)
                        else:
                            roles[key] = role
            else:
                raw_roles = row.get("object_roles")
                if not isinstance(raw_roles, dict):
                    raise InvalidTemplateBatch("Template Analyst вернул неверную структуру ролей")
                for object_id, role in raw_roles.items():
                    key = str(object_id)
                    if key in expected and isinstance(role, str) and role in allowed_roles:
                        roles[key] = role
            for key in conflicts:
                roles.pop(key, None)
            missing = expected - set(roles)
            for key in missing:
                roles[key] = fallback_role(shapes_by_id[key])
            normalized = 0
            for key, role in list(roles.items()):
                shape = shapes_by_id[key]
                if shape.has_text_frame and role in {"fixed_brand", "reusable_asset"}:
                    roles[key] = "replaceable"
                    normalized += 1
                elif (shape.shape_type == MSO_SHAPE_TYPE.GROUP and
                      role in {"fixed_brand", "reusable_asset"} and
                      group_has_old_content(shape)):
                    roles[key] = "unresolved"
                    normalized += 1
            archetype = row.get("archetype")
            archetype = archetype if isinstance(archetype, str) and archetype in {"title", "content", "comparison", "data", "closing"} else "unknown"
            notes = row.get("notes", [])
            analysis_notes = [note[:600] for note in notes[:8] if isinstance(note, str)] if isinstance(notes, list) else []
            classified[item.source_slide_index] = (roles, archetype, analysis_notes, len(missing), normalized)
        return classified

    def save_classification(items: list, classified: dict[int, tuple[dict[str, str], str, list[str], int, int]]) -> None:
        for item in items:
            roles, archetype, notes, missing_count, normalized_count = classified[item.source_slide_index]
            item.object_roles, item.archetype, item.analysis_notes = roles, archetype, notes
            item.analysis_complete = True
            slide_number = item.source_slide_index + 1
            if missing_count:
                template.warnings.append(
                    f"Template Analyst: слайд {slide_number}: роли {missing_count} объектов восстановлены безопасно."
                )
            if normalized_count:
                template.warnings.append(
                    f"Template Analyst: слайд {slide_number}: роли {normalized_count} текстовых или групповых объектов исправлены для удаления старого содержания."
                )
        if checkpoint:
            checkpoint(template)

    for batch_number, items in enumerate(batches, 1):
        logging.getLogger(__name__).info("Template analyst: batch %s/%s", batch_number, len(batches))
        try:
            classified = classify_batch(items)
        except InvalidTemplateBatch as error:
            if len(items) == 1:
                index = items[0].source_slide_index
                raise ModelProviderError(f"Template Analyst: слайд {index + 1} (source_slide_index={index}): {error}") from error
            logging.getLogger(__name__).warning(
                "Template analyst: batch %s invalid (%s); retrying %s slides separately",
                batch_number, error, len(items),
            )
            for item in items:
                index = item.source_slide_index
                try:
                    single_classification = classify_batch([item])
                except InvalidTemplateBatch as single_error:
                    raise ModelProviderError(
                        f"Template Analyst: слайд {index + 1} (source_slide_index={index}): {single_error}"
                    ) from single_error
                save_classification([item], single_classification)
        else:
            save_classification(items, classified)
    template.analysis_mode = "vision_model"
    template.analysis_model = client.vision_model


def critique_content(
    client: ModelClient, slides: list[SlideContent], facts: list[Fact], deadline: float
) -> list[dict[str, Any]]:
    """One independent deck critic sees sources, never another agent's dialogue."""
    slide_ids = [slide.slide_id for slide in slides]
    schema = {
        "type": "object", "additionalProperties": False, "required": ["issues"],
        "properties": {"issues": {
            "type": "array", "maxItems": 40,
            "items": {"type": "object", "additionalProperties": False,
                      "required": ["slide_id", "rule_id", "severity", "evidence"],
                      "properties": {
                          "slide_id": {"type": "string", "enum": slide_ids},
                          "rule_id": {"type": "string", "enum": ["unsupported_claim", "contradiction", "repetition", "narrative_gap"]},
                          "severity": {"type": "string", "enum": ["blocking", "warning"]},
                          "evidence": {"type": "string", "minLength": 1, "maxLength": 1200},
                      }},
        }},
    }
    response = client.complete_json(
        prompt_text("deck_critic"),
        json.dumps({"slides": [slide.to_dict() for slide in slides], "primary_sources": [{"fact_id": fact.fact_id, "excerpt": fact.excerpt} for fact in facts]}, ensure_ascii=False),
        max_tokens=min(4000, max(1200, len(slides) * 180)), deadline=deadline, schema=schema,
    )
    rows = response.get("issues")
    if not isinstance(rows, list) or len(rows) > 40:
        raise ModelProviderError("Deck Critic вернул неверный отчёт")
    indices = {slide.slide_id: index for index, slide in enumerate(slides)}
    issues = []
    for row in rows:
        if not isinstance(row, dict) or row.get("slide_id") not in indices or row.get("rule_id") not in {"unsupported_claim", "contradiction", "repetition", "narrative_gap"}:
            raise ModelProviderError("Deck Critic вернул неизвестный слайд или правило")
        if row.get("severity") not in {"blocking", "warning"} or not isinstance(row.get("evidence"), str) or not row["evidence"].strip():
            raise ModelProviderError("Deck Critic не указал доказательство замечания")
        severity = "blocking" if row["rule_id"] in {"unsupported_claim", "contradiction"} else row["severity"]
        issues.append(_issue(indices[row["slide_id"]], None, row["rule_id"], severity, row["evidence"][:1200], "manual"))
    return issues



def critique_prompt(
    client: ModelClient, brief: str, slides: list[SlideContent], facts: list[Fact], deadline: float
) -> list[dict[str, Any]]:
    """Independently check the deck against the user's original assignment.

    Deck-wide omissions are attached to the first slide because the issue API
    currently requires a slide identifier even for a whole-deck observation.
    """
    if not slides:
        raise ValueError("Prompt critic needs at least one slide")
    if not brief.strip():
        raise ValueError("Prompt critic needs the original brief")
    slide_indices = {slide.slide_id: index for index, slide in enumerate(slides)}
    if len(slide_indices) != len(slides):
        raise ValueError("Prompt critic needs unique slide identifiers")
    rules = {
        "prompt_topic_mismatch",
        "prompt_required_point_missing",
        "prompt_order_mismatch",
        "prompt_unsupported_claim",
    }
    schema = {
        "type": "object", "additionalProperties": False, "required": ["issues"],
        "properties": {"issues": {
            "type": "array", "maxItems": 40,
            "items": {"type": "object", "additionalProperties": False,
                      "required": ["slide_id", "rule_id", "severity", "evidence"],
                      "properties": {
                          "slide_id": {"type": "string", "enum": list(slide_indices)},
                          "rule_id": {"type": "string", "enum": sorted(rules)},
                          "severity": {"type": "string", "enum": ["warning", "blocking"]},
                          "evidence": {"type": "string", "minLength": 1, "maxLength": 1200},
                      }},
        }},
    }
    response = client.complete_json(
        prompt_text("prompt_critic"),
        json.dumps({
            "brief": brief,
            "slides": [slide.to_dict() for slide in slides],
            "primary_sources": [
                {"fact_id": fact.fact_id, "excerpt": fact.excerpt,
                 "source": fact.source, "location": fact.location}
                for fact in facts
            ],
        }, ensure_ascii=False),
        max_tokens=min(4000, max(1200, len(slides) * 180)),
        deadline=deadline,
        schema=schema,
    )
    if not isinstance(response, dict) or set(response) != {"issues"}:
        raise ModelProviderError("Prompt Critic returned an invalid report")
    rows = response.get("issues")
    if not isinstance(rows, list) or len(rows) > 40:
        raise ModelProviderError("Prompt Critic returned an invalid report")
    issues = []
    seen = set()
    for row in rows:
        if (not isinstance(row, dict) or set(row) != {"slide_id", "rule_id", "severity", "evidence"}
                or row.get("slide_id") not in slide_indices or row.get("rule_id") not in rules):
            raise ModelProviderError("Prompt Critic returned an unknown slide or rule")
        if (row.get("severity") not in {"blocking", "warning"}
                or not isinstance(row.get("evidence"), str)
                or not row["evidence"].strip()
                or len(row["evidence"]) > 1200):
            raise ModelProviderError("Prompt Critic returned invalid issue evidence")
        key = (row["slide_id"], row["rule_id"], row["evidence"].strip())
        if key in seen:
            continue
        seen.add(key)
        # Each rule is a concrete violation of the user's explicit request.
        issues.append(_issue(slide_indices[row["slide_id"]], None, row["rule_id"],
                             "blocking", row["evidence"].strip(), "manual"))
    return issues


def critique_visual_slide(client: ModelClient, source: Path, rendered: Path,
                          slide: SlideContent, index: int, deadline: float,
                          variant_theme: str = "source") -> list[dict[str, Any]]:
    """Compare a rendered slide with its actual source composition."""
    if variant_theme not in {"source", "dark", "reflow"}:
        raise ValueError("Unknown variant theme")
    rules = {"visual_text_clipping", "visual_overlap", "visual_low_contrast", "visual_empty_container", "visual_brand_drift"}
    schema = {"type": "object", "additionalProperties": False, "required": ["issues"], "properties": {
        "issues": {"type": "array", "maxItems": 4, "items": {"type": "object", "additionalProperties": False,
            "required": ["rule_id", "severity", "evidence"], "properties": {
                "rule_id": {"type": "string", "enum": sorted(rules)},
                "severity": {"type": "string", "enum": ["warning", "blocking"]},
                "evidence": {"type": "string", "minLength": 1, "maxLength": 180}}}}}}
    response = client.complete_json(
        prompt_text("slide_critic"),
        json.dumps({"expected_title": slide.title, "expected_bullets": slide.bullets,
                    "variant_theme": variant_theme}, ensure_ascii=False),
        vision=True, image_data_url=contact_sheet([source, rendered]), max_tokens=1800, deadline=deadline, schema=schema,
    )
    rows = response.get("issues")
    rules = {"visual_text_clipping", "visual_overlap", "visual_low_contrast", "visual_empty_container", "visual_brand_drift"}
    if not isinstance(rows, list) or len(rows) > 4:
        raise ModelProviderError("Visual Critic returned an invalid issue list")
    issues = []
    for row in rows:
        if (not isinstance(row, dict) or row.get("rule_id") not in rules
                or row.get("severity") not in {"warning", "blocking"}
                or not isinstance(row.get("evidence"), str) or not row["evidence"].strip()):
            raise ModelProviderError("Visual Critic returned an invalid observation")
        severity = "warning" if row["rule_id"] == "visual_empty_container" else row["severity"]
        issues.append(_issue(index, None, row["rule_id"], severity, row["evidence"][:1200], "manual"))
    return issues


def critique_visual_variant(client: ModelClient, sources: list[Path], previews: list[Path],
                            source_slides: list[int], slides: list[SlideContent], deadline: float,
                            variant_theme: str = "source") -> list[dict[str, Any]]:
    from concurrent.futures import ThreadPoolExecutor
    if variant_theme not in {"source", "dark", "reflow"}:
        raise ValueError("Unknown variant theme")
    if len(previews) != len(slides) or len(source_slides) != len(slides):
        raise ValueError("Visual audit slide mapping is incomplete")
    if any(type(index) is not int or not 1 <= index <= len(sources) for index in source_slides):
        raise ValueError("Visual audit references an unknown source slide")
    with ThreadPoolExecutor(max_workers=client.max_parallel) as executor:
        futures = []
        for i, slide in enumerate(slides):
            args = (client, sources[source_slides[i] - 1], previews[i], slide, i, deadline)
            # Preserve the six-argument call for existing source-theme clients.
            futures.append(executor.submit(critique_visual_slide, *args) if variant_theme == "source"
                           else executor.submit(critique_visual_slide, *args, variant_theme))
        try:
            return [issue for future in futures for issue in future.result()]
        except BaseException:
            for future in futures:
                future.cancel()
            raise
