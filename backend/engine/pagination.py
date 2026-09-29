"""Lossless pagination using the same geometry and font metrics as rendering.

Pages are measured at the template's own body and title sizes, so continuation
pages keep its typography; BODY_PT and TITLE_PT serve callers without a template.
"""
from dataclasses import replace
import re
from .models import SlideContent
from .typography import text_height

BODY_PT = 16
TITLE_PT = 28


def dense_geometry(width, height, title, family, title_pt=TITLE_PT, title_family=None):
    """Points, with room for the heading and export/font metric differences."""
    w, h = width / 12700, height / 12700
    margin = min(36, w * .05, h * .06)
    inner = w - 2 * margin
    title_h = text_height([title], inner - 12, title_pt, title_family or family, True) + 12
    top = margin + title_h + 16
    return margin, margin, inner, title_h, top, h - margin - top


def body_fits(lines, width, height, family, body_pt=BODY_PT):
    return text_height(lines, width - 12, body_pt, family) <= (height - 8) * .88


def split_bullets(bullets, width, height, family, body_pt=BODY_PT):
    """Keep complete paragraphs when possible; split huge paragraphs losslessly."""
    pages, page = [], []
    for bullet in bullets:
        rest = bullet
        while rest:
            if body_fits([*page, rest], width, height, family, body_pt):
                page.append(rest)
                break
            if page:
                pages.append(page)
                page = []
                continue
            # Search a prefix without quadratic character-by-character measuring.
            low, high = 0, len(rest)
            while low < high:
                mid = (low + high + 1) // 2
                if body_fits([rest[:mid]], width, height, family, body_pt):
                    low = mid
                else:
                    high = mid - 1
            if low == 0:
                raise ValueError(f"Размер слайда не позволяет разместить строку текста размером {body_pt:g} pt.")
            # Prefer a sentence, then a word boundary. Preserve every character.
            boundaries = [m.end() for m in re.finditer(r"[.!?][ \n]+", rest[:low])]
            boundary = boundaries[-1] if boundaries and boundaries[-1] >= low // 2 else 0
            if not boundary:
                spaces = [m.end() for m in re.finditer(r"\s+", rest[:low])]
                boundary = spaces[-1] if spaces and spaces[-1] >= low // 2 else low
            pages.append([rest[:boundary]])
            rest = rest[boundary:]
    if page:
        pages.append(page)
    return pages or [[]]


def paginate_slides(slides, compositions, width, height, family, body_pt=BODY_PT, title_pt=TITLE_PT,
                    title_family=None):
    """Pin source layouts before expanding so all variants have identical pages."""
    result = []
    for original, composition in zip(slides, compositions):
        continuation_title = original.title + " (продолжение)"
        # Use the longer heading for every part to keep the capacity consistent.
        _, _, body_w, _, _, body_h = dense_geometry(width, height, continuation_title, family,
                                                    title_pt, title_family)
        pages = split_bullets(original.bullets, body_w, body_h, family, body_pt)
        for part, bullets in enumerate(pages):
            result.append(replace(original, slide_id=f"slide_{len(result)+1}",
                title=original.title if part == 0 else continuation_title,
                bullets=bullets, fact_ids=list(original.fact_ids),
                source_slide_index=composition.source_slide_index,
                dense_layout=len(pages) > 1,
                continuation_of=original.slide_id))
    return result
