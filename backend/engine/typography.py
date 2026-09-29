"""Font metrics shared by the writer and the deterministic layout audit."""
from functools import lru_cache
import os
from pathlib import Path
import re
from PIL import ImageFont


def _key(value):
    return re.sub(r'[^\w]', '', value.casefold())


def font_dir() -> Path:
    """The application's own font store: template fonts fetched once, no OS installation."""
    return Path(os.environ.get('AYA_FONT_DIR') or Path(__file__).resolve().parents[1] / 'data' / 'fonts')


@lru_cache(maxsize=1)
def _fonts():
    store = font_dir()
    roots = [Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Fonts',
             Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'Microsoft/Windows/Fonts',
             Path('/usr/share/fonts'), Path('/usr/local/share/fonts'), Path.home() / '.local/share/fonts', store]
    result = {}
    for root in roots:
        if not root.is_dir():
            continue
        for path in sorted(root.rglob('*')):
            if path.suffix.lower() not in {'.ttf', '.otf', '.ttc'}:
                continue
            try:
                family, style = ImageFont.truetype(str(path), 40).getname()
            except OSError:
                continue
            result.setdefault(_key(family), []).append((str(path), style.lower()))
            # A stored font is also known under the name the template uses
            # (its folder), e.g. "Montserrat SemiBold" for a Montserrat 600 file.
            alias = _key(path.parent.name.replace('_', ' '))
            if root == store and path.parent != store and alias != _key(family):
                result.setdefault(alias, []).append((str(path), style.lower()))
    return result


def refresh_fonts():
    """Forget cached lookups after fonts were added to the store."""
    _fonts.cache_clear()
    resolve_font.cache_clear()
    _face.cache_clear()


def font_files(family):
    """Installed or stored files of one family as (path, style) pairs."""
    return list(_fonts().get(_key(family or ''), []))


@lru_cache(maxsize=256)
def resolve_font(family='Arial', bold=False, italic=False):
    catalog = _fonts()
    for requested in [family, 'Arial', 'Liberation Sans', 'DejaVu Sans']:
        entries = catalog.get(_key(requested or ''))
        if entries:
            def score(item):
                style = item[1]
                return (('bold' in style) != bool(bold)) + (('italic' in style or 'oblique' in style) != bool(italic))
            path, _ = min(entries, key=score)
            return ImageFont.truetype(path, 40).getname()[0], path
    raise RuntimeError('Для экспорта нужен установленный шрифт Arial, Liberation Sans или DejaVu Sans.')


@lru_cache(maxsize=512)
def _face(family, size, bold, italic):
    _, path = resolve_font(family, bold, italic)
    return ImageFont.truetype(path, max(1, round(size * 4)))


def text_height(lines, width_points, font_size, family='Arial', bold=False, italic=False, line_factor=1.0):
    """Measure word wrapping, including explicit breaks and long identifiers.

    line_factor is the paragraph's proportional line spacing (0.9 for 90 %).
    """
    face = _face(family, font_size, bold, italic)
    limit = max(1, width_points * 4)
    count = 0
    for text in lines:
        for paragraph in re.split(r'[\n\v]', text):
            current = ''
            count += 1
            for word in paragraph.split():
                candidate = (current + ' ' + word) if current else word
                if face.getlength(candidate) <= limit:
                    current = candidate
                    continue
                if current:
                    count += 1
                    current = ''
                # PowerPoint wraps unbroken strings at character boundaries.
                for char in word:
                    if current and face.getlength(current + char) > limit:
                        count += 1
                        current = ''
                    current += char
    ascent, descent = face.getmetrics()
    return count * max(font_size * 1.2, (ascent + descent) / 4 * 1.15) * line_factor
