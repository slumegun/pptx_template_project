"""Isolated PDFium fallback: native PDFium is not safe in concurrent threads."""
from pathlib import Path
import sys
import pypdfium2 as pdfium


def render(source: Path, output: Path, dpi: int):
    output.mkdir(parents=True, exist_ok=True)
    with pdfium.PdfDocument(source) as document:
        for index in range(len(document)):
            page = document[index]
            bitmap = page.render(scale=dpi / 72)
            image = bitmap.to_pil()
            try:
                image.save(output / f'slide-{index + 1:02d}.png')
            finally:
                image.close()
                bitmap.close()
                page.close()


if __name__ == '__main__':
    render(Path(sys.argv[1]), Path(sys.argv[2]), int(sys.argv[3]))
