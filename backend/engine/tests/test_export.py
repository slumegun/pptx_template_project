from pathlib import Path
import subprocess

import pytest
from pypdf import PdfWriter

from engine import export


def write_pdf(path):
    writer = PdfWriter()
    writer.add_blank_page(width=720, height=540)
    with path.open("wb") as stream:
        writer.write(stream)


def test_export_retries_transient_failure_in_new_profile(tmp_path, monkeypatch):
    source = tmp_path / "Шаблон с пробелами.pptx"
    source.write_bytes(b"immutable source")
    attempts = []
    monkeypatch.setattr(export, "_find_soffice", lambda: "soffice")

    def run(command, cwd, timeout):
        assert (cwd / "input.pptx").read_bytes() == source.read_bytes()
        assert timeout <= 30
        attempts.append(command[1])
        if len(attempts) == 1:
            return subprocess.CompletedProcess(command, 1, "", "Document is empty")
        write_pdf(cwd / "input.pdf")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(export, "_run_converter", run)
    pdf = export.export_pdf(source, tmp_path / "output", timeout=30)
    assert export._valid_pdf(pdf)
    assert len(attempts) == 2 and attempts[0] != attempts[1]
    assert source.read_bytes() == b"immutable source"


@pytest.mark.parametrize("output", [None, b"", b"not a PDF"])
def test_failed_export_does_not_accept_or_replace_stale_pdf(tmp_path, monkeypatch, output):
    source = tmp_path / "source.pptx"
    source.write_bytes(b"source")
    old_pdf = tmp_path / "source.pdf"
    write_pdf(old_pdf)
    before = old_pdf.read_bytes()
    monkeypatch.setattr(export, "_find_soffice", lambda: "soffice")

    def run(command, cwd, timeout):
        if output is not None:
            (cwd / "input.pdf").write_bytes(output)
        return subprocess.CompletedProcess(command, 0, "", "Document is empty")

    monkeypatch.setattr(export, "_run_converter", run)
    with pytest.raises(RuntimeError, match="Document is empty"):
        export.export_pdf(source, tmp_path)
    assert old_pdf.read_bytes() == before


def test_export_timeout_is_actionable(tmp_path, monkeypatch):
    source = tmp_path / "source.pptx"
    source.write_bytes(b"source")
    monkeypatch.setattr(export, "_find_soffice", lambda: "soffice")

    def run(command, cwd, timeout):
        raise subprocess.TimeoutExpired(command, timeout)

    monkeypatch.setattr(export, "_run_converter", run)
    with pytest.raises(RuntimeError, match="время ожидания"):
        export.export_pdf(source, tmp_path)

def test_preview_crash_uses_fallback_and_rejects_partial_results(tmp_path, monkeypatch):
    from PIL import Image
    source = tmp_path / 'deck.pdf'
    write_pdf(source)
    monkeypatch.setattr(export, '_find_pdftoppm', lambda: 'pdftoppm')
    calls = []
    def run(command, **kwargs):
        calls.append(command)
        if command[0] == 'pdftoppm':
            return subprocess.CompletedProcess(command, -1073741819, '', '')
        output = Path(command[-2])
        Image.new('RGB', (10, 10), 'white').save(output/'slide-01.png')
        return subprocess.CompletedProcess(command, 0, '', '')
    monkeypatch.setattr(export.subprocess, 'run', run)
    paths = export.export_previews(source, tmp_path/'previews')
    assert len(calls) == 2
    assert [p.name for p in paths] == ['slide-01.png']
    assert paths[0].is_file()
