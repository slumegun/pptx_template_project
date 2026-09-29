"""Small, bounded checks before an uploaded document reaches a parser."""

import tempfile
import zipfile
from pathlib import Path, PurePosixPath

from fastapi import HTTPException, UploadFile

from .config import get_settings


def stage_upload(file: UploadFile, suffix: str, directory: Path, limit: int) -> Path:
    """Copy an upload to a temporary file, enforcing the size limit while reading."""
    size = 0
    staged_path = None
    try:
        with tempfile.NamedTemporaryFile(dir=directory, suffix=suffix, delete=False) as staged:
            staged_path = Path(staged.name)
            while chunk := file.file.read(1024 * 1024):
                size += len(chunk)
                if size > limit:
                    raise HTTPException(status_code=413, detail="File is too large")
                staged.write(chunk)
    except Exception:
        if staged_path is not None:
            staged_path.unlink(missing_ok=True)
        raise
    finally:
        file.file.close()
    if size == 0:
        staged_path.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail="Empty file")
    return staged_path


def validate_upload(path: Path, suffix: str) -> None:
    if suffix == ".pdf":
        with path.open("rb") as source:
            if not source.read(1024).lstrip().startswith(b"%PDF-"):
                raise ValueError("Invalid PDF")
    if suffix not in {".pptx", ".docx"}:
        return
    required_part = "ppt/presentation.xml" if suffix == ".pptx" else "word/document.xml"
    settings = get_settings()
    try:
        with zipfile.ZipFile(path) as archive:
            entries = archive.infolist()
            if len(entries) > settings.max_archive_entries:
                raise ValueError("Document contains too many archive entries")
            if sum(item.file_size for item in entries) > settings.max_archive_bytes:
                raise ValueError("Unpacked document is too large")
            names = set()
            for item in entries:
                name = item.filename.replace("\\", "/")
                parts = PurePosixPath(name)
                if parts.is_absolute() or ".." in parts.parts or ":" in name or name in names:
                    raise ValueError("Invalid document archive path")
                if item.flag_bits & 1:
                    raise ValueError("Encrypted documents are not supported")
                names.add(name)
            if required_part not in names:
                raise ValueError("Document is missing its main part")
            if archive.testzip() is not None:
                raise ValueError("Document archive is damaged")
    except (zipfile.BadZipFile, RuntimeError, NotImplementedError) as exc:
        raise ValueError("Invalid document archive") from exc
