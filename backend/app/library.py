"""Shared template library: analysis once, reuse by any account, portable packages."""

import json
import logging
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from uuid import uuid4

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from sqlalchemy import select
from sqlalchemy.orm import Session
from starlette.background import BackgroundTask

from engine.fonts import CORE_FONTS, embedded_families, template_families
from engine.ingest import save_prepared, verify_imported
from engine.typography import _key, font_dir, font_files, refresh_fonts

from .auth import current_user
from .config import get_settings
from .db import get_db, get_session_factory, locked_get
from .models import Source, Template, User, uid
from .queueing import enqueue_template
from .schemas import PreparationOut, TemplateOut
from .storage import get_storage
from .uploads import stage_upload, validate_upload

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/api/templates")

PACKAGE_FORMAT = "lukas-template"
PACKAGE_VERSION = 1
MAX_PREVIEWS = 80
MAX_PREVIEW_BYTES = 5 * 1024 * 1024
PREVIEW_NAME = re.compile(r"previews/slide-(\d{1,3})\.png")
FONT_NAME = re.compile(r"fonts/[A-Za-z0-9][A-Za-z0-9 _.-]{0,79}/[A-Za-z0-9][A-Za-z0-9 _.-]{0,79}\.(?:ttf|otf)", re.IGNORECASE)
MAX_FONT_FILES = 24
MAX_FONT_BYTES = 20 * 1024 * 1024
FONT_MAGIC = (b"\x00\x01\x00\x00", b"true", b"OTTO")


def template_out(template: Template) -> TemplateOut:
    ready = template.preparation_status == "ready"
    previews = (template.metadata_json or {}).get("preview_keys") or []
    return TemplateOut(
        id=template.id, name=template.name, filename=template.filename, sha256=template.sha256,
        size_bytes=template.size_bytes, slide_count=template.slide_count, origin=template.origin,
        created_at=template.created_at,
        preparation=PreparationOut(status=template.preparation_status, stage=template.preparation_stage,
                                   error=template.preparation_error, prepared_at=template.prepared_at),
        preview_urls=[f"/api/templates/{template.id}/previews/{number}" for number in range(1, len(previews) + 1)],
        export_url=f"/api/templates/{template.id}/export" if ready else None,
    )


def require_template(db: Session, template_id: str) -> Template:
    template = db.get(Template, template_id)
    if template is None:
        raise HTTPException(status_code=404, detail="Шаблон не найден")
    return template


def schedule_template(db: Session, template: Template) -> TemplateOut:
    try:
        enqueue_template(template.id)
        template = locked_get(db, Template, template.id)
        if template.preparation_status == "pending_enqueue":
            template.preparation_status = "queued"
            template.preparation_stage = "queued"
            db.commit()
    except Exception:
        logger.exception("Could not enqueue template %s; reconciliation will retry", template.id)
        db.rollback()
        template = require_template(db, template.id)
    return template_out(template)


def _display_name(filename: str) -> str:
    return (Path(filename).stem.strip() or "Шаблон")[:200]


@router.get("", response_model=list[TemplateOut])
def list_templates(db: Session = Depends(get_db), _: User = Depends(current_user)):
    return [template_out(item) for item in db.scalars(select(Template).order_by(Template.created_at.desc()))]


@router.get("/{template_id}", response_model=TemplateOut)
def get_template(template_id: str, db: Session = Depends(get_db), _: User = Depends(current_user)):
    return template_out(require_template(db, template_id))


@router.post("", response_model=TemplateOut, status_code=201)
def upload_template(file: UploadFile = File(...), db: Session = Depends(get_db), _: User = Depends(current_user)):
    filename = Path((file.filename or "template.pptx").replace("\\", "/")).name[:200]
    if any(ord(char) < 32 for char in filename) or Path(filename).suffix.lower() != ".pptx":
        raise HTTPException(status_code=422, detail="Для шаблона нужен файл PPTX")
    storage = get_storage()
    staged = stage_upload(file, ".pptx", storage.root, get_settings().max_upload_bytes)
    try:
        try:
            validate_upload(staged, ".pptx")
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
        template_id = uid()
        stored = storage.put_file(staged, f"templates/{template_id}/original.pptx")
    finally:
        staged.unlink(missing_ok=True)
    # The same file needs one analysis: reuse a library entry that is ready or in progress.
    existing = db.scalar(select(Template).where(Template.sha256 == stored.sha256,
                                                Template.preparation_status != "failed")
                         .order_by(Template.created_at))
    if existing is not None:
        storage.delete(stored.key)
        return template_out(existing)
    template = Template(
        id=template_id, name=_display_name(filename), filename=filename, storage_key=stored.key,
        sha256=stored.sha256, size_bytes=stored.size_bytes, origin="upload",
        preparation_status="pending_enqueue", preparation_stage="waiting", metadata_json={},
    )
    db.add(template)
    try:
        db.commit()
    except Exception:
        db.rollback()
        storage.delete(stored.key)
        raise
    db.refresh(template)
    return schedule_template(db, template)


@router.post("/{template_id}/prepare", response_model=TemplateOut, status_code=202)
def retry_template(template_id: str, db: Session = Depends(get_db), _: User = Depends(current_user)):
    template = locked_get(db, Template, template_id)
    if template is None:
        raise HTTPException(status_code=404, detail="Шаблон не найден")
    if template.preparation_status != "failed":
        return template_out(template)
    template.preparation_status = "pending_enqueue"
    template.preparation_stage = "waiting"
    template.preparation_error = None
    db.commit()
    return schedule_template(db, template)


@router.get("/{template_id}/previews/{number}")
def template_preview(template_id: str, number: int, db: Session = Depends(get_db), _: User = Depends(current_user)):
    keys = (require_template(db, template_id).metadata_json or {}).get("preview_keys") or []
    if not 1 <= number <= len(keys):
        raise HTTPException(status_code=404, detail="Превью не найдено")
    path = get_storage().path(keys[number - 1])
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Превью не найдено")
    return FileResponse(path, media_type="image/png", headers={"X-Content-Type-Options": "nosniff"})


def _package_name(name: str) -> str:
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", name).strip(" .") or "template"
    return f"{safe[:120]}.template.zip"


@router.get("/{template_id}/export")
def export_template(template_id: str, db: Session = Depends(get_db), _: User = Depends(current_user)):
    """A self-contained package: the PPTX, its analysis and slide previews."""
    template = require_template(db, template_id)
    if template.preparation_status != "ready" or not template.prepared_key:
        raise HTTPException(status_code=409, detail="Шаблон ещё не подготовлен")
    storage = get_storage()
    previews = [storage.path(key) for key in (template.metadata_json or {}).get("preview_keys") or []]
    return package_response(name=template.name, filename=template.filename, sha256=template.sha256,
                            slide_count=template.slide_count, pptx_path=storage.path(template.storage_key),
                            prepared_path=storage.path(template.prepared_key), previews=previews)


def package_fonts(pptx_path: Path) -> list[tuple[Path, str]]:
    """Files of the template's typefaces that the PPTX does not embed, from the application's font store."""
    from pptx import Presentation

    deck = Presentation(str(pptx_path))
    store = font_dir().resolve()
    embedded = {_key(name) for name in embedded_families(deck)}
    members: dict[str, Path] = {}
    for family in sorted(template_families(deck)):
        if _key(family) in CORE_FONTS or _key(family) in embedded:
            continue
        for path, _ in font_files(family):
            file = Path(path).resolve()
            # System fonts stay on their machine; the store holds open fonts fetched or added for templates.
            if (file.suffix.lower() not in {".ttf", ".otf"} or not file.is_relative_to(store)
                    or file.stat().st_size > MAX_FONT_BYTES):
                continue
            relative = file.relative_to(store)
            folder = relative.parts[0] if len(relative.parts) > 1 else re.sub(r"[^A-Za-z0-9]+", "_", family).strip("_")
            name = f"fonts/{folder}/{file.name}"
            if FONT_NAME.fullmatch(name):
                members.setdefault(name, file)
    return sorted((file, name) for name, file in members.items())[:MAX_FONT_FILES]


def package_response(*, name: str, filename: str, sha256: str, slide_count: int | None, pptx_path: Path,
                     prepared_path: Path, previews: list[Path]) -> FileResponse:
    """A self-contained template package: PPTX, analysis, previews and the fonts it needs."""
    fonts = package_fonts(pptx_path)
    handle = tempfile.NamedTemporaryFile(suffix=".zip", delete=False)
    handle.close()
    package = Path(handle.name)
    manifest = {
        "format": PACKAGE_FORMAT, "version": PACKAGE_VERSION, "name": name,
        "filename": filename, "sha256": sha256, "slide_count": slide_count,
        "exported_at": datetime.now(timezone.utc).isoformat(),
        "previews": [f"previews/slide-{number}.png" for number in range(1, len(previews) + 1)],
        "fonts": [member for _, member in fonts],
    }
    try:
        with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
            archive.write(pptx_path, "template.pptx")
            archive.write(prepared_path, "template_ir.json")
            for number, preview in enumerate(previews, 1):
                archive.write(preview, f"previews/slide-{number}.png")
            for file, member in fonts:
                archive.write(file, member)
    except Exception:
        package.unlink(missing_ok=True)
        raise
    return FileResponse(package, media_type="application/zip", filename=_package_name(name),
                        background=BackgroundTask(package.unlink, missing_ok=True),
                        headers={"X-Content-Type-Options": "nosniff"})


def _read_package(archive_path: Path, target: Path) -> dict:
    """Extract only the known members of a package after bounding its size."""
    settings = get_settings()
    try:
        with zipfile.ZipFile(archive_path) as archive:
            entries = archive.infolist()
            if len(entries) > MAX_PREVIEWS + MAX_FONT_FILES + 3:
                raise ValueError("В пакете слишком много файлов")
            if sum(item.file_size for item in entries) > settings.max_archive_bytes:
                raise ValueError("Пакет шаблона слишком большой")
            names = set()
            for item in entries:
                name = item.filename
                if (item.flag_bits & 1 or PurePosixPath(name).is_absolute() or ".." in PurePosixPath(name).parts
                        or name in names):
                    raise ValueError("Пакет шаблона содержит недопустимый путь")
                if name.endswith("/"):
                    continue
                if (name not in {"manifest.json", "template.pptx", "template_ir.json"}
                        and not PREVIEW_NAME.fullmatch(name) and not FONT_NAME.fullmatch(name)):
                    raise ValueError(f"Неизвестный файл в пакете: {name[:80]}")
                if PREVIEW_NAME.fullmatch(name) and item.file_size > MAX_PREVIEW_BYTES:
                    raise ValueError("Превью в пакете слишком большое")
                if FONT_NAME.fullmatch(name) and item.file_size > MAX_FONT_BYTES:
                    raise ValueError("Шрифт в пакете слишком большой")
                names.add(name)
            if not {"manifest.json", "template.pptx", "template_ir.json"} <= names:
                raise ValueError("Это не пакет шаблона: нет manifest.json, template.pptx или template_ir.json")
            manifest = json.loads(archive.read("manifest.json").decode("utf-8"))
            if not isinstance(manifest, dict) or manifest.get("format") != PACKAGE_FORMAT:
                raise ValueError("Это не пакет шаблона Лукас")
            if manifest.get("version") != PACKAGE_VERSION:
                raise ValueError("Версия пакета шаблона не поддерживается")
            for name in names:
                destination = target / name
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(name) as source, destination.open("wb") as output:
                    shutil.copyfileobj(source, output)
    except (zipfile.BadZipFile, UnicodeDecodeError, json.JSONDecodeError, RuntimeError, NotImplementedError) as exc:
        raise ValueError("Файл пакета шаблона повреждён") from exc
    return manifest


def unpack_package(archive_path: Path, folder: Path):
    """Validated manifest and analysis of an uploaded package; its fonts go to the font store."""
    try:
        manifest = _read_package(archive_path, folder)
        validate_upload(folder / "template.pptx", ".pptx")
        prepared = verify_imported(folder / "template_ir.json", folder / "template.pptx")
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    install_package_fonts(folder)
    return manifest, prepared


def install_package_fonts(folder: Path) -> list[str]:
    """Add the package's fonts to the store so the template renders without installing them."""
    installed = []
    for file in sorted((folder / "fonts").glob("*/*")):
        with file.open("rb") as handle:
            if handle.read(4) not in FONT_MAGIC:
                continue
        target = font_dir() / file.parent.name / file.name
        if target.exists():
            continue
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(file, target)
        installed.append(f"{file.parent.name}/{file.name}")
    if installed:
        refresh_fonts()
    return installed


def package_filename(manifest: dict) -> str:
    filename = manifest.get("filename") if isinstance(manifest.get("filename"), str) else ""
    return Path(filename.replace("\\", "/")).name[:200] or "template.pptx"


def library_previews(db: Session, sha256: str) -> list[Path]:
    """Slide pictures of the same template from the library or the analysis cache."""
    storage = get_storage()
    template = db.scalar(select(Template).where(Template.sha256 == sha256, Template.preparation_status == "ready")
                         .order_by(Template.created_at))
    keys = ((template.metadata_json or {}).get("preview_keys") or []) if template is not None else []
    if keys:
        return [storage.path(key) for key in keys]
    return cached_previews(sha256)


def cached_previews(sha256: str) -> list[Path]:
    cache_root = os.getenv("MODEL_TEMPLATE_CACHE_DIR")
    cached = Path(cache_root) / "visual-references-v1" / sha256 / "slides" if cache_root else None
    if cached is None or not cached.is_dir():
        return []
    images = sorted(cached.glob("*.png"), key=lambda path: int(path.stem.rsplit("-", 1)[-1]))
    return images[:MAX_PREVIEWS]


@router.post("/import", response_model=TemplateOut, status_code=201)
def import_template(file: UploadFile = File(...), db: Session = Depends(get_db), _: User = Depends(current_user)):
    """Add an exported template without a new paid analysis."""
    if Path(file.filename or "").suffix.lower() != ".zip":
        raise HTTPException(status_code=422, detail="Выберите файл шаблона .zip, скачанный из библиотеки")
    storage = get_storage()
    staged = stage_upload(file, ".zip", storage.root, get_settings().max_upload_bytes * 2)
    new_keys: list[str] = []
    try:
        with tempfile.TemporaryDirectory(prefix="aya-import-") as scratch:
            folder = Path(scratch)
            manifest, prepared = unpack_package(staged, folder)
            existing = db.scalar(select(Template).where(Template.sha256 == prepared.template_sha256,
                                                        Template.preparation_status == "ready")
                                 .order_by(Template.created_at))
            if existing is not None:
                return template_out(existing)
            template_id = uid()
            stored = storage.put_file(folder / "template.pptx", f"templates/{template_id}/original.pptx")
            new_keys.append(stored.key)
            prepared_key = f"templates/{template_id}/prepared-{uuid4()}.json"
            # Store the normalized analysis, not the uploaded JSON.
            storage.put_file(save_prepared(prepared, folder / "normalized_ir.json"), prepared_key)
            new_keys.append(prepared_key)
            preview_keys = []
            for number in range(1, min(prepared.slide_count, MAX_PREVIEWS) + 1):
                preview = folder / "previews" / f"slide-{number}.png"
                if not preview.is_file():
                    break
                with preview.open("rb") as image:
                    if image.read(8) != b"\x89PNG\r\n\x1a\n":
                        raise HTTPException(status_code=422, detail="Превью в пакете не является PNG")
                key = f"templates/{template_id}/previews/{uuid4()}-{number}.png"
                storage.put_file(preview, key)
                new_keys.append(key)
                preview_keys.append(key)
            name = manifest.get("name") if isinstance(manifest.get("name"), str) else ""
            filename = package_filename(manifest)
            template = Template(
                id=template_id, name=(name.strip() or _display_name(filename))[:200], filename=filename,
                storage_key=stored.key, sha256=stored.sha256, size_bytes=stored.size_bytes,
                slide_count=prepared.slide_count, origin="import", preparation_status="ready",
                preparation_stage="ready", prepared_key=prepared_key, prepared_at=datetime.now(timezone.utc),
                metadata_json={"preview_keys": preview_keys, "analysis_model": prepared.analysis_model},
            )
            db.add(template)
            db.commit()
            new_keys.clear()
            db.refresh(template)
            return template_out(template)
    finally:
        staged.unlink(missing_ok=True)
        for key in new_keys:
            storage.delete(key)


def adopt_project_templates() -> None:
    """Move templates prepared inside projects into the shared library once."""
    storage = get_storage()
    with get_session_factory()() as db:
        known = set(db.scalars(select(Template.sha256)))
        sources = db.scalars(select(Source).where(Source.kind == "template", Source.preparation_status == "ready",
                                                  Source.prepared_key.is_not(None)).order_by(Source.created_at))
        for source in sources:
            if (source.sha256 in known or (source.metadata_json or {}).get("hidden_from_library")
                    or not storage.exists(source.storage_key) or not storage.exists(source.prepared_key)):
                continue
            known.add(source.sha256)
            template_id = uid()
            preview_keys = []
            for number, image in enumerate(cached_previews(source.sha256), 1):
                key = f"templates/{template_id}/previews/{uuid4()}-{number}.png"
                storage.put_file(image, key)
                preview_keys.append(key)
            try:
                slide_count = json.loads(storage.path(source.prepared_key).read_text(encoding="utf-8")).get("slide_count")
            except (OSError, ValueError):
                slide_count = None
            # Stored files are immutable, so the library entry can share them.
            db.add(Template(
                id=template_id, name=_display_name(source.filename), filename=source.filename,
                storage_key=f"templates/{template_id}/original.pptx", sha256=source.sha256,
                size_bytes=source.size_bytes, slide_count=slide_count if isinstance(slide_count, int) else None,
                origin="project", preparation_status="ready", preparation_stage="ready",
                prepared_key=source.prepared_key, prepared_at=source.prepared_at or datetime.now(timezone.utc),
                metadata_json={"preview_keys": preview_keys, "adopted_source_id": source.id},
            ))
            storage.put_file(storage.path(source.storage_key), f"templates/{template_id}/original.pptx")
        db.commit()
