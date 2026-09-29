from engine.ingest import save_prepared
from engine.provider import configuration_status
from .schemas import SystemOut
import hashlib
import json
import logging
import tempfile
from datetime import datetime, timezone
from contextlib import asynccontextmanager
from pathlib import Path
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from sqlalchemy import select, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session
from starlette.concurrency import run_in_threadpool

from .auth import router as auth_router, current_user
from .config import get_settings
from .db import get_db, get_engine, init_db, locked_get
from .models import Artifact, Issue, Project, Run, Source, Template, User, Version, WorkerCapability, uid
from .queueing import enqueue_preparation, enqueue_run, start_local_jobs, stop_local_jobs
from .schemas import (
    ArtifactOut, IssueOut, PreparationOut, ProjectCreate, ProjectOut, RepairCreate, SlideEditCreate,
    RunCreate, RunOut, SourceOut, VersionOut,
)
from .storage import get_storage
from .library import (adopt_project_templates, library_previews, package_filename, package_response,
                      router as library_router, unpack_package)
from .uploads import stage_upload, validate_upload

logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    try:
        await run_in_threadpool(adopt_project_templates)
    except Exception:
        logger.exception("Could not move project templates into the library")
    local = get_settings().local_jobs and not get_settings().inline_jobs
    if local:
        await run_in_threadpool(start_local_jobs)
    try:
        yield
    finally:
        if local:
            await run_in_threadpool(stop_local_jobs)


app = FastAPI(title="Лукас API", version="0.1.0", lifespan=lifespan, docs_url="/api/docs", openapi_url="/api/openapi.json")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[origin.strip() for origin in get_settings().allowed_origins.split(",") if origin.strip()],
    allow_credentials=True,
    allow_methods=["GET", "POST", "DELETE"],
    allow_headers=["*"],
)

app.include_router(auth_router)
app.include_router(library_router)


def require(db: Session, model, object_id: str):
    value = db.get(model, object_id)
    if value is None:
        raise HTTPException(status_code=404, detail=f"{model.__name__} not found")
    return value



def require_project(db: Session, project_id: str, user: User) -> Project:
    project = require(db, Project, project_id)
    if project.owner_user_id != user.id:
        raise HTTPException(status_code=404, detail="Project not found")
    return project


def source_out(source: Source) -> SourceOut:
    preparation = None
    if source.kind == "template":
        preparation = PreparationOut(
            status=source.preparation_status or "pending_enqueue",
            stage=source.preparation_stage,
            error=source.preparation_error,
            prepared_at=source.prepared_at,
        )
    return SourceOut(
        id=source.id, project_id=source.project_id, kind=source.kind,
        filename=source.filename, sha256=source.sha256,
        size_bytes=source.size_bytes, created_at=source.created_at,
        preparation=preparation,
    )


def run_out(db: Session, run: Run) -> RunOut:
    version_ids = list(db.scalars(select(Version.id).where(Version.run_id == run.id).order_by(Version.ordinal)))
    artifacts = list(db.scalars(select(Artifact).where(Artifact.run_id == run.id, Artifact.version_id.is_(None))))
    config = run.config or {}
    brief = config.get("brief") if isinstance(config.get("brief"), str) else ""
    template_name = None
    if config.get("template_id"):
        library_template = db.get(Template, config["template_id"])
        template_name = library_template.name if library_template else None
    elif config.get("template_source_id"):
        source = db.get(Source, config["template_source_id"])
        template_name = source.filename if source else None
    return RunOut(
        id=run.id, project_id=run.project_id, kind=run.kind,
        parent_run_id=run.parent_run_id, base_version_id=run.base_version_id,
        status=run.status, stage=run.stage, progress=run.progress,
        error=run.error, warnings=run.warnings or [], durations=run.durations or {},
        versions=version_ids,
        artifacts=[{"id": a.id, "kind": a.kind, "url": f"/api/artifacts/{a.id}", "size_bytes": a.size_bytes, "sha256": a.sha256} for a in artifacts],
        created_at=run.created_at, updated_at=run.updated_at,
        brief_preview=" ".join(brief.split())[:160] or None,
        slide_count=config.get("slide_count") if isinstance(config.get("slide_count"), int) else None,
        template_id=config.get("template_id"), template_name=template_name,
    )


def version_out(db: Session, version: Version) -> VersionOut:
    artifacts = list(db.scalars(select(Artifact).where(Artifact.version_id == version.id).order_by(Artifact.created_at)))
    return VersionOut(
        id=version.id, project_id=version.project_id, run_id=version.run_id,
        parent_version_id=version.parent_version_id, variant_id=version.variant_id,
        ordinal=version.ordinal, quality_status=version.quality_status,
        plan=version.plan or {}, created_at=version.created_at,
        preview_urls=[f"/api/artifacts/{a.id}" for a in artifacts if a.kind == "preview"],
        artifacts=[ArtifactOut(id=a.id, kind=a.kind, url=f"/api/artifacts/{a.id}",
                               size_bytes=a.size_bytes, sha256=a.sha256) for a in artifacts],
    )


def request_digest(payload: dict) -> str:
    return hashlib.sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()


def find_idempotent_run(db: Session, project_id: str, kind: str, key: str | None, digest: str):
    if not key:
        return None
    existing = db.scalar(select(Run).where(Run.project_id == project_id, Run.kind == kind, Run.idempotency_key == key))
    if existing and existing.request_hash != digest:
        raise HTTPException(status_code=409, detail="Idempotency-Key was used with a different request")
    return existing


def schedule_run(db: Session, run: Run, force: bool = False) -> RunOut:
    try:
        enqueue_run(run.id, force=True) if force else enqueue_run(run.id)
        run = locked_get(db, Run, run.id)
        if run.status == "pending_enqueue":
            run.status = "queued"
            run.stage = "queued"
            db.commit()
    except Exception:
        logger.exception("Could not enqueue run %s; reconciliation will retry", run.id)
        db.rollback()
        run = require(db, Run, run.id)
    return run_out(db, run)


def schedule_preparation(db: Session, source: Source) -> SourceOut:
    try:
        enqueue_preparation(source.id)
        source = locked_get(db, Source, source.id)
        if source.preparation_status == "pending_enqueue":
            source.preparation_status = "queued"
            source.preparation_stage = "queued"
            db.commit()
    except Exception:
        logger.exception("Could not enqueue source %s; reconciliation will retry", source.id)
        db.rollback()
        source = require(db, Source, source.id)
    return source_out(source)


@app.get("/api/system", response_model=SystemOut)
def system_status(db: Session = Depends(get_db)):
    if get_settings().inline_jobs or get_settings().local_jobs:
        return {**configuration_status(),
                "worker_status": "inline" if get_settings().inline_jobs else "local", "reported_at": None}
    capability = db.get(WorkerCapability, "generation")
    if capability is None:
        return {
            "model_mode": "unknown",
            "model_configured": None,
            "worker_status": "unknown",
            "text_model": None,
            "vision_model": None,
            "reported_at": None,
        }
    return {
        **(capability.configuration_json or {}),
        "model_mode": capability.model_mode,
        "model_configured": capability.model_mode == "configured_api",
        "worker_status": "reported",
        "text_model": capability.text_model,
        "vision_model": capability.vision_model,
        "reported_at": capability.updated_at,
    }

@app.get("/api/health")
def health():
    try:
        with get_engine().connect() as connection:
            connection.execute(text("SELECT 1"))
    except Exception:
        raise HTTPException(status_code=503, detail="Database unavailable")
    return {"status": "ok"}


@app.post("/api/projects", response_model=ProjectOut, status_code=201)
def create_project(body: ProjectCreate, db: Session = Depends(get_db), user: User = Depends(current_user)):
    project = Project(name=body.name.strip(), owner_user_id=user.id)
    if not project.name:
        raise HTTPException(status_code=422, detail="Name cannot be blank")
    db.add(project)
    db.commit()
    db.refresh(project)
    return project


@app.get("/api/projects", response_model=list[ProjectOut])
def list_projects(db: Session = Depends(get_db), user: User = Depends(current_user)):
    return list(db.scalars(select(Project).where(Project.owner_user_id == user.id).order_by(Project.created_at.desc())))


@app.get("/api/projects/{project_id}", response_model=ProjectOut)
def get_project(project_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    return require_project(db, project_id, user)


@app.post("/api/projects/{project_id}/sources", response_model=SourceOut, status_code=201)
def upload_source(
    project_id: str,
    file: UploadFile = File(...),
    kind: str = Form(...),
    db: Session = Depends(get_db),
    user: User = Depends(current_user),
):
    require_project(db, project_id, user)
    if kind not in {"template", "content"}:
        raise HTTPException(status_code=422, detail="kind must be template or content")
    filename = Path((file.filename or "upload").replace("\\", "/")).name[:200]
    if any(ord(char) < 32 for char in filename):
        raise HTTPException(status_code=422, detail="Invalid filename")
    suffix = Path(filename).suffix.lower()
    if kind == "template" and suffix != ".pptx":
        raise HTTPException(status_code=422, detail="Template must be a .pptx file")
    if kind == "content" and suffix not in {".txt", ".md", ".csv", ".json", ".pdf", ".docx"}:
        raise HTTPException(status_code=422, detail="Unsupported content file format")
    storage = get_storage()
    staged_path = stage_upload(file, suffix, storage.root, get_settings().max_upload_bytes)
    try:
        validate_upload(staged_path, suffix)
    except ValueError as exc:
        staged_path.unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    source_id = uid()
    key = f"sources/{source_id}/original{suffix}"
    try:
        stored = storage.put_file(staged_path, key)
    finally:
        staged_path.unlink(missing_ok=True)
    source = Source(
        id=source_id, project_id=project_id, kind=kind, filename=filename,
        storage_key=stored.key, sha256=stored.sha256, size_bytes=stored.size_bytes,
        metadata_json={},
        preparation_status="pending_enqueue" if kind == "template" else None,
        preparation_stage="waiting" if kind == "template" else None,
    )
    db.add(source)
    try:
        db.commit()
    except Exception:
        db.rollback()
        storage.delete(stored.key)
        raise
    db.refresh(source)
    return schedule_preparation(db, source) if kind == "template" else source_out(source)


def hidden(source: Source) -> bool:
    """A template the user deleted from the project; presentations made with it keep working."""
    return bool((source.metadata_json or {}).get("hidden_from_library"))


@app.get("/api/projects/{project_id}/sources", response_model=list[SourceOut])
def list_sources(project_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    require_project(db, project_id, user)
    sources = db.scalars(select(Source).where(Source.project_id == project_id).order_by(Source.created_at.desc()))
    return [source_out(source) for source in sources if not hidden(source)]


@app.delete("/api/projects/{project_id}/sources/{source_id}", status_code=204)
def delete_template(project_id: str, source_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    require_project(db, project_id, user)
    source = require(db, Source, source_id)
    if source.project_id != project_id or source.kind != "template" or hidden(source):
        raise HTTPException(status_code=404, detail="Template not found in project")
    # Versions, retries and running jobs still read the file, so the template is only hidden.
    source.metadata_json = {**(source.metadata_json or {}), "hidden_from_library": True}
    db.commit()


@app.get("/api/projects/{project_id}/sources/{source_id}/export")
def export_project_template(project_id: str, source_id: str, db: Session = Depends(get_db),
                            user: User = Depends(current_user)):
    """The prepared template as a package another installation imports without a new analysis."""
    require_project(db, project_id, user)
    source = require(db, Source, source_id)
    if source.project_id != project_id or source.kind != "template" or hidden(source):
        raise HTTPException(status_code=404, detail="Template not found in project")
    if source.preparation_status != "ready" or not source.prepared_key:
        raise HTTPException(status_code=409, detail="Шаблон ещё не подготовлен")
    storage = get_storage()
    prepared_path = storage.path(source.prepared_key)
    try:
        slide_count = json.loads(prepared_path.read_text(encoding="utf-8")).get("slide_count")
    except (OSError, ValueError):
        slide_count = None
    return package_response(name=Path(source.filename).stem or "Шаблон", filename=source.filename,
                            sha256=source.sha256, slide_count=slide_count if isinstance(slide_count, int) else None,
                            pptx_path=storage.path(source.storage_key), prepared_path=prepared_path,
                            previews=library_previews(db, source.sha256))


@app.post("/api/projects/{project_id}/sources/import", response_model=SourceOut, status_code=201)
def import_project_template(project_id: str, file: UploadFile = File(...), db: Session = Depends(get_db),
                            user: User = Depends(current_user)):
    """Add an exported template package to the project, ready at once: no paid analysis."""
    require_project(db, project_id, user)
    if Path(file.filename or "").suffix.lower() != ".zip":
        raise HTTPException(status_code=422, detail="Выберите пакет шаблона .zip, скачанный кнопкой «Скачать»")
    storage = get_storage()
    staged = stage_upload(file, ".zip", storage.root, get_settings().max_upload_bytes * 2)
    new_keys: list[str] = []
    try:
        with tempfile.TemporaryDirectory(prefix="aya-import-") as scratch:
            folder = Path(scratch)
            manifest, prepared = unpack_package(staged, folder)
            for existing in db.scalars(select(Source).where(
                    Source.project_id == project_id, Source.kind == "template", Source.preparation_status == "ready",
                    Source.sha256 == prepared.template_sha256)):
                if not hidden(existing):
                    return source_out(existing)
            source_id = uid()
            stored = storage.put_file(folder / "template.pptx", f"sources/{source_id}/original.pptx")
            new_keys.append(stored.key)
            prepared_key = f"prepared/{source_id}/{uuid4()}.json"
            # Store the normalized analysis, not the uploaded JSON.
            storage.put_file(save_prepared(prepared, folder / "normalized_ir.json"), prepared_key)
            new_keys.append(prepared_key)
            source = Source(
                id=source_id, project_id=project_id, kind="template", filename=package_filename(manifest),
                storage_key=stored.key, sha256=stored.sha256, size_bytes=stored.size_bytes,
                metadata_json={"imported_package": True, "analysis_model": prepared.analysis_model},
                preparation_status="ready", preparation_stage="ready", prepared_key=prepared_key,
                prepared_at=datetime.now(timezone.utc),
            )
            db.add(source)
            db.commit()
            new_keys.clear()
            db.refresh(source)
            return source_out(source)
    finally:
        staged.unlink(missing_ok=True)
        for key in new_keys:
            storage.delete(key)


@app.get("/api/sources/{source_id}", response_model=SourceOut)
def get_source(source_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    source = require(db, Source, source_id)
    require_project(db, source.project_id, user)
    if hidden(source):
        raise HTTPException(status_code=404, detail="Source not found")
    return source_out(source)


@app.post("/api/projects/{project_id}/sources/{source_id}/prepare", response_model=SourceOut, status_code=202)
def prepare_source(project_id: str, source_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    require_project(db, project_id, user)
    source = locked_get(db, Source, source_id)
    if source is None:
        raise HTTPException(status_code=404, detail="Source not found")
    if source.project_id != project_id or hidden(source):
        raise HTTPException(status_code=404, detail="Source not found in project")
    if source.kind != "template" or (source.metadata_json or {}).get("hidden_from_library"):
        raise HTTPException(status_code=422, detail="Template is unavailable")
    if source.preparation_status in {"ready", "running", "queued"}:
        return source_out(source)
    source.preparation_status = "pending_enqueue"
    source.preparation_stage = "waiting"
    source.preparation_error = None
    db.commit()
    return schedule_preparation(db, source)


@app.post("/api/projects/{project_id}/runs", response_model=RunOut, status_code=202)
def create_run(
    project_id: str, body: RunCreate, db: Session = Depends(get_db),
    idempotency_key: str | None = Header(default=None, max_length=200, alias="Idempotency-Key"),
    user: User = Depends(current_user),
):
    require_project(db, project_id, user)
    payload = body.model_dump()
    digest = request_digest(payload)
    existing = find_idempotent_run(db, project_id, "generation", idempotency_key, digest)
    if existing:
        return run_out(db, existing)
    if body.template_id is not None:
        template = require(db, Template, body.template_id)
    else:
        template = require(db, Source, body.template_source_id)
        if template.project_id != project_id or template.kind != "template" or hidden(template):
            raise HTTPException(status_code=422, detail="Invalid template_source_id")
    if template.preparation_status != "ready" or not template.prepared_key:
        raise HTTPException(status_code=409, detail="Template preparation is not complete")
    for source_id in body.content_source_ids:
        source = require(db, Source, source_id)
        if source.project_id != project_id or source.kind != "content":
            raise HTTPException(status_code=422, detail=f"Invalid content source: {source_id}")
    run = Run(
        project_id=project_id, kind="generation", status="pending_enqueue", stage="waiting",
        config=payload, idempotency_key=idempotency_key, request_hash=digest,
    )
    db.add(run)
    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        existing = find_idempotent_run(db, project_id, "generation", idempotency_key, digest)
        if existing:
            return run_out(db, existing)
        raise
    db.refresh(run)
    return schedule_run(db, run)


@app.get("/api/projects/{project_id}/runs", response_model=list[RunOut])
def list_runs(project_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    """Generated presentations of the workspace, newest first."""
    require_project(db, project_id, user)
    runs = db.scalars(select(Run).where(Run.project_id == project_id, Run.kind == "generation")
                      .order_by(Run.created_at.desc()).limit(200))
    return [run_out(db, run) for run in runs]


@app.get("/api/runs/{run_id}", response_model=RunOut)
def get_run(run_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    run = require(db, Run, run_id)
    require_project(db, run.project_id, user)
    return run_out(db, run)


@app.post("/api/runs/{run_id}/cancel", response_model=RunOut)
def cancel_run(run_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    run = locked_get(db, Run, run_id)
    if run is None:
        raise HTTPException(status_code=404, detail="Run not found")
    require_project(db, run.project_id, user)
    if run.status in {"completed", "completed_with_warnings", "failed", "cancelled"}:
        return run_out(db, run)
    run.cancel_requested = True
    if run.status in {"pending_enqueue", "queued", "interrupted"}:
        run.status = "cancelled"
        run.stage = "cancelled"
    else:
        run.stage = "cancelling"
    db.commit()
    db.refresh(run)
    return run_out(db, run)


@app.post("/api/runs/{run_id}/retry", response_model=RunOut, status_code=202)
def retry_run(run_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    run = locked_get(db, Run, run_id)
    if run is None:
        raise HTTPException(status_code=404, detail="Run not found")
    require_project(db, run.project_id, user)
    if run.status not in {"failed", "interrupted"}:
        raise HTTPException(status_code=409, detail="Only failed or interrupted runs can be retried")
    run.status = "pending_enqueue"
    run.stage = "waiting"
    run.error = None
    run.progress = None
    run.warnings = []
    run.durations = {}
    run.cancel_requested = False
    run.lease_token = None
    db.commit()
    return schedule_run(db, run, force=True)


@app.get("/api/projects/{project_id}/versions", response_model=list[VersionOut])
def list_versions(project_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    require_project(db, project_id, user)
    versions = db.scalars(select(Version).where(Version.project_id == project_id).order_by(Version.created_at.desc()))
    return [version_out(db, version) for version in versions]


@app.get("/api/versions/{version_id}", response_model=VersionOut)
def get_version(version_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    version = require(db, Version, version_id)
    require_project(db, version.project_id, user)
    return version_out(db, version)


@app.get("/api/versions/{version_id}/issues", response_model=list[IssueOut])
def list_issues(version_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    version = require(db, Version, version_id)
    require_project(db, version.project_id, user)
    return list(db.scalars(select(Issue).where(Issue.version_id == version_id).order_by(Issue.severity, Issue.id)))


@app.post("/api/versions/{version_id}/repairs", response_model=RunOut, status_code=202)
def create_repair(
    version_id: str, body: RepairCreate, db: Session = Depends(get_db),
    idempotency_key: str | None = Header(default=None, max_length=200, alias="Idempotency-Key"),
    user: User = Depends(current_user),
):
    version = locked_get(db, Version, version_id)
    if version is None:
        raise HTTPException(status_code=404, detail="Version not found")
    require_project(db, version.project_id, user)
    issue_ids = sorted(set(body.issue_ids))
    payload = {"base_version_id": version_id, "issue_ids": issue_ids}
    digest = request_digest(payload)
    existing = find_idempotent_run(db, version.project_id, "repair", idempotency_key, digest)
    if existing:
        return run_out(db, existing)
    issues = list(db.scalars(select(Issue).where(Issue.version_id == version_id, Issue.id.in_(issue_ids))))
    if len(issues) != len(issue_ids):
        raise HTTPException(status_code=422, detail="Every selected issue must belong to the base version")
    unsupported = [issue.rule_id for issue in issues if issue.repairability != "automatic"]
    if unsupported:
        raise HTTPException(status_code=422, detail="Manual repair required: " + ", ".join(unsupported))
    active_repair = db.scalar(select(Run).where(
        Run.base_version_id == version_id,
        Run.kind == "repair",
        Run.status.in_(["pending_enqueue", "queued", "running"]),
    ))
    if active_repair is not None:
        raise HTTPException(status_code=409, detail="A repair is already running for this version")
    successor = db.scalar(select(Version.id).where(Version.parent_version_id == version.id).limit(1))
    if successor is not None:
        raise HTTPException(status_code=409, detail="A newer version of this variant exists")
    source_run = require(db, Run, version.run_id)
    run = Run(
        project_id=version.project_id, kind="repair", parent_run_id=source_run.id,
        base_version_id=version.id, status="pending_enqueue", stage="waiting",
        config={**source_run.config, "issue_ids": issue_ids},
        idempotency_key=idempotency_key, request_hash=digest,
    )
    for issue in issues:
        issue.selected = True
    db.add(run)
    try:
        db.commit()
    except IntegrityError:
        db.rollback()
        existing = find_idempotent_run(db, version.project_id, "repair", idempotency_key, digest)
        if existing:
            return run_out(db, existing)
        raise
    db.refresh(run)
    return schedule_run(db, run)


@app.post("/api/versions/{version_id}/edits", response_model=RunOut, status_code=202)
def create_slide_edit(
    version_id: str, body: SlideEditCreate, db: Session = Depends(get_db),
    user: User = Depends(current_user),
):
    version = locked_get(db, Version, version_id)
    if version is None:
        raise HTTPException(status_code=404, detail="Version not found")
    require_project(db, version.project_id, user)
    source_run = require(db, Run, version.run_id)
    slide_count = (version.plan or {}).get("metrics", {}).get("slide_count", source_run.config.get("slide_count"))
    if not isinstance(slide_count, int) or body.slide_index > slide_count:
        raise HTTPException(status_code=422, detail="Слайд не найден")
    active = db.scalar(select(Run.id).where(
        Run.base_version_id == version_id,
        Run.kind.in_(["repair", "edit"]),
        Run.status.in_(["pending_enqueue", "queued", "running"]),
    ))
    if active is not None:
        raise HTTPException(status_code=409, detail="Правки этой презентации уже выполняются")
    successor = db.scalar(select(Version.id).where(Version.parent_version_id == version_id).limit(1))
    if successor is not None:
        raise HTTPException(status_code=409, detail="Откройте последнюю версию презентации")
    run = Run(
        project_id=version.project_id, kind="edit", parent_run_id=source_run.id,
        base_version_id=version.id, status="pending_enqueue", stage="waiting",
        config={**source_run.config, "slide_index": body.slide_index, "edit_prompt": body.prompt},
    )
    db.add(run)
    db.commit()
    db.refresh(run)
    return schedule_run(db, run)


@app.get("/api/artifacts/{artifact_id}")
def download_artifact(artifact_id: str, db: Session = Depends(get_db), user: User = Depends(current_user)):
    artifact = require(db, Artifact, artifact_id)
    if artifact.version_id:
        project_id = require(db, Version, artifact.version_id).project_id
    elif artifact.run_id:
        project_id = require(db, Run, artifact.run_id).project_id
    else:
        raise HTTPException(status_code=404, detail="Artifact not found")
    require_project(db, project_id, user)
    path = get_storage().path(artifact.storage_key)
    if not path.is_file():
        raise HTTPException(status_code=404, detail="Artifact file is missing")
    extension = {"pptx": ".pptx", "pdf": ".pdf", "html": ".html", "preview": ".png", "audit": ".json"}.get(artifact.kind, "")
    media_type = {"pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation", "pdf": "application/pdf",
                  "html": "text/html", "preview": "image/png", "audit": "application/json"}.get(artifact.kind, "application/octet-stream")
    return FileResponse(
        path, filename=f"aya-{artifact.id}{extension}", media_type=media_type,
        content_disposition_type="inline" if artifact.kind == "preview" else "attachment",
        headers={"X-Content-Type-Options": "nosniff", "Content-Security-Policy": "sandbox; default-src 'none'; img-src data:; style-src 'unsafe-inline'"},
    )










