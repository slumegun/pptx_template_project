import json
import logging
import os
import shutil
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from pathlib import Path
from threading import Event, Thread
from uuid import uuid4

from engine.provider import configuration_status

from redis import Redis
from rq import Worker
from sqlalchemy import func, select, or_

from .config import get_settings
from .db import get_session_factory, init_db, locked_get
from .models import Artifact, Execution, Issue, Run, Source, Template, Version, WorkerCapability
from .queueing import enqueue_preparation, enqueue_run, enqueue_template, get_queue
from .storage import get_storage

logger = logging.getLogger(__name__)


class RunCancelled(Exception):
    pass


class StaleExecution(Exception):
    pass


def publish_capability() -> None:
    """Worker reports model availability without storing or sharing its API key."""
    init_db()
    profile = configuration_status()
    with get_session_factory()() as db:
        capability = db.get(WorkerCapability, "generation", with_for_update=True)
        if capability is None:
            capability = WorkerCapability(id="generation")
            db.add(capability)
        capability.model_mode = profile["model_mode"]
        capability.text_model = profile["text_model"]
        capability.vision_model = profile["vision_model"]
        capability.configuration_json = profile
        capability.updated_at = datetime.now(timezone.utc)
        db.commit()

def prepare_source(source_id: str) -> None:
    """Analyze one uploaded template outside the measured generation run."""
    init_db()
    sessions = get_session_factory()
    lease = str(uuid4())
    with sessions() as db:
        source = locked_get(db, Source, source_id)
        if source is None or source.kind != "template" or source.preparation_status not in {"pending_enqueue", "queued"}:
            return
        source.preparation_status = "running"
        source.preparation_stage = "analyzing"
        source.preparation_error = None
        source.metadata_json = {**(source.metadata_json or {}), "preparation_lease": lease,
                                "preparation_started_at": datetime.now(timezone.utc).isoformat(),
                                "preparation_heartbeat_at": datetime.now(timezone.utc).isoformat()}
        key = source.storage_key
        db.commit()
    started = time.perf_counter()
    storage = get_storage()
    prepared_key = None
    committed = False
    try:
        from engine.pipeline import prepare

        with preparation_heartbeat(source_id, lease), tempfile.TemporaryDirectory(prefix=f"aya-prepare-{source_id}-") as scratch:
            prepared_path = Path(prepare(storage.path(key), Path(scratch)))
            if not prepared_path.resolve().is_relative_to(Path(scratch).resolve()) or not prepared_path.is_file():
                raise RuntimeError("Template analysis did not produce a file")
            prepared_key = f"prepared/{source_id}/{uuid4()}.json"
            storage.put_file(prepared_path, prepared_key)
        with sessions() as db:
            source = locked_get(db, Source, source_id)
            if source is None or (source.metadata_json or {}).get("preparation_lease") != lease:
                raise StaleExecution()
            source.prepared_key = prepared_key
            source.preparation_status = "ready"
            source.preparation_stage = "ready"
            source.prepared_at = datetime.now(timezone.utc)
            source.metadata_json = {**(source.metadata_json or {}), "analysis_seconds": round(time.perf_counter() - started, 3)}
            db.commit()
            committed = True
    except Exception as exc:
        logger.exception("Template preparation failed: %s", source_id)
        with sessions() as db:
            source = locked_get(db, Source, source_id)
            if source is not None and (source.metadata_json or {}).get("preparation_lease") == lease:
                source.preparation_status = "failed"
                source.preparation_stage = "failed"
                source.preparation_error = public_error(exc)
                db.commit()
        if not isinstance(exc, StaleExecution):
            raise
    finally:
        if prepared_key and not committed:
            storage.delete(prepared_key)


@contextmanager
def preparation_heartbeat(source_id: str, lease: str, model=Source):
    """Refresh a running analysis (a project Source or a library Template) every 30 s."""
    stopped = Event()

    def heartbeat():
        while not stopped.wait(30):
            try:
                with get_session_factory()() as db:
                    source = locked_get(db, model, source_id)
                    if source is None or source.preparation_status != "running" or (source.metadata_json or {}).get("preparation_lease") != lease:
                        return
                    source.metadata_json = {**source.metadata_json, "preparation_heartbeat_at": datetime.now(timezone.utc).isoformat()}
                    db.commit()
            except Exception:
                logger.exception("Could not refresh preparation heartbeat for %s", source_id)

    thread = Thread(target=heartbeat, daemon=True)
    thread.start()
    try:
        yield
    finally:
        stopped.set()
        thread.join(timeout=35)


MAX_TEMPLATE_PREVIEWS = 80


def prepare_template(template_id: str) -> None:
    """Analyze a library template once; its result is reusable by every account."""
    init_db()
    sessions = get_session_factory()
    lease = str(uuid4())
    with sessions() as db:
        template = locked_get(db, Template, template_id)
        if template is None or template.preparation_status not in {"pending_enqueue", "queued"}:
            return
        template.preparation_status = "running"
        template.preparation_stage = "analyzing"
        template.preparation_error = None
        template.metadata_json = {**(template.metadata_json or {}), "preparation_lease": lease,
                                  "preparation_started_at": datetime.now(timezone.utc).isoformat(),
                                  "preparation_heartbeat_at": datetime.now(timezone.utc).isoformat()}
        key = template.storage_key
        db.commit()
    started = time.perf_counter()
    storage = get_storage()
    new_keys: list[str] = []
    committed = False
    try:
        from engine.pipeline import prepare, template_previews

        with (preparation_heartbeat(template_id, lease, Template),
              tempfile.TemporaryDirectory(prefix=f"aya-template-{template_id}-") as scratch):
            prepared_path = Path(prepare(storage.path(key), Path(scratch)))
            if not prepared_path.resolve().is_relative_to(Path(scratch).resolve()) or not prepared_path.is_file():
                raise RuntimeError("Template analysis did not produce a file")
            prepared_key = f"templates/{template_id}/prepared-{uuid4()}.json"
            storage.put_file(prepared_path, prepared_key)
            new_keys.append(prepared_key)
            preview_keys = []
            for number, preview in enumerate(template_previews(storage.path(key), prepared_path, Path(scratch))[:MAX_TEMPLATE_PREVIEWS], 1):
                preview_key = f"templates/{template_id}/previews/{uuid4()}-{number}.png"
                storage.put_file(Path(preview), preview_key)
                new_keys.append(preview_key)
                preview_keys.append(preview_key)
            slide_count = json.loads(prepared_path.read_text(encoding="utf-8")).get("slide_count")
        with sessions() as db:
            template = locked_get(db, Template, template_id)
            if template is None or (template.metadata_json or {}).get("preparation_lease") != lease:
                raise StaleExecution()
            template.prepared_key = prepared_key
            template.slide_count = slide_count if isinstance(slide_count, int) else None
            template.preparation_status = "ready"
            template.preparation_stage = "ready"
            template.prepared_at = datetime.now(timezone.utc)
            template.metadata_json = {**(template.metadata_json or {}), "preview_keys": preview_keys,
                                      "analysis_seconds": round(time.perf_counter() - started, 3)}
            db.commit()
            committed = True
    except Exception as exc:
        logger.exception("Template library preparation failed: %s", template_id)
        with sessions() as db:
            template = locked_get(db, Template, template_id)
            if template is not None and (template.metadata_json or {}).get("preparation_lease") == lease:
                template.preparation_status = "failed"
                template.preparation_stage = "failed"
                template.preparation_error = public_error(exc)
                db.commit()
        if not isinstance(exc, StaleExecution):
            raise
    finally:
        if not committed:
            for new_key in new_keys:
                storage.delete(new_key)


def public_error(exc: Exception) -> str:
    message = str(exc)
    for name, value in os.environ.items():
        if (name.endswith(("API_KEY", "TOKEN", "PASSWORD")) and len(value) >= 6):
            message = message.replace(value, "[redacted]")
    return message[:1000]


@contextmanager
def execution_heartbeat(run_id: str, lease: str):
    stopped = Event()

    def heartbeat():
        while not stopped.wait(30):
            try:
                with get_session_factory()() as db:
                    run = locked_get(db, Run, run_id)
                    if run is None or run.lease_token != lease or run.status != "running":
                        return
                    execution = db.scalar(select(Execution).where(Execution.run_id == run_id, Execution.lease_token == lease))
                    if execution:
                        execution.heartbeat_at = datetime.now(timezone.utc)
                    db.commit()
            except Exception:
                logger.exception("Could not refresh heartbeat for %s", run_id)

    thread = Thread(target=heartbeat, daemon=True)
    thread.start()
    try:
        yield
    finally:
        stopped.set()
        thread.join(timeout=35)


def _update_progress(run_id: str, lease: str, stage: str, percent: int | None) -> None:
    with get_session_factory()() as db:
        run = locked_get(db, Run, run_id)
        if run is None or run.lease_token != lease:
            raise StaleExecution()
        if run.cancel_requested:
            raise RunCancelled()
        run.stage = stage[:64]
        run.progress = max(0, min(100, int(percent))) if percent is not None else None
        execution = db.scalar(select(Execution).where(Execution.run_id == run_id, Execution.lease_token == lease))
        if execution:
            execution.heartbeat_at = datetime.now(timezone.utc)
        db.commit()


def _artifact_from_file(storage, source: Path, kind: str, version_id: str | None, run_id: str):
    if not source.is_file():
        raise RuntimeError(f"Missing {kind} output: {source}")
    artifact_id = str(uuid4())
    suffix = source.suffix.lower() or ".bin"
    stored = storage.put_file(source, f"artifacts/{artifact_id}/file{suffix}")
    return Artifact(
        id=artifact_id, version_id=version_id, run_id=run_id, kind=kind,
        storage_key=stored.key, sha256=stored.sha256, size_bytes=stored.size_bytes,
    )


def _quality(issues: list[dict], missing_exports: bool) -> str:
    if missing_exports or any(str(i.get("severity", "")).lower() in {"blocking", "critical", "error"} for i in issues):
        return "blocked"
    if issues:
        return "warnings"
    return "passed"


def _serialize_issue(version_id: str, raw: dict) -> Issue:
    evidence = raw.get("evidence") or {}
    if not isinstance(evidence, dict):
        evidence = {"message": str(evidence)}
    if raw.get("message") and "message" not in evidence:
        evidence["message"] = raw["message"]
    return Issue(
        version_id=version_id,
        slide_id=str(raw["slide_id"]) if raw.get("slide_id") is not None else (
            str(raw["slide_index"]) if raw.get("slide_index") is not None else None
        ),
        object_id=str(raw["object_id"]) if raw.get("object_id") is not None else None,
        rule_id=str(raw.get("rule_id") or raw.get("code") or "unspecified"),
        severity=str(raw.get("severity") or "warning"),
        repairability=str(raw.get("repairability") or "manual"),
        engine_issue_id=str(raw["issue_id"]) if raw.get("issue_id") else None,
        evidence=evidence,
    )


def execute_run(run_id: str) -> None:
    """Execute an RQ job. A lease prevents a late worker from publishing stale output."""
    init_db()
    sessions = get_session_factory()
    lease = str(uuid4())
    started = time.perf_counter()
    with sessions() as db:
        run = locked_get(db, Run, run_id)
        if run is None or run.status not in {"pending_enqueue", "queued"}:
            return
        if run.cancel_requested:
            run.status = "cancelled"
            run.stage = "cancelled"
            db.commit()
            return
        attempt_no = (db.scalar(select(func.max(Execution.attempt_no)).where(Execution.run_id == run_id)) or 0) + 1
        execution = Execution(run_id=run_id, attempt_no=attempt_no, lease_token=lease)
        db.add(execution)
        run.lease_token = lease
        run.status = "running"
        run.stage = "starting"
        run.progress = None
        config = dict(run.config)
        kind = run.kind
        base_version_id = run.base_version_id
        project_id = run.project_id
        db.commit()

    storage = get_storage()
    published_keys: list[str] = []
    publication_committed = False
    try:
        with sessions() as db:
            if config.get("template_id"):
                template = db.get(Template, config["template_id"])
                if template is None or template.preparation_status != "ready" or not template.prepared_key:
                    raise RuntimeError("Template preparation is unavailable")
            else:
                template = db.get(Source, config["template_source_id"])
                if (template is None or template.project_id != project_id or template.kind != "template"
                        or template.preparation_status != "ready" or not template.prepared_key):
                    raise RuntimeError("Template preparation is unavailable")
            template_path = storage.path(template.storage_key)
            prepared_path = storage.path(template.prepared_key)
            analysis_seconds = (template.metadata_json or {}).get("analysis_seconds")
            content_sources = []
            for source_id in config.get("content_source_ids", []):
                source = db.get(Source, source_id)
                if source is None or source.project_id != project_id or source.kind != "content":
                    raise RuntimeError(f"Content source is unavailable: {source_id}")
                content_sources.append((source.id, source.filename, storage.path(source.storage_key)))
            base = db.get(Version, base_version_id) if base_version_id else None
            selected_issues = []
            base_pptx_path = None
            if kind in {"repair", "edit"}:
                if base is None:
                    raise RuntimeError("Base version is unavailable")
                base_pptx = db.scalar(select(Artifact).where(
                    Artifact.version_id == base.id, Artifact.kind == "pptx"
                ))
                if base_pptx is None:
                    raise RuntimeError("Base PPTX is unavailable")
                base_pptx_path = storage.path(base_pptx.storage_key)
                if kind == "repair":
                    issue_ids = config.get("issue_ids", [])
                    issues = list(db.scalars(select(Issue).where(Issue.version_id == base.id, Issue.id.in_(issue_ids))))
                    if len(issues) != len(issue_ids):
                        raise RuntimeError("Selected issues no longer match the base version")
                    selected_issues = [{
                        "issue_id": i.engine_issue_id or i.id, "slide_id": i.slide_id, "object_id": i.object_id,
                        "repairability": i.repairability,
                        "rule_id": i.rule_id, "severity": i.severity, "evidence": i.evidence,
                    } for i in issues]

        def progress(stage: str, percent: int | None = None):
            _update_progress(run_id, lease, stage, percent)

        from engine.pipeline import generate
        with execution_heartbeat(run_id, lease), tempfile.TemporaryDirectory(prefix=f"aya-run-{run_id}-") as scratch:
            output_dir = Path(scratch)
            content_dir = output_dir / "content"
            content_dir.mkdir()
            content_paths = []
            for source_id, filename, stored_path in content_sources:
                safe_name = Path(filename.replace("\\", "/")).name
                target = content_dir / f"{source_id}_{safe_name}"
                shutil.copyfile(stored_path, target)
                content_paths.append(target)
            _update_progress(run_id, lease, "generating" if kind == "generation" else "repairing", None)
            engine_started = time.perf_counter()
            if kind == "repair":
                try:
                    from engine.pipeline import repair
                except ImportError as exc:
                    raise RuntimeError("Selected repair is not implemented by the engine") from exc
                variants = repair(
                    template_path, base_pptx_path, config["brief"], config["slide_count"],
                    selected_issues, output_dir, progress=progress, content_paths=content_paths,
                )
            elif kind == "edit":
                from engine.pipeline import edit_slide
                variants = edit_slide(
                    template_path, base_pptx_path, config["brief"], config["slide_count"],
                    config["slide_index"], config["edit_prompt"], output_dir,
                    progress=progress, content_paths=content_paths,
                )
            else:
                variants = generate(
                    template_path, config["brief"], config["slide_count"], output_dir,
                    progress=progress, content_paths=content_paths, prepared_path=prepared_path,
                )
            engine_seconds = round(time.perf_counter() - engine_started, 3)
            if not variants:
                raise RuntimeError("Engine produced no variants")
            if kind == "generation" and len(variants) != 3:
                raise RuntimeError(f"Engine produced {len(variants)} variants; expected 3")
            if kind in {"repair", "edit"} and len(variants) != 1:
                raise RuntimeError("Repair must produce exactly one version")
            if len({str(item.variant_id) for item in variants}) != len(variants):
                raise RuntimeError("Engine returned duplicate variant identifiers")
            _update_progress(run_id, lease, "publishing", None)
            publication_started = time.perf_counter()
            versions_to_add = []
            artifacts_to_add = []
            issues_to_add = []
            warnings: list[str] = []
            for ordinal, result in enumerate(variants, start=1):
                version_id = str(uuid4())
                raw_issues = list(result.issues or [])
                missing_exports = result.pdf_path is None or result.html_path is None
                if result.pdf_path is None:
                    warnings.append(f"Variant {result.variant_id}: PDF export unavailable")
                if result.html_path is None:
                    warnings.append(f"Variant {result.variant_id}: HTML export unavailable")
                version = Version(
                    id=version_id, project_id=project_id, run_id=run_id,
                    parent_version_id=base_version_id,
                    variant_id=base.variant_id if base is not None else str(result.variant_id),
                    ordinal=(base.ordinal + 1 if base is not None else 1),
                    quality_status=_quality(raw_issues, missing_exports),
                    plan={"metrics": result.metrics or {}},
                )
                versions_to_add.append(version)
                paths = [("pptx", result.pptx_path), ("pdf", result.pdf_path), ("html", result.html_path)]
                paths += [("preview", path) for path in (result.preview_paths or [])]
                for artifact_kind, path in paths:
                    if path is None:
                        continue
                    if not Path(path).resolve().is_relative_to(output_dir.resolve()):
                        raise RuntimeError("Engine output is outside the execution directory")
                    artifact = _artifact_from_file(storage, Path(path), artifact_kind, version_id, run_id)
                    published_keys.append(artifact.storage_key)
                    artifacts_to_add.append(artifact)
                issues_to_add.extend(_serialize_issue(version_id, issue) for issue in raw_issues)
            report_path = output_dir / "run_report.json"
            if not report_path.is_file():
                report_path.write_text(json.dumps({
                    "run_id": run_id,
                    "kind": kind,
                    "base_version_id": base_version_id,
                    "variants": [
                        {"variant_id": item.variant_id, "metrics": item.metrics or {}, "issues": item.issues or []}
                        for item in variants
                    ],
                }, ensure_ascii=False, indent=2), encoding="utf-8")
            try:
                report_warnings = json.loads(report_path.read_text(encoding="utf-8")).get("warnings", [])
                if isinstance(report_warnings, list):
                    warnings.extend(str(item)[:500] for item in report_warnings if isinstance(item, str) and item.strip())
            except (OSError, ValueError, TypeError):
                logger.warning("Could not read engine warnings from %s", report_path)
            report_artifact = _artifact_from_file(storage, report_path, "audit", None, run_id)
            published_keys.append(report_artifact.storage_key)
            artifacts_to_add.append(report_artifact)
            publication_seconds = round(time.perf_counter() - publication_started, 3)
            with sessions() as db:
                run = locked_get(db, Run, run_id)
                if run is None or run.lease_token != lease:
                    raise StaleExecution()
                if run.cancel_requested:
                    raise RunCancelled()
                db.add_all(versions_to_add)
                db.flush()
                db.add_all(artifacts_to_add + issues_to_add)
                run.status = "completed_with_warnings" if warnings or any(v.quality_status != "passed" for v in versions_to_add) else "completed"
                run.stage = "completed"
                run.progress = 100
                run.warnings = warnings
                created_at = run.created_at if run.created_at.tzinfo else run.created_at.replace(tzinfo=timezone.utc)
                request_to_result_seconds = round((datetime.now(timezone.utc) - created_at).total_seconds(), 3)
                run.durations = {
                    "analysis_seconds": analysis_seconds,
                    "generation_seconds": engine_seconds,
                    "publication_seconds": publication_seconds,
                    "total_run_seconds": round(time.perf_counter() - started, 3),
                    "request_to_result_seconds": request_to_result_seconds,
                    "within_300_seconds": request_to_result_seconds <= 300,
                }
                execution = db.scalar(select(Execution).where(Execution.run_id == run_id, Execution.lease_token == lease))
                if execution:
                    execution.status = "completed"
                    execution.finished_at = datetime.now(timezone.utc)
                db.commit()
                publication_committed = True
    except Exception as exc:
        if publication_committed:
            logger.exception("Run %s completed, but temporary file cleanup failed", run_id)
            return
        for key in published_keys:
            try:
                storage.delete(key)
            except Exception:
                logger.exception("Could not remove orphaned artifact %s", key)
        cancelled = isinstance(exc, RunCancelled)
        stale = isinstance(exc, StaleExecution)
        with sessions() as db:
            run = locked_get(db, Run, run_id)
            if run is not None and run.lease_token == lease:
                run.status = "cancelled" if cancelled else "failed"
                run.stage = "cancelled" if cancelled else "failed"
                run.error = None if cancelled else public_error(exc)
                run.durations = {**(run.durations or {}), "total_run_seconds": round(time.perf_counter() - started, 3)}
                execution = db.scalar(select(Execution).where(Execution.run_id == run_id, Execution.lease_token == lease))
                if execution:
                    execution.status = "cancelled" if cancelled else "failed"
                    execution.finished_at = datetime.now(timezone.utc)
                db.commit()
        if not cancelled and not stale:
            logger.exception("Run failed: %s", run_id)
            raise


def reconcile_pending() -> None:
    """Recover the transaction/enqueue gap after a Redis or API interruption."""
    init_db()
    with get_session_factory()() as db:
        from engine.models import generation_budget_seconds

        now = datetime.now(timezone.utc)
        cutoff = now - timedelta(minutes=10)

        def aware(value):
            return value.replace(tzinfo=timezone.utc) if value is not None and value.tzinfo is None else value

        stale_executions = []
        for execution in db.scalars(select(Execution).where(Execution.status == "running")):
            run = db.get(Run, execution.run_id)
            count = (run.config or {}).get("slide_count") if run is not None else None
            # The generation budget grows with the deck; 40 s more cover publication.
            limit = timedelta(seconds=generation_budget_seconds(count if isinstance(count, int) else 10) + 40)
            heartbeat = aware(execution.heartbeat_at)
            if (heartbeat is not None and heartbeat < cutoff) or aware(execution.started_at) < now - limit:
                stale_executions.append(execution)
        for execution in stale_executions:
            current = db.get(Run, execution.run_id, populate_existing=True, with_for_update=True)
            if current is None or current.status != "running" or current.lease_token != execution.lease_token:
                continue
            current.status = "interrupted"
            current.stage = "interrupted"
            current.error = "Обработчик прерван или превышен лимит выполнения. Можно повторить генерацию."
            current.lease_token = None
            execution.status = "interrupted"
            execution.finished_at = datetime.now(timezone.utc)
            db.commit()
        # A stopped worker must not leave a template permanently "running".
        preparing = [(Source, source.id) for source in db.scalars(select(Source).where(
            Source.kind == "template", Source.preparation_status == "running"))]
        preparing += [(Template, template.id) for template in db.scalars(select(Template).where(
            Template.preparation_status == "running"))]
        for model, record_id in preparing:
            current = locked_get(db, model, record_id)
            if current is None or current.preparation_status != "running":
                continue
            metadata = current.metadata_json or {}
            timestamp = metadata.get("preparation_heartbeat_at") or metadata.get("preparation_started_at")
            try:
                heartbeat = datetime.fromisoformat(timestamp)
                if heartbeat.tzinfo is None:
                    heartbeat = heartbeat.replace(tzinfo=timezone.utc)
            except (TypeError, ValueError):
                heartbeat = current.created_at.replace(tzinfo=timezone.utc)
            if heartbeat < cutoff:
                current.preparation_status = "failed"
                current.preparation_stage = "failed"
                current.preparation_error = "Подготовка прервана из-за остановки обработчика. Нажмите «Повторить»."
                current.metadata_json = {**metadata, "preparation_lease": None}
                db.commit()
        sources = list(db.scalars(select(Source).where(Source.kind == "template", Source.preparation_status.in_(["pending_enqueue", "queued"]))))
        runs = list(db.scalars(select(Run).where(Run.status.in_(["pending_enqueue", "queued"]), Run.cancel_requested.is_(False))))
        for source in sources:
            try:
                enqueue_preparation(source.id)
                current = db.get(Source, source.id, populate_existing=True, with_for_update=True)
                if current is not None and current.preparation_status == "pending_enqueue":
                    current.preparation_status = "queued"
                    current.preparation_stage = "queued"
                    db.commit()
            except Exception:
                db.rollback()
                logger.exception("Could not reconcile source %s", source.id)
        templates = list(db.scalars(select(Template).where(Template.preparation_status.in_(["pending_enqueue", "queued"]))))
        for template in templates:
            try:
                enqueue_template(template.id)
                current = db.get(Template, template.id, populate_existing=True, with_for_update=True)
                if current is not None and current.preparation_status == "pending_enqueue":
                    current.preparation_status = "queued"
                    current.preparation_stage = "queued"
                    db.commit()
            except Exception:
                db.rollback()
                logger.exception("Could not reconcile template %s", template.id)
        for run in runs:
            try:
                enqueue_run(run.id)
                current = db.get(Run, run.id, populate_existing=True, with_for_update=True)
                if current is not None and current.status == "pending_enqueue":
                    current.status = "queued"
                    current.stage = "queued"
                    db.commit()
            except Exception:
                db.rollback()
                logger.exception("Could not reconcile run %s", run.id)


def main() -> None:
    logging.basicConfig(level=logging.INFO)
    publish_capability()
    reconcile_pending()
    queue = get_queue()
    worker = Worker([queue], connection=Redis.from_url(get_settings().redis_url))
    worker.work()


if __name__ == "__main__":
    main()










