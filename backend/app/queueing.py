import logging
from concurrent.futures import ThreadPoolExecutor
from threading import Event, RLock, Thread

from redis import Redis
from rq import Queue

from .config import get_settings

logger = logging.getLogger(__name__)
_executor: ThreadPoolExecutor | None = None
_pending = {}
_lock = RLock()
_stopped = Event()
_reconciler: Thread | None = None


def _submit_local(job_id, function, argument):
    global _executor
    with _lock:
        if job_id in _pending:
            return
        if len(_pending) >= get_settings().local_queue_capacity:
            raise RuntimeError("Local queue is full; reconciliation will retry")
        if _executor is None:
            _executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="aya-worker")
        future = _executor.submit(function, argument)
        _pending[job_id] = future

        def finished(result):
            with _lock:
                _pending.pop(job_id, None)
            if not result.cancelled() and result.exception() is not None:
                logger.error("Local job failed: %s", job_id)

        future.add_done_callback(finished)


def start_local_jobs():
    global _reconciler
    from .jobs import publish_capability, reconcile_pending

    publish_capability()
    _stopped.clear()

    def reconcile_loop():
        while not _stopped.is_set():
            try:
                reconcile_pending()
            except Exception:
                logger.exception("Local job reconciliation failed")
            _stopped.wait(30)

    _reconciler = Thread(target=reconcile_loop, daemon=True, name="aya-reconciler")
    _reconciler.start()


def stop_local_jobs():
    global _executor, _reconciler
    _stopped.set()
    if _reconciler is not None:
        _reconciler.join(timeout=35)
        _reconciler = None
    if _executor is not None:
        _executor.shutdown(wait=True)
        _executor = None


def get_queue() -> Queue:
    return Queue("aya", connection=Redis.from_url(get_settings().redis_url), default_timeout=1800)


def enqueue_run(run_id: str, force: bool = False) -> None:
    if get_settings().inline_jobs:
        from .jobs import execute_run
        execute_run(run_id)
        return
    if get_settings().local_jobs:
        from .jobs import execute_run
        _submit_local(run_id, execute_run, run_id)
        return
    queue = get_queue()
    job = queue.fetch_job(run_id)
    if job:
        if not force and job.get_status(refresh=True) in {"queued", "started", "deferred", "scheduled"}:
            return
        job.delete()
    queue.enqueue("app.jobs.execute_run", run_id, job_id=run_id, job_timeout=1800, result_ttl=3600, failure_ttl=86400)


def enqueue_template(template_id: str, force: bool = False) -> None:
    if get_settings().inline_jobs:
        from .jobs import prepare_template
        prepare_template(template_id)
        return
    if get_settings().local_jobs:
        from .jobs import prepare_template
        _submit_local(f"template-{template_id}", prepare_template, template_id)
        return
    queue = get_queue()
    job_id = f"template-{template_id}"
    job = queue.fetch_job(job_id)
    if job:
        if not force and job.get_status(refresh=True) in {"queued", "started", "deferred", "scheduled"}:
            return
        job.delete()
    queue.enqueue("app.jobs.prepare_template", template_id, job_id=job_id, job_timeout=1800, result_ttl=3600, failure_ttl=86400)


def enqueue_preparation(source_id: str, force: bool = False) -> None:
    if get_settings().inline_jobs:
        from .jobs import prepare_source
        prepare_source(source_id)
        return
    if get_settings().local_jobs:
        from .jobs import prepare_source
        _submit_local(f"prepare-{source_id}", prepare_source, source_id)
        return
    queue = get_queue()
    job_id = f"prepare-{source_id}"
    job = queue.fetch_job(job_id)
    if job:
        if not force and job.get_status(refresh=True) in {"queued", "started", "deferred", "scheduled"}:
            return
        job.delete()
    queue.enqueue("app.jobs.prepare_source", source_id, job_id=job_id, job_timeout=900, result_ttl=3600, failure_ttl=86400)



