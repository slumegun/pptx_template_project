from datetime import datetime, timedelta, timezone
from test_api_worker import environment, create_prepared_template, create_run
from app import jobs, db as db_module
from app.models import Run, Execution


def test_fresh_heartbeat_cannot_keep_an_overdue_run_loading_forever(environment):
    client, pipeline, tmp_path = environment
    project, source = create_prepared_template(client, pipeline, tmp_path)
    run_id = create_run(client, project, source).json()['id']
    with db_module.get_session_factory()() as db:
        run = db.get(Run, run_id)
        run.status, run.lease_token = 'running', 'overdue'
        db.add(Execution(run_id=run_id, attempt_no=1, lease_token='overdue', status='running',
                         started_at=datetime.now(timezone.utc)-timedelta(minutes=6),
                         heartbeat_at=datetime.now(timezone.utc)))
        db.commit()
    jobs.reconcile_pending()
    status = client.get(f'/api/runs/{run_id}').json()
    assert status['status'] == 'interrupted'
    assert status['error']
    with db_module.get_session_factory()() as db:
        assert db.get(Run, run_id).lease_token is None


def test_a_long_deck_gets_a_longer_budget_before_it_counts_as_overdue(environment):
    client, pipeline, tmp_path = environment
    project, source = create_prepared_template(client, pipeline, tmp_path)
    run_id = create_run(client, project, source).json()['id']
    with db_module.get_session_factory()() as db:
        run = db.get(Run, run_id)
        run.status, run.lease_token = 'running', 'long-deck'
        run.config = {**run.config, 'slide_count': 50}
        db.add(Execution(run_id=run_id, attempt_no=1, lease_token='long-deck', status='running',
                         started_at=datetime.now(timezone.utc)-timedelta(minutes=6),
                         heartbeat_at=datetime.now(timezone.utc)))
        db.commit()
    jobs.reconcile_pending()
    assert client.get(f'/api/runs/{run_id}').json()['status'] == 'running'
