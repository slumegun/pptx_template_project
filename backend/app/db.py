from collections.abc import Generator
from functools import lru_cache

from sqlalchemy import create_engine, event, inspect, update
from sqlalchemy.orm import DeclarativeBase, Session, sessionmaker

from .config import get_settings


class Base(DeclarativeBase):
    pass


@lru_cache
def get_engine():
    url = get_settings().database_url
    kwargs = {"check_same_thread": False, "timeout": 30} if url.startswith("sqlite") else {}
    engine = create_engine(url, pool_pre_ping=True, connect_args=kwargs)
    if url.startswith("sqlite"):
        @event.listens_for(engine, "connect")
        def enable_foreign_keys(connection, _):
            connection.execute("PRAGMA foreign_keys=ON")
    return engine


def locked_get(db: Session, model, object_id: str):
    """Serialize mutations on both PostgreSQL and the local SQLite backend."""
    if db.bind.dialect.name == "sqlite":
        # SQLite ignores FOR UPDATE. Acquire its write lock before reading the
        # value so a concurrent cancellation/retry cannot be overwritten.
        db.execute(update(model).where(model.id == object_id).values(id=model.id))
    return db.get(model, object_id, populate_existing=True, with_for_update=True)


@lru_cache
def get_session_factory():
    return sessionmaker(bind=get_engine(), expire_on_commit=False, class_=Session)


def get_db() -> Generator[Session, None, None]:
    with get_session_factory()() as db:
        yield db


def init_db() -> None:
    from . import models  # noqa: F401

    engine = get_engine()
    if get_settings().database_url.startswith("sqlite"):
        Base.metadata.create_all(bind=engine)
        return
    with engine.connect() as connection:
        if not inspect(connection).has_table("alembic_version"):
            raise RuntimeError("Database schema is missing; run python -m alembic upgrade head")


