import os
from functools import lru_cache
from pathlib import Path

from dotenv import dotenv_values
from pydantic import Field
from pydantic_settings import BaseSettings, SettingsConfigDict


# BaseSettings handles AYA_*; the model adapter reads MODEL_* from os.environ.
# Load only model variables here so local .env works without changing test or
# Docker database settings. Explicit process variables remain authoritative.
ENV_FILE = Path(__file__).resolve().parents[1] / ".env"
for key, value in dotenv_values(ENV_FILE).items():
    if (key.startswith(("MODEL_", "VISION_", "OPENROUTER_")) or key == "TEXT_MODEL") and value is not None:
        os.environ.setdefault(key, value)


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_prefix="AYA_", env_file=ENV_FILE, extra="ignore")

    database_url: str = "postgresql+psycopg://aya:aya@localhost:5432/aya"
    redis_url: str = "redis://localhost:6379/0"
    storage_root: Path = Path("./data")
    max_upload_bytes: int = Field(default=50 * 1024 * 1024, gt=0)
    max_archive_bytes: int = Field(default=256 * 1024 * 1024, gt=0)
    max_archive_entries: int = Field(default=10_000, gt=0)
    allowed_origins: str = "http://localhost:5173"
    inline_jobs: bool = False
    local_jobs: bool = False
    local_queue_capacity: int = Field(default=20, ge=1, le=100)


@lru_cache
def get_settings() -> Settings:
    return Settings()

