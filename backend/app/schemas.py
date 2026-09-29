from datetime import datetime
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class ProjectCreate(BaseModel):
    name: str = Field(min_length=1, max_length=160)


class ProjectOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: str
    name: str
    created_at: datetime


class PreparationOut(BaseModel):
    status: str
    stage: str | None = None
    error: str | None = None
    prepared_at: datetime | None = None


class SourceOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: str
    project_id: str
    kind: str
    filename: str
    sha256: str
    size_bytes: int
    created_at: datetime
    preparation: PreparationOut | None = None


class TemplateOut(BaseModel):
    id: str
    name: str
    filename: str
    sha256: str
    size_bytes: int
    slide_count: int | None = None
    origin: str
    created_at: datetime
    preparation: PreparationOut
    preview_urls: list[str] = Field(default_factory=list)
    export_url: str | None = None


class RunCreate(BaseModel):
    brief: str = Field(min_length=1, max_length=100_000)
    slide_count: int = Field(default=10, ge=1, le=60)
    # A library template (template_id) or a legacy project upload (template_source_id).
    template_id: str | None = Field(default=None, min_length=1, max_length=36)
    template_source_id: str | None = Field(default=None, min_length=1, max_length=36)
    content_source_ids: list[str] = Field(default_factory=list, max_length=20)

    @model_validator(mode="after")
    def one_template(self):
        if (self.template_id is None) == (self.template_source_id is None):
            raise ValueError("Choose exactly one template")
        return self

    @field_validator("brief")
    @classmethod
    def nonblank_brief(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("Brief cannot be blank")
        return value

    @field_validator("content_source_ids")
    @classmethod
    def unique_sources(cls, values: list[str]) -> list[str]:
        if any(not value or len(value) > 36 for value in values):
            raise ValueError("Invalid content source identifier")
        return list(dict.fromkeys(values))


class RepairCreate(BaseModel):
    issue_ids: list[str] = Field(min_length=1, max_length=500)


class SlideEditCreate(BaseModel):
    slide_index: int = Field(ge=1)
    prompt: str = Field(min_length=3, max_length=2000)

    @field_validator("prompt")
    @classmethod
    def nonblank_prompt(cls, value: str) -> str:
        value = value.strip()
        if len(value) < 3:
            raise ValueError("Опишите правку для слайда")
        return value


class RunOut(BaseModel):
    id: str
    project_id: str
    kind: str
    parent_run_id: str | None
    base_version_id: str | None
    status: str
    stage: str
    progress: int | None
    error: str | None
    warnings: list[str]
    durations: dict[str, Any]
    versions: list[str]
    artifacts: list[dict[str, Any]]
    created_at: datetime
    updated_at: datetime
    brief_preview: str | None = None
    slide_count: int | None = None
    template_id: str | None = None
    template_name: str | None = None


class ArtifactOut(BaseModel):
    id: str
    kind: str
    url: str
    size_bytes: int
    sha256: str


class VersionOut(BaseModel):
    id: str
    project_id: str
    run_id: str
    parent_version_id: str | None
    variant_id: str
    ordinal: int
    quality_status: str
    plan: dict[str, Any]
    created_at: datetime
    preview_urls: list[str]
    artifacts: list[ArtifactOut]


class IssueOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: str
    version_id: str
    slide_id: str | None
    object_id: str | None
    rule_id: str
    severity: str
    repairability: str
    evidence: dict[str, Any]
    selected: bool
    resolution: str | None


class ErrorOut(BaseModel):
    detail: str







class SystemOut(BaseModel):
    """Additive frontend contract; configuration does not prove live API balance."""
    model_mode: str
    model_configured: bool | None = None
    worker_status: str
    text_model: str | None = None
    vision_model: str | None = None
    reported_at: datetime | None = None
    provider: str | None = None
    endpoint_host: str | None = None
    configuration_status: str = "unknown"
    configuration_error: str | None = None
    agents: list[dict[str, Any]] = Field(default_factory=list)
    model_manifest: dict[str, Any] = Field(default_factory=dict)
    features: dict[str, Any] = Field(default_factory=dict)
    live_verified: bool = False
