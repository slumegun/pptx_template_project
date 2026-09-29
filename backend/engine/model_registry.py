"""Reviewed open weights: at most 35B total parameters, Apache-2.0 or MIT.

Adding a model requires reviewing its official weights and licence first.
The total parameter count includes inactive MoE parameters.
"""
from dataclasses import asdict, dataclass
import os

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"


@dataclass(frozen=True)
class ModelSpec:
    model_id: str
    parameters_billion: float
    license: str
    weights_url: str
    vision: bool = True
    reasoning: bool = True


APPROVED_MODELS = {
    spec.model_id: spec for spec in (
        ModelSpec("qwen/qwen3.8-27b", 27, "Apache-2.0", "https://huggingface.co/Qwen/Qwen3.8-27B"),
        ModelSpec("qwen/qwen3-vl-32b-instruct", 32, "Apache-2.0", "https://huggingface.co/Qwen/Qwen3-VL-32B-Instruct", reasoning=False),
        ModelSpec("google/gemma-4-31b-it", 31, "Apache-2.0", "https://huggingface.co/google/gemma-4-31B-it"),
    )
}

# Independent contexts even when several roles use the same model.
ROLE_DEFAULTS = {
    "template_analyst": "qwen/qwen3-vl-32b-instruct",
    "deck_planner": "qwen/qwen3.8-27b",
    "slide_worker": "qwen/qwen3.8-27b",
    "infographic_designer": "qwen/qwen3.8-27b",
    "dark_theme_designer": "qwen/qwen3.8-27b",
    "deck_critic": "google/gemma-4-31b-it",
    "prompt_critic": "google/gemma-4-31b-it",
    "visual_critic": "google/gemma-4-31b-it",
    "slide_repair": "qwen/qwen3.8-27b",
}
VISION_ROLES = {"template_analyst", "visual_critic"}


def role_model(role: str) -> ModelSpec:
    if role not in ROLE_DEFAULTS:
        raise ValueError("Unknown agent role")
    setting = "MODEL_VISUAL_CRITIC_MODEL" if role == "visual_critic" else f"MODEL_{role.upper()}"
    model_id = os.getenv(setting, ROLE_DEFAULTS[role]).strip()
    if model_id not in APPROVED_MODELS:
        raise ValueError(f"{setting} must name an approved hackathon model")
    spec = APPROVED_MODELS[model_id]
    if spec.parameters_billion > 35 or spec.license not in {"Apache-2.0", "MIT"}:
        raise ValueError("Model exceeds the hackathon size or licence restriction")
    if role in VISION_ROLES and not spec.vision:
        raise ValueError("This agent role requires image input")
    return spec


def model_manifest() -> dict:
    return {role: asdict(role_model(role)) for role in ROLE_DEFAULTS}
