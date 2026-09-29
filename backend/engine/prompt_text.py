"""Versioned prompts shipped with the engine package."""
from functools import lru_cache
from pathlib import Path

PROMPT_VERSION = "2026-09-29.1"

@lru_cache
def prompt_text(role: str) -> str:
    if role not in {"template_analyst", "deck_planner", "slide_worker", "infographic_designer", "deck_critic", "slide_critic", "slide_repair", "slide_edit", "dark_theme_designer", "prompt_critic"}:
        raise ValueError(f"Unknown agent role: {role}")
    return (Path(__file__).with_name("prompts") / f"{role}.txt").read_text(encoding="utf-8").strip()
