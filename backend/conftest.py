"""Tests must never use developer credentials or make paid network requests."""
import os
import pytest
from engine import provider


@pytest.fixture(autouse=True)
def isolated_model_environment(monkeypatch):
    for key in list(os.environ):
        if key.startswith(("MODEL_", "VISION_", "OPENROUTER_")) or key == "TEXT_MODEL":
            monkeypatch.delenv(key)
    monkeypatch.setenv("OPENROUTER_API_KEY", "")
    monkeypatch.setenv("MODEL_REQUESTS_PER_MINUTE", "0")
    # Template fonts come from the system and the local store only.
    monkeypatch.setenv("AYA_FONT_DOWNLOAD", "0")
    # The adaptive layout tests set their own mode; strict template tests switch it on.
    monkeypatch.setenv("AYA_STRICT_TEMPLATE", "0")
    def unexpected_network(*args, **kwargs):
        pytest.fail("A test tried to contact the live model API")
    monkeypatch.setattr(provider, "_open", unexpected_network)
