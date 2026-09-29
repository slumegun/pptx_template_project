from pathlib import Path

import pytest
from engine.models import SlideContent
from engine.provider import ModelProviderError
from engine.swarm import critique_visual_slide, critique_visual_variant


class Critic:
    max_parallel = 2
    def __init__(self, response):
        self.response = response
    def complete_json(self, *args, **kwargs):
        assert kwargs["vision"] is True
        assert kwargs["deadline"] == 100
        return self.response


def test_visual_issues_keep_target_slide_and_cannot_trigger_geometry_repair(monkeypatch):
    monkeypatch.setattr("engine.swarm.contact_sheet", lambda paths: "test")
    client = Critic({"issues": [{"rule_id": "visual_empty_container", "severity": "blocking", "evidence": "Empty right card"}]})
    issues = critique_visual_slide(client, Path("source"), Path("output"), SlideContent("slide_3", "Title", ["Fact"], ["f1"]), 2, 100)
    assert issues[0]["slide_id"] == "slide_3"
    assert issues[0]["severity"] == "warning"
    assert issues[0]["repairability"] == "manual"


def test_visual_critic_rejects_unrecognized_rule(monkeypatch):
    monkeypatch.setattr("engine.swarm.contact_sheet", lambda paths: "test")
    client = Critic({"issues": [{"rule_id": "delete_everything", "severity": "blocking", "evidence": "bad"}]})
    with pytest.raises(ModelProviderError):
        critique_visual_slide(client, Path("source"), Path("out"), SlideContent("slide_1", "Title", [], []), 0, 100)


def test_visual_comparison_uses_actual_source_mapping(monkeypatch):
    def inspect(client, source, rendered, slide, index, deadline):
        assert source.name == {0: "second", 1: "first"}[index]
        return []
    monkeypatch.setattr("engine.swarm.critique_visual_slide", inspect)
    slides = [SlideContent("slide_1", "A", [], []), SlideContent("slide_2", "B", [], [])]
    assert critique_visual_variant(Critic({}), [Path("first"), Path("second")], [Path("a"), Path("b")], [2, 1], slides, 100) == []
    with pytest.raises(ValueError):
        critique_visual_variant(Critic({}), [Path("first")], [Path("a"), Path("b")], [2, 1], slides, 100)


def test_source_preview_cache_is_sorted_and_skips_export(tmp_path, monkeypatch):
    from types import SimpleNamespace
    from engine.pipeline import _visual_references
    monkeypatch.setenv("MODEL_TEMPLATE_CACHE_DIR", str(tmp_path))
    folder = tmp_path / "visual-references-v1" / "hash" / "slides"
    folder.mkdir(parents=True)
    for index in range(1, 12):
        (folder / f"slide-{index}.png").write_bytes(b"cached")
    def unexpected(*args, **kwargs):
        raise AssertionError("A complete cache must not rerender the source")
    monkeypatch.setattr("engine.pipeline.export_pdf", unexpected)
    template = SimpleNamespace(template_sha256="hash", compositions=[None] * 11)
    paths = _visual_references(Path("source.pptx"), template, tmp_path)
    assert [p.name for p in paths] == [f"slide-{i}.png" for i in range(1, 12)]
