import json
from pathlib import Path

import pytest

from engine.models import Fact, SlideContent
from engine.provider import ModelProviderError
from engine.swarm import critique_content, critique_prompt, critique_visual_slide, critique_visual_variant


class Critic:
    max_parallel = 2

    def __init__(self, response):
        self.response = response
        self.calls = []

    def complete_json(self, system, user, **kwargs):
        self.calls.append((system, json.loads(user), kwargs))
        return self.response


def sample():
    slides = [
        SlideContent("slide_1", "Project goal", ["The project creates slides."], ["f1"]),
        SlideContent("slide_2", "Result", ["The file remains editable."], ["f2"]),
    ]
    facts = [
        Fact("f1", "The project creates slides.", "brief", "line 1", "The project creates slides."),
        Fact("f2", "The file remains editable.", "notes.txt", "line 3", "The file remains editable."),
    ]
    return slides, facts


def test_prompt_critic_checks_original_assignment_and_returns_bounded_issue():
    slides, facts = sample()
    brief = "Explain the project, then the result. Include the launch date."
    client = Critic({"issues": [{
        "slide_id": "slide_1",
        "rule_id": "prompt_required_point_missing",
        "severity": "warning",
        "evidence": "The brief requires a launch date, but no slide gives one.",
    }]})
    issues = critique_prompt(client, brief, slides, facts, 123)
    system, payload, kwargs = client.calls[0]
    assert "Prompt Compliance Critic" in system
    assert payload["brief"] == brief
    assert [slide["slide_id"] for slide in payload["slides"]] == ["slide_1", "slide_2"]
    assert payload["primary_sources"][1]["source"] == "notes.txt"
    assert kwargs["deadline"] == 123
    assert kwargs["schema"]["properties"]["issues"]["maxItems"] == 40
    assert kwargs["schema"]["properties"]["issues"]["items"]["properties"]["rule_id"]["enum"] == [
        "prompt_order_mismatch", "prompt_required_point_missing",
        "prompt_topic_mismatch", "prompt_unsupported_claim",
    ]
    assert issues[0]["slide_id"] == "slide_1"
    assert issues[0]["rule_id"] == "prompt_required_point_missing"
    assert issues[0]["severity"] == "blocking"
    assert issues[0]["repairability"] == "manual"


@pytest.mark.parametrize("response", [
    {"issues": [{"slide_id": "slide_9", "rule_id": "prompt_topic_mismatch",
                 "severity": "blocking", "evidence": "Wrong topic"}]},
    {"issues": [{"slide_id": "slide_1", "rule_id": "unknown_rule",
                 "severity": "blocking", "evidence": "Wrong topic"}]},
    {"issues": [{"slide_id": "slide_1", "rule_id": "prompt_topic_mismatch",
                 "severity": "critical", "evidence": "Wrong topic"}]},
    {"issues": [{"slide_id": "slide_1", "rule_id": "prompt_topic_mismatch",
                 "severity": "blocking", "evidence": "  "}]},
    {"issues": [{"slide_id": "slide_1", "rule_id": "prompt_topic_mismatch",
                 "severity": "blocking", "evidence": "x" * 1201}]},
    {"issues": [{}] * 41},
    {"issues": "none"},
    {"issues": [], "extra": "unexpected"},
    {"issues": [{"slide_id": "slide_1", "rule_id": "prompt_topic_mismatch",
                 "severity": "blocking", "evidence": "Wrong topic", "extra": "unexpected"}]},
    [],
])
def test_prompt_critic_rejects_invalid_model_reports(response):
    slides, facts = sample()
    with pytest.raises(ModelProviderError):
        critique_prompt(Critic(response), "Explain the project.", slides, facts, 123)


def test_prompt_critic_deduplicates_identical_observations():
    slides, facts = sample()
    observation = {"slide_id": "slide_2", "rule_id": "prompt_order_mismatch",
                   "severity": "blocking", "evidence": "  The result came before the project.  "}
    issues = critique_prompt(Critic({"issues": [observation, observation.copy()]}),
                             "Project first, then result.", slides, facts, 123)
    assert len(issues) == 1
    assert issues[0]["slide_id"] == "slide_2"
    assert issues[0]["evidence"] == "The result came before the project."


def test_prompt_critic_needs_an_assignment_and_unique_slides():
    slides, facts = sample()
    client = Critic({"issues": []})
    with pytest.raises(ValueError):
        critique_prompt(client, " ", slides, facts, 123)
    with pytest.raises(ValueError):
        critique_prompt(client, "Topic", [], facts, 123)
    with pytest.raises(ValueError):
        critique_prompt(client, "Topic", [slides[0], slides[0]], facts, 123)
    assert not client.calls


def test_visual_critic_receives_theme_without_treating_dark_as_invalid(monkeypatch):
    monkeypatch.setattr("engine.swarm.contact_sheet", lambda paths: "test-image")
    client = Critic({"issues": []})
    slides, _ = sample()
    assert critique_visual_slide(client, Path("source"), Path("output"), slides[0], 0, 123,
                                 variant_theme="dark") == []
    assert client.calls[0][1]["variant_theme"] == "dark"
    assert client.calls[0][2]["vision"] is True
    assert critique_visual_variant(client, [Path("source")], [Path("one"), Path("two")],
                                   [1, 1], slides, 123, variant_theme="dark") == []
    assert all(call[1]["variant_theme"] == "dark" for call in client.calls)
    with pytest.raises(ValueError, match="theme"):
        critique_visual_slide(client, Path("source"), Path("output"), slides[0], 0, 123,
                              variant_theme="unknown")


def test_deck_critic_constrains_live_report_to_known_slides_and_rules():
    slides, facts = sample()
    client = Critic({"issues": []})
    assert critique_content(client, slides, facts, 123) == []
    schema = client.calls[0][2]["schema"]
    fields = schema["properties"]["issues"]["items"]["properties"]
    assert fields["slide_id"]["enum"] == ["slide_1", "slide_2"]
    assert "unsupported_claim" in fields["rule_id"]["enum"]
    assert schema["properties"]["issues"]["maxItems"] == 40
