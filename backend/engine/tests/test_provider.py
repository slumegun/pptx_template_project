import io
import json
import time
import urllib.error

import pytest

from engine import provider


@pytest.fixture
def client(monkeypatch):
    for name in list(provider.os.environ):
        if name.startswith(("MODEL_", "VISION_", "OPENROUTER_")) or name == "TEXT_MODEL":
            monkeypatch.delenv(name)
    for name, value in {
        "OPENROUTER_API_KEY": "test-secret-key",
        "MODEL_REQUESTS_PER_MINUTE": "0", "MODEL_RETRY_BACKOFF_SECONDS": "0",
    }.items():
        monkeypatch.setenv(name, value)
    return provider.ModelClient()


def response(content, **choice):
    return json.dumps({"choices": [{"message": {"content": content}, **choice}]}).encode()


def test_json_request_and_safe_summary(client, monkeypatch):
    captured = []
    def open_request(request, timeout):
        captured.append(request)
        return io.BytesIO(response('{"ok": true}'))
    monkeypatch.setattr(provider, "_open", open_request)
    assert client.complete_json("Return JSON", "test", max_tokens=20) == {"ok": True}
    request = captured[0]
    assert request.get_header("User-agent") == "aya-backend/0.1"
    assert json.loads(request.data)["model"] == client.text_model
    assert "test-secret-key" not in json.dumps(client.configuration_summary())


@pytest.mark.parametrize("body", [b'{}', b'<html>gateway</html>', response('[]'), response('{"n":NaN}'), response(None)])
def test_malformed_responses_fail_safely(client, body):
    with pytest.raises(provider.ModelProviderError):
        client._parse_response(body)


def test_truncation_is_not_accepted(client):
    with pytest.raises(provider.ModelProviderError, match="truncated"):
        client._parse_response(response('{"ok": true}', finish_reason="length"))


def test_fenced_json_is_supported(client):
    assert client._parse_response(response('```json\n{"ok":true}\n```')) == {"ok": True}


def test_transient_error_retries(client, monkeypatch):
    attempts = []
    def open_request(request, timeout):
        attempts.append(1)
        if len(attempts) == 1:
            raise urllib.error.HTTPError(request.full_url, 503, "unavailable", {}, io.BytesIO(b""))
        return io.BytesIO(response('{"ok": true}'))
    monkeypatch.setattr(provider, "_open", open_request)
    assert client.complete_json("JSON", "test") == {"ok": True}
    assert len(attempts) == 2


def test_auth_failure_does_not_retry_or_echo_body(client, monkeypatch):
    attempts = []
    def open_request(request, timeout):
        attempts.append(1)
        raise urllib.error.HTTPError(request.full_url, 401, "test-secret-key", {}, io.BytesIO(b'test-secret-key'))
    monkeypatch.setattr(provider, "_open", open_request)
    with pytest.raises(provider.ModelProviderError) as error:
        client.complete_json("JSON", "test")
    assert "test-secret-key" not in str(error.value)
    assert len(attempts) == 1


def test_deadline_prevents_network_call(client, monkeypatch):
    monkeypatch.setattr(provider, "_open", lambda *args: pytest.fail("network called"))
    with pytest.raises(provider.ModelProviderError, match="deadline"):
        client.complete_json("JSON", "test", deadline=time.monotonic() - 1)


@pytest.mark.parametrize("name,value", [
    ("MODEL_BASE_URL", "https://api.groq.com/openai/v1"),
    ("VISION_BASE_URL", "https://other.example/v1"),
    ("OPENROUTER_BASE_URL", "https://openrouter.ai.evil.example/api/v1"),
    ("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1?redirect=evil"),
    ("MODEL_BASE_URL", "http://localhost:1234/v1"),
    ("MODEL_PROVIDER", "openai_compatible"),
    ("MODEL_DUAL_PROVIDER", "1"),
    ("MODEL_API_KEY", "old-provider-key"),
])
def test_non_router_and_legacy_config_fails_before_network(client, monkeypatch, name, value):
    monkeypatch.setenv(name, value)
    with pytest.raises(provider.ModelProviderError):
        provider.ModelGateway()


@pytest.mark.parametrize("model", ["openrouter/auto", "qwen/qwen3.8-27b:free", "qwen/qwen3-235b-a22b", "openai/gpt-4o"])
def test_unreviewed_models_rejected(client, monkeypatch, model):
    monkeypatch.setenv("MODEL_DECK_PLANNER", model)
    with pytest.raises(provider.ModelProviderError, match="approved"):
        provider.ModelGateway()


def test_free_only_rejects_paid_profile(client, monkeypatch):
    monkeypatch.setenv("MODEL_FREE_ONLY", "1")
    with pytest.raises(provider.ModelProviderError, match="paid"):
        provider.ModelClient()


def test_remote_image_rejected(client):
    with pytest.raises(provider.ModelProviderError, match="inline"):
        client.complete_json("JSON", "test", vision=True, image_data_url="https://private.example/image")


def test_roles_share_one_key_and_account_limits(client):
    gateway = provider.ModelGateway()
    analyst = gateway.client("template_analyst")
    critic = gateway.client("deck_critic")
    worker = gateway.client("slide_worker")
    assert critic.text_model == "google/gemma-4-31b-it"
    assert worker.text_model == "qwen/qwen3.8-27b"
    assert analyst.text_model == "qwen/qwen3-vl-32b-instruct"
    assert critic.api_key == analyst.api_key == worker.api_key == "test-secret-key"
    assert critic._slots is analyst._slots is worker._slots
    assert critic._budget is worker._budget
    assert "test-secret-key" not in json.dumps(provider.configuration_status())


def test_openrouter_reasoning_disabled_uses_gateway_parameter(client, monkeypatch):
    client.base_url = "https://openrouter.ai/api/v1"
    client.reasoning_effort = "none"
    client.free_only = False
    def inspect(request, timeout):
        payload = json.loads(request.data)
        assert payload["reasoning"] == {"enabled": False}
        assert "reasoning_effort" not in payload
        return io.BytesIO(response('{"ok":true}'))
    monkeypatch.setattr(provider, "_open", inspect)
    assert client.complete_json("JSON", "test") == {"ok": True}


def test_agent_schema_is_forwarded_as_strict_json_schema(client, monkeypatch):
    schema = {"type": "object", "properties": {"ok": {"type": "boolean"}}, "required": ["ok"], "additionalProperties": False}
    def inspect(request, timeout):
        payload = json.loads(request.data)
        assert payload["response_format"]["type"] == "json_schema"
        assert payload["response_format"]["json_schema"]["strict"] is True
        assert payload["response_format"]["json_schema"]["schema"] == schema
        return io.BytesIO(response('{"ok":true}'))
    monkeypatch.setattr(provider, "_open", inspect)
    assert client.complete_json("JSON", "test", schema=schema) == {"ok": True}


def test_payment_failure_is_not_retried_and_stops_other_roles(client, monkeypatch):
    gateway = provider.ModelGateway()
    attempts = []
    def depleted(request, timeout):
        attempts.append(request)
        raise urllib.error.HTTPError(request.full_url, 402, "secret upstream", {}, io.BytesIO(b"secret upstream"))
    monkeypatch.setattr(provider, "_open", depleted)
    for role in ("deck_planner", "slide_worker", "visual_critic"):
        with pytest.raises(provider.ModelProviderError, match="balance exhausted"):
            gateway.client(role).complete_json("JSON", "test")
    assert len(attempts) == 1
    assert gateway.calls[0]["http_status"] == 402
    assert "secret upstream" not in json.dumps(gateway.calls)


def test_no_model_fallback_and_all_calls_only_openrouter(client, monkeypatch):
    calls = []
    def inspect(request, timeout):
        assert request.full_url == "https://openrouter.ai/api/v1/chat/completions"
        payload = json.loads(request.data)
        assert "models" not in payload
        assert "route" not in payload
        assert payload["provider"] == {"sort": "throughput", "require_parameters": True, "allow_fallbacks": True}
        assert request.get_header("Authorization") == "Bearer test-secret-key"
        assert len(payload["messages"]) == 2
        calls.append(payload["model"])
        return io.BytesIO(response('{"ok":true}'))
    monkeypatch.setattr(provider, "_open", inspect)
    gateway = provider.ModelGateway()
    for role in gateway.clients:
        gateway.client(role).complete_json("JSON", role)
    assert len(calls) == 9
    assert {row["role"] for row in gateway.calls} == set(gateway.clients)


def test_retry_after_larger_than_deadline_does_not_retry(client, monkeypatch):
    attempts = []
    def limited(request, timeout):
        attempts.append(1)
        raise urllib.error.HTTPError(request.full_url, 429, "limited", {"Retry-After": "60"}, io.BytesIO())
    monkeypatch.setattr(provider, "_open", limited)
    with pytest.raises(provider.ModelProviderError, match="remaining"):
        client.complete_json("JSON", "test", deadline=time.monotonic() + 0.5)
    assert len(attempts) == 1


def test_missing_key_never_silently_becomes_local_draft(monkeypatch):
    status = provider.configuration_status()
    assert status["model_mode"] == "unknown"
    assert status["configuration_status"] == "missing_key"
    monkeypatch.setenv("MODEL_ALLOW_LOCAL_DRAFT", "1")
    assert provider.configuration_status()["model_mode"] == "unknown"


def test_redirect_handler_never_forwards_credentials(client):
    handler = provider._NoRedirect()
    request = provider.urllib.request.Request(client.base_url, headers={"Authorization": "Bearer secret"})
    assert handler.redirect_request(request, None, 307, "redirect", {}, "https://other.example") is None
