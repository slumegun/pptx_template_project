"""OpenRouter-only model gateway with shared quotas and isolated agent contexts."""

from __future__ import annotations

import base64
import binascii
import http.client
import json
import math
import os
import re
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import deque
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from typing import Any

from .model_registry import OPENROUTER_BASE_URL, ROLE_DEFAULTS, role_model, model_manifest


class ModelProviderError(RuntimeError):
    """A user-visible error without request bodies or credentials."""


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # A redirected POST must never carry the API key to another endpoint.
        return None


def _open(request: urllib.request.Request, timeout: float):
    return urllib.request.build_opener(_NoRedirect()).open(request, timeout=timeout)


def _number(name: str, default: str, minimum: float, maximum: float, *, integer: bool = False):
    try:
        value = int(os.getenv(name, default)) if integer else float(os.getenv(name, default))
    except ValueError:
        raise ModelProviderError(f"{name} must be a number") from None
    if not math.isfinite(value) or not minimum <= value <= maximum:
        raise ModelProviderError(f"{name} must be between {minimum:g} and {maximum:g}")
    return value


def _validate_gateway() -> None:
    _choice("MODEL_PROVIDER", "openrouter", {"openrouter"})
    for name in ("OPENROUTER_BASE_URL", "MODEL_BASE_URL", "VISION_BASE_URL"):
        value = os.getenv(name, "").strip().rstrip("/")
        if value and value != OPENROUTER_BASE_URL:
            raise ModelProviderError(f"{name}: only https://openrouter.ai/api/v1 is allowed")
    if os.getenv("MODEL_DUAL_PROVIDER", "0") != "0":
        raise ModelProviderError("MODEL_DUAL_PROVIDER was removed; all agents use OpenRouter")
    if any(os.getenv(name, "").strip() for name in ("MODEL_API_KEY", "VISION_API_KEY")):
        raise ModelProviderError("Legacy model credentials must be removed; use OPENROUTER_API_KEY only")
    if _choice("MODEL_FREE_ONLY", "0", {"0", "1"}) == "1":
        raise ModelProviderError("This reviewed profile uses paid model IDs; set MODEL_FREE_ONLY=0")


def _choice(name: str, default: str, values: set[str]) -> str:
    value = os.getenv(name, default).strip()
    if value not in values:
        raise ModelProviderError(f"{name} has an unsupported value")
    return value


def _remaining(deadline: float) -> float:
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise ModelProviderError("Model request deadline exceeded")
    return remaining


def _pause(seconds: float, deadline: float) -> None:
    if seconds >= _remaining(deadline):
        raise ModelProviderError("Model retry or rate-limit wait exceeds the remaining time budget")
    if seconds > 0:
        time.sleep(seconds)


def _retry_after(value: str | None) -> float:
    if not value:
        return 0
    try:
        seconds = float(value.rstrip("s"))
    except ValueError:
        try:
            date = parsedate_to_datetime(value)
            if date.tzinfo is None:
                date = date.replace(tzinfo=timezone.utc)
            seconds = (date - datetime.now(timezone.utc)).total_seconds()
        except (ValueError, OverflowError, TypeError):
            return 0
    return max(0, seconds) if math.isfinite(seconds) else 0


def _invalid_constant(value: str):
    raise ValueError("Non-finite JSON number")


class _RequestBudget:
    """One account-level semaphore, rate limiter and ledger for every role."""
    def __init__(self):
        self.max_parallel = _number("MODEL_MAX_PARALLEL", "4", 1, 8, integer=True)
        self.requests_per_minute = _number("MODEL_REQUESTS_PER_MINUTE", "120", 0, 6000)
        self.tokens_per_minute = _number("MODEL_TOKENS_PER_MINUTE", "0", 0, 10000000, integer=True)
        self.slots = threading.BoundedSemaphore(self.max_parallel)
        self._rate_lock = threading.Lock()
        self._next_request = 0.0
        self._token_reservations: deque[tuple[float, int]] = deque()
        self.calls: list[dict[str, Any]] = []
        self.calls_lock = threading.Lock()
        self.blocked: str | None = None

    def wait(self, tokens: int, deadline: float) -> None:
        if self.blocked:
            raise ModelProviderError(self.blocked)
        if self.tokens_per_minute and tokens > self.tokens_per_minute:
            raise ModelProviderError("Model request exceeds the configured token budget; shorten the context or max_tokens")
        while True:
            with self._rate_lock:
                if self.blocked:
                    raise ModelProviderError(self.blocked)
                now = time.monotonic()
                while self._token_reservations and self._token_reservations[0][0] <= now - 60:
                    self._token_reservations.popleft()
                wait = max(0, self._next_request - now)
                reserved = sum(count for _, count in self._token_reservations)
                if self.tokens_per_minute and reserved + tokens > self.tokens_per_minute:
                    wait = max(wait, self._token_reservations[0][0] + 60 - now)
                if wait <= 0:
                    _remaining(deadline)
                    if self.requests_per_minute:
                        self._next_request = now + 60 / self.requests_per_minute
                    if self.tokens_per_minute:
                        self._token_reservations.append((now, tokens))
                    return
            _pause(wait, deadline)


class ModelClient:
    def __init__(self, role: str = "deck_planner", *, budget: _RequestBudget | None = None) -> None:
        _validate_gateway()
        try:
            self.spec = role_model(role)
        except ValueError as error:
            raise ModelProviderError(str(error)) from None
        self.role = role
        self.provider = "openrouter"
        self.base_url = self.vision_base_url = OPENROUTER_BASE_URL
        self.api_key = self.vision_api_key = os.getenv("OPENROUTER_API_KEY", "").strip()
        self.text_model = self.vision_model = self.spec.model_id
        self.timeout = _number("MODEL_TIMEOUT_SECONDS", "45", 1, 300)
        self.request_budget = _number("MODEL_REQUEST_BUDGET_SECONDS", "75", 1, 300)
        self.max_retries = _number("MODEL_MAX_RETRIES", "2", 0, 4, integer=True)
        self.backoff = _number("MODEL_RETRY_BACKOFF_SECONDS", "1", 0, 30)
        self.temperature = _number(f"MODEL_{role.upper()}_TEMPERATURE", "0.1" if "critic" in role else "0.2", 0, 2)
        self.json_mode = self.vision_json_mode = "json_object"
        self.reasoning_effort = self.vision_reasoning_effort = _choice(
            f"MODEL_{role.upper()}_REASONING", "none", {"none", "minimal", "low", "medium", "high"})
        if not self.spec.reasoning and self.reasoning_effort != "none":
            raise ModelProviderError("This model does not support configurable reasoning")
        self.provider_sort = _choice("OPENROUTER_PROVIDER_SORT", "throughput", {"throughput", "latency", "price"})
        self.free_only = False
        self._budget = budget or _RequestBudget()
        self.max_parallel = self._budget.max_parallel
        self._slots = self._budget.slots
        self.calls = self._budget.calls
        self._calls_lock = self._budget.calls_lock

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)

    @property
    def vision_enabled(self) -> bool:
        return self.enabled and self.spec.vision

    def configuration_summary(self) -> dict[str, Any]:
        return {
            "provider": self.provider, "role": self.role,
            "configured": self.enabled, "vision_configured": self.vision_enabled,
            "model": self.spec.model_id, "text_model": self.text_model, "vision_model": self.vision_model,
            "parameters_billion": self.spec.parameters_billion, "license": self.spec.license,
            "weights_url": self.spec.weights_url,
            "endpoint_host": "openrouter.ai", "vision_endpoint_host": "openrouter.ai",
            "timeout_seconds": self.timeout, "request_budget_seconds": self.request_budget,
            "max_parallel": self.max_parallel, "max_retries": self.max_retries,
            "requests_per_minute": self._budget.requests_per_minute,
            "tokens_per_minute": self._budget.tokens_per_minute,
            "json_mode": self.json_mode, "reasoning_effort": self.reasoning_effort,
            "temperature": self.temperature, "provider_sort": self.provider_sort,
            "model_fallbacks": False,
        }

    def complete_json(
        self,
        system: str,
        user: str,
        *,
        vision: bool = False,
        image_data_url: str | None = None,
        max_tokens: int = 1800,
        deadline: float | None = None,
        schema: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Use a fresh context; deadline is an absolute time.monotonic() value."""
        if not (self.vision_enabled if vision else self.enabled):
            needed = "OPENROUTER_API_KEY"
            raise ModelProviderError(f"{needed} are required")
        if not isinstance(max_tokens, int) or isinstance(max_tokens, bool) or not 1 <= max_tokens <= 32768:
            raise ModelProviderError("max_tokens must be between 1 and 32768")
        if image_data_url and not vision:
            raise ModelProviderError("Image input requires vision=True")
        if deadline is not None and not math.isfinite(deadline):
            raise ModelProviderError("Model deadline must be finite")
        # A long answer (the plan of a big deck) streams longer than a usual request:
        # the attempt and its retries get time for about 50 tokens a second.
        writing = max_tokens / 50
        attempt_timeout = max(self.timeout, writing)
        stop_at = time.monotonic() + max(self.request_budget, writing * 1.6)
        if deadline is not None:
            stop_at = min(stop_at, deadline)
        user_content: str | list[dict[str, Any]] = user
        if image_data_url:
            if len(image_data_url) > 20 * 1024 * 1024 or not re.match(r"^data:image/(png|jpeg|webp);base64,", image_data_url):
                raise ModelProviderError("Vision input must be an inline PNG, JPEG or WebP under 20 MiB")
            try:
                base64.b64decode(image_data_url.split(",", 1)[1], validate=True)
            except (ValueError, binascii.Error):
                raise ModelProviderError("Vision input contains invalid base64 data") from None
            user_content = [
                {"type": "text", "text": user},
                {"type": "image_url", "image_url": {"url": image_data_url}},
            ]
        # Recheck immediately before sending; callers cannot retarget a client.
        if self.base_url != OPENROUTER_BASE_URL or self.vision_base_url != OPENROUTER_BASE_URL:
            raise ModelProviderError("Only the OpenRouter endpoint is allowed")
        if self.text_model != self.spec.model_id or self.vision_model != self.spec.model_id:
            raise ModelProviderError("Model assignment changed outside the reviewed registry")
        model = self.spec.model_id
        base_url = OPENROUTER_BASE_URL
        api_key = self.api_key
        json_mode = self.vision_json_mode if vision else self.json_mode
        effort = self.vision_reasoning_effort if vision else self.reasoning_effort
        system = system + "\nReturn exactly one valid JSON object, without Markdown or commentary."
        payload: dict[str, Any] = {
            "model": model,
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": user_content}],
            "temperature": self.temperature,
            "max_tokens": max_tokens,
        }
        if schema is not None:
            payload["response_format"] = {"type": "json_schema", "json_schema": {"name": "agent_response", "strict": True, "schema": schema}}
        elif json_mode == "json_object":
            payload["response_format"] = {"type": "json_object"}
        if self.spec.reasoning:
            payload["reasoning"] = {"enabled": False} if effort == "none" else {"effort": effort}
        # OpenRouter may fail over between hosts serving this exact model only.
        # Strict parameter support avoids silently dropping response_format.
        payload["provider"] = {"sort": self.provider_sort, "require_parameters": True, "allow_fallbacks": True}
        request = urllib.request.Request(
            base_url + "/chat/completions",
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={"Authorization": "Bearer " + api_key, "Content-Type": "application/json", "Accept": "application/json", "User-Agent": "aya-backend/0.1"},
            method="POST",
        )
        # An estimate, not a tokenizer: keep a margin and still honor provider 429/Retry-After.
        tokens = math.ceil(len((system + user).encode("utf-8")) / 2) + 64 + max_tokens
        if image_data_url:
            tokens += 2048
        if not self._slots.acquire(timeout=_remaining(stop_at)):
            raise ModelProviderError("Model concurrency wait exceeds the remaining time budget")
        try:
            return self._request(request, tokens, stop_at, attempt_timeout)
        finally:
            self._slots.release()

    def _request(self, request: urllib.request.Request, tokens: int, deadline: float,
                 timeout: float | None = None) -> dict[str, Any]:
        for attempt in range(self.max_retries + 1):
            self._budget.wait(tokens, deadline)
            retry_delay = self.backoff * 2**attempt
            try:
                started = time.monotonic()
                if self._budget.blocked:
                    raise ModelProviderError(self._budget.blocked)
                with _open(request, min(timeout or self.timeout, _remaining(deadline))) as response:
                    data = response.read(4 * 1024 * 1024 + 1)
                _remaining(deadline)
                if len(data) > 4 * 1024 * 1024:
                    raise ModelProviderError("Model response exceeds the 4 MiB limit")
                result = self._parse_response(data)
                body = json.loads(data)
                sent = json.loads(request.data)
                usage = body.get("usage") or {}
                with self._calls_lock:
                    self.calls.append({"role": self.role, "model": sent["model"], "status": "succeeded", "attempt": attempt + 1, "host": urllib.parse.urlsplit(request.full_url).hostname,
                        "started": started, "finished": time.monotonic(),
                        "prompt_tokens": usage.get("prompt_tokens"), "completion_tokens": usage.get("completion_tokens"),
                        "cost_usd": usage.get("cost"), "upstream_provider": body.get("provider"), "generation_id": body.get("id"), "served_model": body.get("model")})
                return result
            except urllib.error.HTTPError as error:
                status = error.code
                self._record_failure(started, attempt, status)
                retry_delay = max(retry_delay, _retry_after(error.headers.get("Retry-After") if error.headers else None))
                if status == 429 and error.headers and error.headers.get("x-should-retry") == "false":
                    error.close()
                    raise ModelProviderError("Model request exceeds a provider rate window; reduce max_tokens or the request batch") from None
                if status == 429:
                    retry_delay = max(retry_delay, 2)
                error.close()
                reason = {
                    400: "request rejected; check model settings and context size",
                    401: "authentication failed; check the API key",
                    402: "OpenRouter balance exhausted; top up the account and retry the job",
                    403: "access denied by the provider",
                    404: "endpoint or configured model is unavailable",
                    413: "request exceeds provider token or image size limits; reduce the batch",
                    429: "provider rate limit exceeded",
                }.get(status, "provider request failed")
                message = f"Model API HTTP {status}: {reason}"
                if status in {401, 402, 403}:
                    self._budget.blocked = message
                if status not in {408, 429, 500, 502, 503, 504}:
                    raise ModelProviderError(message) from None
            except (urllib.error.URLError, TimeoutError, ConnectionError, http.client.HTTPException):
                self._record_failure(started, attempt, None)
                message = "Model API unavailable or timed out"
            if attempt == self.max_retries:
                raise ModelProviderError(message) from None
            _pause(retry_delay, deadline)
        raise ModelProviderError("Model API request failed")

    def _record_failure(self, started: float, attempt: int, status: int | None) -> None:
        with self._calls_lock:
            self.calls.append({"role": self.role, "model": self.spec.model_id, "host": "openrouter.ai",
                               "status": "failed", "http_status": status, "attempt": attempt + 1,
                               "started": started, "finished": time.monotonic(), "cost_usd": None})

    @staticmethod
    def _parse_response(data: bytes) -> dict[str, Any]:
        try:
            body = json.loads(data, parse_constant=_invalid_constant)
            choice = body["choices"][0]
            if choice.get("finish_reason") == "length":
                raise ModelProviderError("Model response was truncated; reduce the task or increase max_tokens")
            message = choice["message"]
            if message.get("refusal") or choice.get("finish_reason") == "content_filter":
                raise ModelProviderError("Model provider declined this request")
            content = message["content"]
            if isinstance(content, list):
                content = "".join(part["text"] for part in content if part.get("type") == "text")
            content = content.strip()
            fence = re.fullmatch(r"```(?:json)?\s*\n(.*?)\n?```", content, flags=re.DOTALL)
            if fence:
                content = fence.group(1).strip()
            result = json.loads(content, parse_constant=_invalid_constant)
        except (KeyError, IndexError, TypeError, ValueError, AttributeError, RecursionError):
            raise ModelProviderError("Model response is not a valid JSON object") from None
        if not isinstance(result, dict):
            raise ModelProviderError("Model response is not a JSON object")
        return result


class ModelGateway:
    """Validate the whole profile before the first call and share account limits."""
    def __init__(self):
        self.budget = _RequestBudget()
        self.clients = {role: ModelClient(role, budget=self.budget) for role in ROLE_DEFAULTS}
        self.enable_visual_critic = _choice("MODEL_VISUAL_CRITIC", "1", {"0", "1"}) == "1"
        self.repair_rounds = _number("MODEL_REPAIR_ROUNDS", "1", 0, 1, integer=True)

    def client(self, role: str) -> ModelClient:
        return self.clients[role]

    @property
    def enabled(self) -> bool:
        return self.client("deck_planner").enabled

    @property
    def calls(self) -> list[dict[str, Any]]:
        return sorted(self.budget.calls, key=lambda item: item["started"])

    def configuration_summary(self) -> list[dict[str, Any]]:
        return [client.configuration_summary() for client in self.clients.values()]


def configuration_status() -> dict[str, Any]:
    """Safe, offline worker capability report; configuration is not a balance check."""
    try:
        gateway = ModelGateway()
        configured = gateway.enabled
        return {
            "model_mode": "configured_api" if configured else "unknown",
            "model_configured": configured,
            "configuration_status": "configured" if configured else "missing_key",
            "provider": "openrouter", "endpoint_host": "openrouter.ai",
            "text_model": gateway.client("deck_planner").text_model,
            "vision_model": gateway.client("template_analyst").vision_model,
            "agents": gateway.configuration_summary(), "model_manifest": model_manifest(),
            "features": {"slide_workers": True, "visual_critic": gateway.enable_visual_critic,
                         "semantic_repair_rounds": gateway.repair_rounds},
            "live_verified": False,
            "configuration_error": None if configured else "Set OPENROUTER_API_KEY on the worker",
        }
    except ModelProviderError as error:
        return {"model_mode": "unknown", "model_configured": False, "configuration_status": "invalid",
                "provider": "openrouter", "endpoint_host": "openrouter.ai", "text_model": None,
                "vision_model": None, "agents": [], "model_manifest": {}, "live_verified": False,
                "configuration_error": str(error)}
