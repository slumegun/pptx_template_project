"""Validate OpenRouter roles offline; use --catalog or --live explicitly."""
from __future__ import annotations

import argparse
import base64
import io
import json
import sys
import time
import urllib.request
from pathlib import Path

from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--live", action="store_true", help="Send paid synthetic probes to all nine agent roles")
    parser.add_argument("--catalog", action="store_true", help="Verify the public model catalog without a key or inference")
    parser.add_argument("--text-only", action="store_true", help="With --live, skip the two image roles")
    parser.add_argument("--env-file", type=Path, help="Explicit alternative profile; replaces the default env file")
    parser.add_argument("--output", type=Path, help="Write the non-secret JSON report")
    args = parser.parse_args()
    env_file = args.env_file or ROOT / ".env"
    if args.env_file and not env_file.is_file():
        parser.error("Settings file does not exist")
    load_dotenv(env_file, override=False)
    from engine.provider import ModelGateway, ModelProviderError, configuration_status, _open
    from engine.model_registry import OPENROUTER_BASE_URL, VISION_ROLES
    report = {"configuration": configuration_status(), "checks": [], "live_verified": False}
    try:
        gateway = ModelGateway()
        if args.catalog:
            request = urllib.request.Request(OPENROUTER_BASE_URL + "/models", headers={"Accept": "application/json"})
            with _open(request, 30) as response:
                catalog = {row["id"]: row for row in json.load(response)["data"]}
            for model in sorted({client.spec.model_id for client in gateway.clients.values()}):
                if model not in catalog:
                    raise ModelProviderError("A configured model is absent from the public OpenRouter catalog")
                row = catalog[model]
                modalities = row.get("architecture", {}).get("input_modalities", [])
                parameters = row.get("supported_parameters", [])
                if "image" not in modalities or "response_format" not in parameters:
                    raise ModelProviderError("Configured model lacks image input or JSON response support in the catalog")
                report["checks"].append({"model": model, "status": "catalog_found", "pricing": row.get("pricing"),
                                         "input_modalities": modalities, "supported_parameters": parameters})
        if args.live:
            if not gateway.enabled:
                raise ModelProviderError("--live requires OPENROUTER_API_KEY")
            from PIL import Image
            buffer = io.BytesIO()
            Image.new("RGB", (64, 64), (255, 0, 0)).save(buffer, format="PNG")
            image = "data:image/png;base64," + base64.b64encode(buffer.getvalue()).decode("ascii")
            for role, client in gateway.clients.items():
                vision = role in VISION_ROLES
                if args.text_only and vision:
                    continue
                started = time.monotonic()
                field, expected = ("color", "red") if vision else ("status", "ok")
                schema = {"type": "object", "additionalProperties": False, "required": [field],
                          "properties": {field: {"type": "string"}}}
                system = 'Identify the dominant color. Return JSON with color.' if vision else 'Return JSON {"status":"ok"}.'
                answer = client.complete_json(system, "Synthetic connection test", vision=vision,
                    image_data_url=image if vision else None, max_tokens=256,
                    schema=schema if role == "visual_critic" else None)
                if answer.get(field) != expected:
                    raise ModelProviderError("An agent returned an unexpected probe result")
                report["checks"].append({"role": role, "model": client.spec.model_id, "status": "ok",
                                         "seconds": round(time.monotonic() - started, 3)})
            report["model_calls"] = gateway.calls
            report["live_verified"] = not args.text_only
        report["status"] = "passed"
    except (ModelProviderError, OSError, ValueError, KeyError) as error:
        # Never print transport exception strings: they may include remote response data.
        report.update(status="failed", error=str(error) if isinstance(error, ModelProviderError) else "Catalog or transport check failed")
    encoded = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded)
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
