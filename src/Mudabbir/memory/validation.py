"""Validation helpers for memory backend settings."""

from __future__ import annotations

import re
from typing import Any
from urllib.parse import urlparse


def _get_field(source: Any, key: str, default: Any = None) -> Any:
    """Read field from dict-like payloads or Settings-like objects."""
    if isinstance(source, dict):
        return source.get(key, default)
    return getattr(source, key, default)


def _looks_like_model_name(value: str) -> bool:
    """Heuristic: detect model-like values accidentally used as URL."""
    value = (value or "").strip()
    if not value or "://" in value:
        return False
    if "/" in value:
        return False
    # Typical model tags: "name:tag", "deepseek-v3.1:671b-cloud", etc.
    return bool(re.match(r"^[a-zA-Z0-9._-]+:[a-zA-Z0-9._-]+$", value))


def _is_http_url(value: str) -> tuple[bool, str]:
    """Validate a URL for Ollama host usage."""
    value = (value or "").strip()
    if not value:
        return False, "is empty"
    try:
        parsed = urlparse(value)
        if parsed.scheme not in {"http", "https"}:
            return False, "must start with http:// or https://"
        if not parsed.hostname:
            return False, "is missing a hostname"
        # Accessing .port raises ValueError for invalid port strings.
        _ = parsed.port
        return True, ""
    except ValueError as exc:
        return False, str(exc)
    except Exception:
        return False, "is not a valid URL"


def validate_mem0_settings(settings_or_payload: Any) -> list[str]:
    """Return a list of validation errors for Mem0 configuration."""
    memory_backend = str(_get_field(settings_or_payload, "memory_backend", "file") or "file").strip()
    if memory_backend.lower() != "mem0":
        return []

    llm_provider = str(_get_field(settings_or_payload, "mem0_llm_provider", "") or "").strip().lower()
    embedder_provider = str(
        _get_field(settings_or_payload, "mem0_embedder_provider", "") or ""
    ).strip().lower()
    llm_model = str(_get_field(settings_or_payload, "mem0_llm_model", "") or "").strip()
    embedder_model = str(_get_field(settings_or_payload, "mem0_embedder_model", "") or "").strip()
    ollama_base_url = str(
        _get_field(settings_or_payload, "mem0_ollama_base_url", "") or ""
    ).strip()

    errors: list[str] = []
    uses_ollama = llm_provider == "ollama" or embedder_provider == "ollama"

    if uses_ollama:
        ok, reason = _is_http_url(ollama_base_url)
        if not ok:
            hint = (
                "Use a valid endpoint URL, for example: http://localhost:11434 "
                "(not a model name)."
            )
            msg = f"mem0_ollama_base_url {reason}. {hint}"
            if _looks_like_model_name(ollama_base_url):
                msg = (
                    "mem0_ollama_base_url looks like a model name "
                    f"('{ollama_base_url}'), not an endpoint URL. {hint}"
                )
            errors.append(msg)

    if llm_provider == "ollama":
        if not llm_model:
            errors.append(
                "mem0_llm_model is required when mem0_llm_provider=ollama "
                "(example: llama3.2)."
            )
        if llm_model.startswith(("http://", "https://")):
            errors.append(
                "mem0_llm_model must be a model name, not a URL "
                "(example: llama3.2)."
            )

    if embedder_provider == "ollama":
        if not embedder_model:
            errors.append(
                "mem0_embedder_model is required when mem0_embedder_provider=ollama "
                "(example: nomic-embed-text)."
            )
        if embedder_model.startswith(("http://", "https://")):
            errors.append(
                "mem0_embedder_model must be a model name, not a URL "
                "(example: nomic-embed-text)."
            )

    return errors

