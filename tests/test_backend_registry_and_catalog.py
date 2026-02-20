"""Backend registry tests for expanded multi-backend catalog."""

from __future__ import annotations

from Mudabbir.agents.registry import (
    backend_summary,
    get_backend_info,
    install_hint_text,
    is_backend_available,
    list_backend_summaries,
)


def test_backend_summary_shape() -> None:
    summary = backend_summary("claude_agent_sdk")
    assert summary["name"] == "claude_agent_sdk"
    assert "displayName" in summary
    assert isinstance(summary["capabilities"], list)
    assert isinstance(summary["installHint"], dict)


def test_backend_catalog_contains_expanded_backends() -> None:
    names = {item["name"] for item in list_backend_summaries()}
    assert {"openai_agents", "google_adk", "codex_cli", "opencode", "copilot_sdk"} <= names


def test_install_hint_text_renders() -> None:
    info = get_backend_info("google_adk")
    hint = install_hint_text(info)
    assert "pip install" in hint


def test_backend_availability_returns_boolean() -> None:
    assert isinstance(is_backend_available("claude_agent_sdk"), bool)
    assert isinstance(is_backend_available("codex_cli"), bool)
