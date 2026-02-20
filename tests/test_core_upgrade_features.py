"""Tests for core-upgrade features backported from upstream."""

from __future__ import annotations

import asyncio

from Mudabbir.agents.registry import get_backend_info, list_backend_names, normalize_backend_name
from Mudabbir.bus.commands import get_command_handler
from Mudabbir.bus.events import Channel, InboundMessage
from Mudabbir.deep_work.goal_parser import GoalParser


def test_backend_registry_normalization() -> None:
    """Registry should expose canonical names and alias normalization."""
    names = list_backend_names()
    assert "claude_agent_sdk" in names
    assert "Mudabbir_native" in names
    assert "open_interpreter" in names
    assert "openai_agents" in names
    assert "google_adk" in names
    assert "codex_cli" in names
    assert "opencode" in names
    assert "copilot_sdk" in names
    assert normalize_backend_name("claude_code") == "claude_agent_sdk"
    assert normalize_backend_name("adk") == "google_adk"
    assert get_backend_info("open_interpreter").display_name == "Open Interpreter"


def test_goal_parser_outputs_valid_shape() -> None:
    """Goal parser should return a structured analysis with valid enums."""
    parser = GoalParser()
    analysis = asyncio.run(
        parser.parse("Build a React + FastAPI dashboard with auth and deployment in 2 weeks")
    )
    assert analysis.domain in {"code", "business", "creative", "education", "events", "home", "hybrid"}
    assert analysis.complexity in {"S", "M", "L", "XL"}
    assert analysis.suggested_research_depth in {"none", "quick", "standard", "deep"}
    assert analysis.estimated_phases >= 1


def test_backends_command_available() -> None:
    """Command handler should support /backends response."""
    handler = get_command_handler()
    message = InboundMessage(channel=Channel.CLI, sender_id="tester", chat_id="cli", content="/backends")
    response = asyncio.run(handler.handle(message))
    assert response is not None
    assert "Available Backends" in response.content

