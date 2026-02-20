"""Tests for Mem0 validation and MCP cleanup hardening."""

from __future__ import annotations

import asyncio

from Mudabbir.mcp.config import MCPServerConfig
from Mudabbir.mcp.manager import MCPManager, _ServerState
from Mudabbir.memory.validation import validate_mem0_settings


def test_validate_mem0_settings_rejects_model_like_ollama_url() -> None:
    """Model names must not be accepted as mem0_ollama_base_url."""
    errors = validate_mem0_settings(
        {
            "memory_backend": "mem0",
            "mem0_llm_provider": "ollama",
            "mem0_embedder_provider": "ollama",
            "mem0_llm_model": "deepseek-v3.1:671b-cloud",
            "mem0_embedder_model": "nomic-embed-text",
            "mem0_ollama_base_url": "deepseek-v3.1:671b-cloud",
        }
    )
    assert errors
    assert any("mem0_ollama_base_url" in err for err in errors)


def test_validate_mem0_settings_accepts_valid_ollama_url() -> None:
    """A valid http(s) endpoint must pass when mem0 uses ollama."""
    errors = validate_mem0_settings(
        {
            "memory_backend": "mem0",
            "mem0_llm_provider": "ollama",
            "mem0_embedder_provider": "ollama",
            "mem0_llm_model": "llama3.2",
            "mem0_embedder_model": "nomic-embed-text",
            "mem0_ollama_base_url": "http://localhost:11434",
        }
    )
    assert errors == []


def test_mcp_cleanup_state_swallows_cancelled_error() -> None:
    """Shutdown cleanup should not raise when context managers cancel."""

    class _CancelledContext:
        async def __aexit__(self, exc_type, exc, tb):
            raise asyncio.CancelledError()

    manager = MCPManager()
    state = _ServerState(config=MCPServerConfig(name="test"))
    state.session = _CancelledContext()
    state.client = _CancelledContext()

    asyncio.run(manager._cleanup_state(state))
    assert state.connected is False

