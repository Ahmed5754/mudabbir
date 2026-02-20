"""AI response pipeline and capability registry tests."""

from __future__ import annotations

import pytest

from Mudabbir.agents.loop import AgentLoop
from Mudabbir.bus.adapters import telegram_adapter
from Mudabbir.tools.capabilities.registry import DEFAULT_CAPABILITY_REGISTRY
from Mudabbir.tools.capabilities.windows_catalog import WINDOWS_CAPABILITIES


def test_stage_a_capability_catalog_is_present() -> None:
    stage_a = [cap for cap in WINDOWS_CAPABILITIES if cap.stage == "A"]
    assert stage_a
    assert any(cap.id == "audio.volume" for cap in stage_a)
    assert any(cap.id == "display.brightness" for cap in stage_a)


def test_registry_exposes_stage_a_desktop_actions() -> None:
    actions = DEFAULT_CAPABILITY_REGISTRY.allowed_actions_stage_a()
    assert "volume" in actions
    assert "brightness" in actions
    assert "launch_start_app" in actions


def test_telegram_static_fallback_helper_removed() -> None:
    assert not hasattr(telegram_adapter, "_infer_safe_fallback_text")


@pytest.mark.asyncio
async def test_response_composer_uses_fallback_when_llm_unavailable() -> None:
    loop = AgentLoop()

    async def _fail_llm(*args, **kwargs):  # noqa: ANN002, ANN003
        return None

    loop._llm_one_shot_text = _fail_llm  # type: ignore[method-assign]
    result = await loop._compose_response(
        user_query="كم مستوى الصوت؟",
        events=[{"type": "result", "metadata": {"facts": {"level_percent": 37}}}],
        fallback_text="Current volume: 37% (muted=False)",
    )
    assert result == "Current volume: 37% (muted=False)"
