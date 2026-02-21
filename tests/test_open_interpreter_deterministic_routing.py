from __future__ import annotations

import pytest

from Mudabbir.agents.open_interpreter import OpenInterpreterAgent
from Mudabbir.config import Settings


@pytest.mark.asyncio
async def test_deterministic_route_executes_desktop_action(monkeypatch: pytest.MonkeyPatch) -> None:
    async def _fake_execute(self, action: str, **params: object) -> str:  # noqa: ANN001
        assert action == "volume"
        assert params.get("mode") == "set"
        assert params.get("level") == 40
        return '{"ok": true, "level_percent": 40}'

    monkeypatch.setattr(OpenInterpreterAgent, "_initialize", lambda self: None)
    from Mudabbir.tools.builtin.desktop import DesktopTool

    monkeypatch.setattr(DesktopTool, "execute", _fake_execute)
    agent = OpenInterpreterAgent(Settings())
    result = await agent._try_intent_map_desktop_response("set volume to 40")
    assert result is not None
    assert result["metadata"]["resolved_action"] == "volume"
    assert result["metadata"]["facts"]["ok"] is True


@pytest.mark.asyncio
async def test_destructive_requires_confirmation(monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list[tuple[str, dict]] = []

    async def _fake_execute(self, action: str, **params: object) -> str:  # noqa: ANN001
        calls.append((action, dict(params)))
        return '{"ok": true}'

    monkeypatch.setattr(OpenInterpreterAgent, "_initialize", lambda self: None)
    from Mudabbir.tools.builtin.desktop import DesktopTool

    monkeypatch.setattr(DesktopTool, "execute", _fake_execute)
    agent = OpenInterpreterAgent(Settings())

    first = await agent._try_intent_map_desktop_response("shutdown now")
    assert first is not None
    assert first["metadata"]["facts"]["status"] == "awaiting_confirmation"
    assert calls == []

    second = await agent._try_intent_map_desktop_response("yes")
    assert second is not None
    assert second["metadata"]["facts"]["ok"] is True
    assert calls[0][0] == "system_power"
    assert calls[0][1]["mode"] == "shutdown"


@pytest.mark.asyncio
async def test_whatsapp_latest_message_uses_vision_describe_screen(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[tuple[str, dict]] = []

    async def _fake_execute(self, action: str, **params: object) -> str:  # noqa: ANN001
        calls.append((action, dict(params)))
        assert action == "vision_tools"
        assert params.get("mode") == "describe_screen"
        return (
            '{"ok": true, "mode": "describe_screen", "source": "vision", '
            '"top_app": "WhatsApp", "ui_summary": "Recent chat list is visible."}'
        )

    monkeypatch.setattr(OpenInterpreterAgent, "_initialize", lambda self: None)
    from Mudabbir.tools.builtin.desktop import DesktopTool

    monkeypatch.setattr(DesktopTool, "execute", _fake_execute)
    agent = OpenInterpreterAgent(Settings())
    result = await agent._try_direct_desktop_response("مين آخر واحد بعتلي على واتساب؟", history=[])
    assert result is not None
    assert result["type"] == "message"
    assert "واتساب" in result["content"] or "الشاشة" in result["content"]
    assert calls and calls[0][0] == "vision_tools"
