import pytest

from Mudabbir.agents.loop import AgentLoop


def test_sanitize_stream_chunk_hides_mcp_noise() -> None:
    noisy = 'mcp_sequential-thinking__sequentialthinking<｜tool▁sep｜>{"thought":"x"}'
    assert AgentLoop._sanitize_stream_chunk(noisy) == ""


def test_timeout_message_is_provider_aware() -> None:
    loop = AgentLoop()
    text = loop._timeout_message(backend="open_interpreter", provider="ollama")
    assert "Ollama" in text
    assert "Claude Code CLI" not in text


@pytest.mark.asyncio
async def test_global_fastpath_volume_get(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "volume"
            assert kwargs.get("mode") == "get"
            return '{"level_percent": 37, "muted": false}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)

    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="كم نسبة الصوت", session_key="s1"
    )
    assert handled is True
    assert "37" in str(reply)


@pytest.mark.asyncio
async def test_global_fastpath_destructive_confirmation_flow(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    calls: list[tuple[str, dict]] = []

    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            calls.append((action, kwargs))
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)

    loop = AgentLoop()
    first_handled, first_reply = await loop._try_global_windows_fastpath(
        text="shutdown now", session_key="s2"
    )
    assert first_handled is True
    assert "destructive" in str(first_reply).lower() or "خط" in str(first_reply)
    assert calls == []

    second_handled, second_reply = await loop._try_global_windows_fastpath(
        text="yes", session_key="s2"
    )
    assert second_handled is True
    assert calls
    assert isinstance(second_reply, str)


@pytest.mark.asyncio
async def test_global_fastpath_brightness_get(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "brightness"
            assert kwargs.get("mode") == "get"
            return '{"brightness_percent": 62}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="كم نسبة الاضاءة", session_key="s3"
    )
    assert handled is True
    assert "62" in str(reply)


@pytest.mark.asyncio
async def test_global_fastpath_battery_get(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "system_info"
            assert kwargs.get("mode") == "battery"
            return '{"available": true, "percent": 81, "plugged": false}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="كم نسبة البطارية", session_key="s4"
    )
    assert handled is True
    assert "81" in str(reply)
