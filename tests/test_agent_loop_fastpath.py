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


@pytest.mark.asyncio
async def test_global_fastpath_app_tools_human_reply(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "app_tools"
            assert kwargs.get("mode") == "open_task_manager"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="افتح مدير المهام", session_key="s5"
    )
    assert handled is True
    assert "مدير المهام" in str(reply)


@pytest.mark.asyncio
async def test_global_fastpath_shell_tools_human_reply(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "shell_tools"
            assert kwargs.get("mode") == "quick_settings"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="افتح الاعدادات السريعة", session_key="s6"
    )
    assert handled is True
    assert "السريعة" in str(reply)


@pytest.mark.asyncio
async def test_global_fastpath_network_tools_human_reply(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "network_tools"
            assert kwargs.get("mode") == "wifi_on"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="شغل الواي فاي", session_key="s7"
    )
    assert handled is True
    assert "الواي فاي" in str(reply)


@pytest.mark.asyncio
async def test_global_fastpath_media_and_window_replies(monkeypatch: pytest.MonkeyPatch) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "media_control":
                assert kwargs.get("mode") == "next"
            if action == "window_control":
                assert kwargs.get("mode") == "minimize"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_media, reply_media = await loop._try_global_windows_fastpath(
        text="المقطع التالي", session_key="s8"
    )
    assert handled_media is True
    assert "التالي" in str(reply_media)

    handled_window, reply_window = await loop._try_global_windows_fastpath(
        text="تصغير النافذة", session_key="s9"
    )
    assert handled_window is True
    assert "تصغير" in str(reply_window)


@pytest.mark.asyncio
async def test_global_fastpath_process_top_cpu_human_reply(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "process_tools"
            assert kwargs.get("mode") == "top_cpu"
            return '{"ok": true, "mode": "top_cpu", "items": [{"pid": 1234, "name": "chrome.exe", "cpu": 67.5}]}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="اكثر العمليات استهلاكا للمعالج", session_key="s10"
    )
    assert handled is True
    assert "chrome.exe" in str(reply)
    assert "1234" in str(reply)


@pytest.mark.asyncio
async def test_global_fastpath_service_and_security_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "service_tools":
                assert kwargs.get("mode") == "restart"
                assert kwargs.get("name") == "Spooler"
                return '{"ok": true}'
            if action == "security_tools":
                assert kwargs.get("mode") == "firewall_status"
                return '{"ok": true, "mode": "firewall_status"}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_service, reply_service = await loop._try_global_windows_fastpath(
        text="اعادة تشغيل خدمة Spooler", session_key="s11"
    )
    assert handled_service is True
    assert "Spooler" in str(reply_service)

    handled_security, reply_security = await loop._try_global_windows_fastpath(
        text="حالة الجدار الناري", session_key="s12"
    )
    assert handled_security is True
    assert "جدار" in str(reply_security)


@pytest.mark.asyncio
async def test_global_fastpath_startup_background_and_performance_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "startup_tools":
                assert kwargs.get("mode") == "startup_list"
                return '{"ok": true, "items": []}'
            if action == "background_tools":
                assert kwargs.get("mode") == "count_background"
                return '{"ok": true, "total": 120, "background": 87}'
            if action == "performance_tools":
                assert kwargs.get("mode") == "total_cpu_percent"
                return '{"ok": true, "percent": 41.2}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_startup, reply_startup = await loop._try_global_windows_fastpath(
        text="قائمة برامج بدء التشغيل", session_key="s13"
    )
    assert handled_startup is True
    assert "بدء التشغيل" in str(reply_startup)

    handled_bg, reply_bg = await loop._try_global_windows_fastpath(
        text="تعداد التطبيقات المشغلة في الخلفية", session_key="s14"
    )
    assert handled_bg is True
    assert "الخلفية" in str(reply_bg)

    handled_perf, reply_perf = await loop._try_global_windows_fastpath(
        text="اجمالي استهلاك المعالج", session_key="s15"
    )
    assert handled_perf is True
    assert "41.2" in str(reply_perf)
