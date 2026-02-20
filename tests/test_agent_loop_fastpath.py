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


@pytest.mark.asyncio
async def test_global_fastpath_browser_task_user_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "browser_control":
                assert kwargs.get("mode") == "new_tab"
                return '{"ok": true}'
            if action == "task_tools":
                assert kwargs.get("mode") == "run"
                assert kwargs.get("name") == "BackupTask"
                return '{"ok": true}'
            if action == "user_tools":
                assert kwargs.get("mode") == "delete"
                assert kwargs.get("username") == "TestUser"
                return '{"ok": true}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_browser, reply_browser = await loop._try_global_windows_fastpath(
        text="فتح تبويب جديد", session_key="s16"
    )
    assert handled_browser is True
    assert "تبويب" in str(reply_browser)

    handled_task, reply_task = await loop._try_global_windows_fastpath(
        text="تشغيل مهمة مجدولة BackupTask", session_key="s17"
    )
    assert handled_task is True
    assert "BackupTask" in str(reply_task)

    handled_user, reply_user = await loop._try_global_windows_fastpath(
        text="حذف مستخدم TestUser", session_key="s18"
    )
    assert handled_user is True
    assert "خط" in str(reply_user) or "yes" in str(reply_user).lower()
    handled_user_confirm, reply_user_confirm = await loop._try_global_windows_fastpath(
        text="نعم", session_key="s18"
    )
    assert handled_user_confirm is True
    assert "TestUser" in str(reply_user_confirm)


@pytest.mark.asyncio
async def test_global_fastpath_update_remote_disk_registry_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "update_tools":
                assert kwargs.get("mode") == "check_updates"
                return '{"ok": true, "mode": "check_updates"}'
            if action == "remote_tools":
                assert kwargs.get("mode") == "vpn_connect"
                assert kwargs.get("host") == "OfficeVPN"
                return '{"ok": true, "mode": "vpn_connect"}'
            if action == "disk_tools":
                assert kwargs.get("mode") == "temp_files_clean"
                return '{"ok": true, "mode": "temp_files_clean"}'
            if action == "registry_tools":
                assert kwargs.get("mode") == "backup"
                return '{"ok": true, "mode": "backup"}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_update, reply_update = await loop._try_global_windows_fastpath(
        text="فحص التحديثات", session_key="s19"
    )
    assert handled_update is True
    assert "تحديث" in str(reply_update)

    handled_remote, reply_remote = await loop._try_global_windows_fastpath(
        text="تشغيل vpn OfficeVPN", session_key="s20"
    )
    assert handled_remote is True
    assert "OfficeVPN" in str(reply_remote)

    handled_disk, reply_disk = await loop._try_global_windows_fastpath(
        text="تنظيف الملفات المؤقتة", session_key="s21"
    )
    assert handled_disk is True
    assert "المؤقتة" in str(reply_disk) or "temp" in str(reply_disk).lower()

    handled_reg, reply_reg = await loop._try_global_windows_fastpath(
        text="registry backup HKCU\\Software", session_key="s22"
    )
    assert handled_reg is True
    assert "سجل" in str(reply_reg) or "registry" in str(reply_reg).lower()


@pytest.mark.asyncio
async def test_global_fastpath_network_security_search_web_api_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "network_tools":
                assert kwargs.get("mode") == "open_ports"
                return '{"ok": true}'
            if action == "security_tools":
                assert kwargs.get("mode") == "disable_usb"
                return '{"ok": true}'
            if action == "search_tools":
                assert kwargs.get("mode") == "find_images"
                return '{"ok": true, "count": 12}'
            if action == "web_tools":
                assert kwargs.get("mode") == "open_url"
                return '{"ok": true}'
            if action == "api_tools":
                assert kwargs.get("mode") == "currency"
                return '{"ok": true}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_network, reply_network = await loop._try_global_windows_fastpath(
        text="المنافذ المفتوحة", session_key="s23"
    )
    assert handled_network is True
    assert "منافذ" in str(reply_network) or "ports" in str(reply_network).lower()

    handled_security, reply_security = await loop._try_global_windows_fastpath(
        text="تعطيل منافذ usb", session_key="s24"
    )
    assert handled_security is True
    assert "خط" in str(reply_security) or "destructive" in str(reply_security).lower()
    handled_security_confirm, reply_security_confirm = await loop._try_global_windows_fastpath(
        text="نعم", session_key="s24"
    )
    assert handled_security_confirm is True
    assert "usb" in str(reply_security_confirm).lower() or "منافذ" in str(reply_security_confirm)

    handled_search, reply_search = await loop._try_global_windows_fastpath(
        text="ايجاد جميع الصور", session_key="s25"
    )
    assert handled_search is True
    assert "الصور" in str(reply_search) or "image" in str(reply_search).lower()

    handled_web, reply_web = await loop._try_global_windows_fastpath(
        text="افتح رابط https://example.com", session_key="s26"
    )
    assert handled_web is True
    assert "رابط" in str(reply_web) or "url" in str(reply_web).lower()

    handled_api, reply_api = await loop._try_global_windows_fastpath(
        text="اسعار العملات", session_key="s27"
    )
    assert handled_api is True
    assert "عملات" in str(reply_api) or "currency" in str(reply_api).lower()


@pytest.mark.asyncio
async def test_global_fastpath_browserdeep_office_driver_info_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "browser_deep_tools":
                assert kwargs.get("mode") == "clear_chrome_cache"
                return '{"ok": true}'
            if action == "office_tools":
                assert kwargs.get("mode") == "open_word_new"
                return '{"ok": true}'
            if action == "driver_tools":
                assert kwargs.get("mode") == "drivers_list"
                return '{"ok": true}'
            if action == "info_tools":
                assert kwargs.get("mode") == "system_language"
                return '{"ok": true}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_deep, reply_deep = await loop._try_global_windows_fastpath(
        text="مسح الكاش لمتصفح chrome", session_key="s28"
    )
    assert handled_deep is True
    assert "chrome" in str(reply_deep).lower() or "كاش" in str(reply_deep)

    handled_office, reply_office = await loop._try_global_windows_fastpath(
        text="فتح ملف word جديد", session_key="s29"
    )
    assert handled_office is True
    assert "word" in str(reply_office).lower() or "ورد" in str(reply_office)

    handled_driver, reply_driver = await loop._try_global_windows_fastpath(
        text="قائمة التعريفات المثبتة", session_key="s30"
    )
    assert handled_driver is True
    assert "تعريفات" in str(reply_driver) or "driver" in str(reply_driver).lower()

    handled_info, reply_info = await loop._try_global_windows_fastpath(
        text="لغة النظام الحالية", session_key="s31"
    )
    assert handled_info is True
    assert "لغة" in str(reply_info) or "language" in str(reply_info).lower()
