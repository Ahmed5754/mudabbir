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
async def test_global_fastpath_network_diagnostics_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "network_tools"
            assert kwargs.get("mode") in {"ipconfig_all", "tracert", "pathping", "nslookup", "netstat_active", "display_dns", "getmac", "arp_table", "nbtstat_cache", "nbtstat_host", "net_view"}
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_ipall, reply_ipall = await loop._try_global_windows_fastpath(
        text="ipconfig /all", session_key="s7i"
    )
    assert handled_ipall is True
    assert "شبك" in str(reply_ipall) or "network" in str(reply_ipall).lower()

    handled_tracert, reply_tracert = await loop._try_global_windows_fastpath(
        text="تتبع المسار google.com", session_key="s7t"
    )
    assert handled_tracert is True
    assert "تتبع" in str(reply_tracert) or "trace" in str(reply_tracert).lower()

    handled_pathping, reply_pathping = await loop._try_global_windows_fastpath(
        text="pathping google.com", session_key="s7p"
    )
    assert handled_pathping is True
    assert "فقدان" in str(reply_pathping) or "path ping" in str(reply_pathping).lower()

    handled_nslookup, reply_nslookup = await loop._try_global_windows_fastpath(
        text="nslookup google.com", session_key="s7n"
    )
    assert handled_nslookup is True
    assert "dns" in str(reply_nslookup).lower() or "استعلام" in str(reply_nslookup)

    handled_netstat, reply_netstat = await loop._try_global_windows_fastpath(
        text="الاتصالات النشطة", session_key="s7s"
    )
    assert handled_netstat is True
    assert "نشط" in str(reply_netstat) or "active" in str(reply_netstat).lower()

    handled_dns, reply_dns = await loop._try_global_windows_fastpath(
        text="عرض dns", session_key="s7d"
    )
    assert handled_dns is True
    assert "dns" in str(reply_dns).lower()

    handled_getmac, reply_getmac = await loop._try_global_windows_fastpath(
        text="getmac", session_key="s7m"
    )
    assert handled_getmac is True
    assert "mac" in str(reply_getmac).lower()

    handled_arp, reply_arp = await loop._try_global_windows_fastpath(
        text="جدول arp", session_key="s7a"
    )
    assert handled_arp is True
    assert "arp" in str(reply_arp).lower()

    handled_nbt, reply_nbt = await loop._try_global_windows_fastpath(
        text="nbtstat -c", session_key="s7b"
    )
    assert handled_nbt is True
    assert "netbios" in str(reply_nbt).lower() or "bios" in str(reply_nbt).lower()

    handled_nbthost, reply_nbthost = await loop._try_global_windows_fastpath(
        text="nbtstat -a FILESRV", session_key="s7bh"
    )
    assert handled_nbthost is True
    assert "netbios" in str(reply_nbthost).lower() or "bios" in str(reply_nbthost).lower()

    handled_netview, reply_netview = await loop._try_global_windows_fastpath(
        text="net view", session_key="s7v"
    )
    assert handled_netview is True
    assert "شبك" in str(reply_netview) or "network" in str(reply_netview).lower()


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
async def test_global_fastpath_service_list_and_dependencies_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "service_tools"
            mode = kwargs.get("mode")
            if mode == "list":
                return '{"ok": true, "mode":"list", "data":[]}'
            if mode == "dependencies":
                assert kwargs.get("name") == "Spooler"
                return '{"ok": true, "mode":"dependencies", "data":{}}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_list, reply_list = await loop._try_global_windows_fastpath(
        text="قائمة الخدمات", session_key="s12a"
    )
    assert handled_list is True
    assert "الخدمات" in str(reply_list) or "services" in str(reply_list).lower()

    handled_deps, reply_deps = await loop._try_global_windows_fastpath(
        text="تبعيات الخدمة Spooler", session_key="s12b"
    )
    assert handled_deps is True
    assert "تبعيات" in str(reply_deps) or "dependencies" in str(reply_deps).lower()


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
                assert kwargs.get("mode") in {
                    "new_tab",
                    "reopen_tab",
                    "history",
                    "downloads",
                    "zoom_in",
                    "zoom_out",
                    "zoom_reset",
                    "save_pdf",
                }
                return '{"ok": true}'
            if action == "task_tools":
                assert kwargs.get("mode") in {"running", "last_run", "run", "end", "enable", "disable"}
                if kwargs.get("mode") in {"run", "end", "enable", "disable"}:
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

    handled_reopen, reply_reopen = await loop._try_global_windows_fastpath(
        text="إعادة فتح التبويب المغلق", session_key="s16r"
    )
    assert handled_reopen is True
    assert "تبويب" in str(reply_reopen) or "tab" in str(reply_reopen).lower()

    handled_history, reply_history = await loop._try_global_windows_fastpath(
        text="سجل التصفح", session_key="s16h"
    )
    assert handled_history is True
    assert "سجل" in str(reply_history) or "history" in str(reply_history).lower()

    handled_downloads, reply_downloads = await loop._try_global_windows_fastpath(
        text="تنزيلات المتصفح", session_key="s16d"
    )
    assert handled_downloads is True
    assert "تنزيل" in str(reply_downloads) or "download" in str(reply_downloads).lower()

    handled_zoom_in, reply_zoom_in = await loop._try_global_windows_fastpath(
        text="تكبير الصفحة", session_key="s16zi"
    )
    assert handled_zoom_in is True
    assert "تكبير" in str(reply_zoom_in) or "zoom" in str(reply_zoom_in).lower()

    handled_zoom_out, reply_zoom_out = await loop._try_global_windows_fastpath(
        text="تصغير الصفحة", session_key="s16zo"
    )
    assert handled_zoom_out is True
    assert "تصغير" in str(reply_zoom_out) or "zoom" in str(reply_zoom_out).lower()

    handled_zoom_reset, reply_zoom_reset = await loop._try_global_windows_fastpath(
        text="ارجاع الزوم 100", session_key="s16zr"
    )
    assert handled_zoom_reset is True
    assert "100" in str(reply_zoom_reset) or "zoom" in str(reply_zoom_reset).lower()

    handled_save_pdf, reply_save_pdf = await loop._try_global_windows_fastpath(
        text="حفظ الصفحة pdf", session_key="s16p"
    )
    assert handled_save_pdf is True
    assert "pdf" in str(reply_save_pdf).lower() or "صفحة" in str(reply_save_pdf)

    handled_task, reply_task = await loop._try_global_windows_fastpath(
        text="تشغيل مهمة مجدولة BackupTask", session_key="s17"
    )
    assert handled_task is True
    assert "BackupTask" in str(reply_task)

    handled_task_running, reply_task_running = await loop._try_global_windows_fastpath(
        text="المهام المجدولة الجارية", session_key="s17r"
    )
    assert handled_task_running is True
    assert "المهام" in str(reply_task_running) or "tasks" in str(reply_task_running).lower()

    handled_task_last_run, reply_task_last_run = await loop._try_global_windows_fastpath(
        text="آخر تشغيل للمهام", session_key="s17lr"
    )
    assert handled_task_last_run is True
    assert "تشغيل" in str(reply_task_last_run) or "last run" in str(reply_task_last_run).lower()

    handled_task_end, reply_task_end = await loop._try_global_windows_fastpath(
        text="إنهاء مهمة مجدولة BackupTask", session_key="s17e"
    )
    assert handled_task_end is True
    assert "BackupTask" in str(reply_task_end)

    handled_task_enable, reply_task_enable = await loop._try_global_windows_fastpath(
        text="تمكين مهمة مجدولة BackupTask", session_key="s17en"
    )
    assert handled_task_enable is True
    assert "BackupTask" in str(reply_task_enable)

    handled_task_disable, reply_task_disable = await loop._try_global_windows_fastpath(
        text="تعطيل مهمة مجدولة BackupTask", session_key="s17d"
    )
    assert handled_task_disable is True
    assert "BackupTask" in str(reply_task_disable)

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


@pytest.mark.asyncio
async def test_global_fastpath_control_panel_and_mmc_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            if action == "app_tools":
                assert kwargs.get("mode") == "open_sound_cpl"
                return '{"ok": true}'
            if action == "dev_tools":
                assert kwargs.get("mode") == "open_task_scheduler"
                return '{"ok": true}'
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_app, reply_app = await loop._try_global_windows_fastpath(
        text="افتح mmsys.cpl", session_key="s32"
    )
    assert handled_app is True
    assert "الصوت" in str(reply_app) or "sound" in str(reply_app).lower()

    handled_dev, reply_dev = await loop._try_global_windows_fastpath(
        text="taskschd.msc", session_key="s33"
    )
    assert handled_dev is True
    assert "المهام" in str(reply_dev) or "scheduler" in str(reply_dev).lower()


@pytest.mark.asyncio
async def test_global_fastpath_event_errors_and_bsod_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "dev_tools"
            assert kwargs.get("mode") in {"event_errors", "analyze_bsod"}
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_events, reply_events = await loop._try_global_windows_fastpath(
        text="آخر أخطاء النظام", session_key="s33e"
    )
    assert handled_events is True
    assert "خط" in str(reply_events) or "event" in str(reply_events).lower()

    handled_bsod, reply_bsod = await loop._try_global_windows_fastpath(
        text="تحليل شاشة الموت", session_key="s33b"
    )
    assert handled_bsod is True
    assert "bsod" in str(reply_bsod).lower() or "انهيار" in str(reply_bsod)


@pytest.mark.asyncio
async def test_global_fastpath_app_tools_extra_replies(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "app_tools"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()

    handled_mixer, reply_mixer = await loop._try_global_windows_fastpath(
        text="volume mixer", session_key="s34"
    )
    assert handled_mixer is True
    assert "mixer" in str(reply_mixer).lower() or "خالط" in str(reply_mixer)

    handled_mic, reply_mic = await loop._try_global_windows_fastpath(
        text="microphone settings", session_key="s35"
    )
    assert handled_mic is True
    assert "microphone" in str(reply_mic).lower() or "الميكروفون" in str(reply_mic)

    handled_inetcpl, reply_inetcpl = await loop._try_global_windows_fastpath(
        text="inetcpl.cpl", session_key="s35i"
    )
    assert handled_inetcpl is True
    assert "internet" in str(reply_inetcpl).lower() or "الإنترنت" in str(reply_inetcpl) or "الانترنت" in str(reply_inetcpl)

    handled_deskcpl, reply_deskcpl = await loop._try_global_windows_fastpath(
        text="desk.cpl", session_key="s35d"
    )
    assert handled_deskcpl is True
    assert "display" in str(reply_deskcpl).lower() or "العرض" in str(reply_deskcpl)

    handled_appwiz, reply_appwiz = await loop._try_global_windows_fastpath(
        text="appwiz.cpl", session_key="s36"
    )
    assert handled_appwiz is True
    assert "programs" in str(reply_appwiz).lower() or "البرامج" in str(reply_appwiz)

    handled_mouse, reply_mouse = await loop._try_global_windows_fastpath(
        text="control mouse", session_key="s36m"
    )
    assert handled_mouse is True
    assert "mouse" in str(reply_mouse).lower() or "فأرة" in str(reply_mouse) or "الفاره" in str(reply_mouse)

    handled_keyboard, reply_keyboard = await loop._try_global_windows_fastpath(
        text="control keyboard", session_key="s36k"
    )
    assert handled_keyboard is True
    assert "keyboard" in str(reply_keyboard).lower() or "لوحة المفاتيح" in str(reply_keyboard)

    handled_printers, reply_printers = await loop._try_global_windows_fastpath(
        text="control printers", session_key="s36p"
    )
    assert handled_printers is True
    assert "printer" in str(reply_printers).lower() or "طابعات" in str(reply_printers)

    handled_users, reply_users = await loop._try_global_windows_fastpath(
        text="control userpasswords2", session_key="s36u"
    )
    assert handled_users is True
    assert "user" in str(reply_users).lower() or "حساب" in str(reply_users)

    handled_users2, reply_users2 = await loop._try_global_windows_fastpath(
        text="control userpasswords", session_key="s36u2"
    )
    assert handled_users2 is True
    assert "user" in str(reply_users2).lower() or "حساب" in str(reply_users2)

    handled_bt, reply_bt = await loop._try_global_windows_fastpath(
        text="bthprops.cpl", session_key="s36bt"
    )
    assert handled_bt is True
    assert "bluetooth" in str(reply_bt).lower() or "بلوتوث" in str(reply_bt)

    handled_access, reply_access = await loop._try_global_windows_fastpath(
        text="access.cpl", session_key="s36ac"
    )
    assert handled_access is True
    assert "access" in str(reply_access).lower() or "الوصول" in str(reply_access)

    handled_admintools, reply_admintools = await loop._try_global_windows_fastpath(
        text="control admintools", session_key="s36a"
    )
    assert handled_admintools is True
    assert "administrative" in str(reply_admintools).lower() or "اداري" in str(reply_admintools) or "إداري" in str(reply_admintools)

    handled_schedtasks, reply_schedtasks = await loop._try_global_windows_fastpath(
        text="control schedtasks", session_key="s36s"
    )
    assert handled_schedtasks is True
    assert "task" in str(reply_schedtasks).lower() or "المهام" in str(reply_schedtasks)

    handled_netconn, reply_netconn = await loop._try_global_windows_fastpath(
        text="control netconnections", session_key="s36n"
    )
    assert handled_netconn is True
    assert "network" in str(reply_netconn).lower() or "الشبكة" in str(reply_netconn)

    handled_folders, reply_folders = await loop._try_global_windows_fastpath(
        text="control folders", session_key="s36f"
    )
    assert handled_folders is True
    assert "folder" in str(reply_folders).lower() or "مجلد" in str(reply_folders)

    handled_color, reply_color = await loop._try_global_windows_fastpath(
        text="control color", session_key="s36c"
    )
    assert handled_color is True
    assert "color" in str(reply_color).lower() or "لون" in str(reply_color) or "ألوان" in str(reply_color)

    handled_desktop, reply_desktop = await loop._try_global_windows_fastpath(
        text="control desktop", session_key="s36d"
    )
    assert handled_desktop is True
    assert "desktop" in str(reply_desktop).lower() or "سطح المكتب" in str(reply_desktop)


@pytest.mark.asyncio
async def test_global_fastpath_open_settings_page_reply(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "open_settings_page"
            assert kwargs.get("page") == "privacy"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="افتح إعدادات الخصوصية", session_key="s37"
    )
    assert handled is True
    assert "الخصوصية" in str(reply) or "privacy" in str(reply).lower()


@pytest.mark.asyncio
async def test_global_fastpath_open_settings_page_update_reply(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    class DummyDesktopTool:
        async def execute(self, action: str, **kwargs):
            assert action == "open_settings_page"
            assert kwargs.get("page") == "windowsupdate"
            return '{"ok": true}'

    monkeypatch.setattr("Mudabbir.tools.builtin.desktop.DesktopTool", DummyDesktopTool)
    loop = AgentLoop()
    handled, reply = await loop._try_global_windows_fastpath(
        text="افتح تحديثات ويندوز", session_key="s38"
    )
    assert handled is True
    assert "تحديثات" in str(reply) or "update" in str(reply).lower()
