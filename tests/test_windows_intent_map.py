import pytest

from Mudabbir.tools.capabilities.windows_intent_map import (
    is_confirmation_message,
    resolve_windows_intent,
)


def test_resolve_arabic_volume_set() -> None:
    result = resolve_windows_intent("خلي الصوت 33%")
    assert result.matched is True
    assert result.action == "volume"
    assert result.params.get("mode") == "set"
    assert result.params.get("level") == 33


def test_resolve_shutdown_is_destructive() -> None:
    result = resolve_windows_intent("shutdown the pc now")
    assert result.matched is True
    assert result.action == "system_power"
    assert result.params.get("mode") == "shutdown"
    assert result.risk_level == "destructive"


def test_resolve_unsupported_audio_output() -> None:
    result = resolve_windows_intent("change audio output to headset")
    assert result.matched is True
    assert result.unsupported is True
    assert "not implemented" in result.unsupported_reason.lower()


def test_confirmation_message_detection() -> None:
    assert is_confirmation_message("yes") is True
    assert is_confirmation_message("نعم نفذ") is True
    assert is_confirmation_message("cancel") is False


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("تفعيل وضع الطيران", "power_user_tools", "airplane_on"),
        ("تعطيل وضع الطيران", "power_user_tools", "airplane_off"),
        ("اعطني نسبة البطارية", "system_info", "battery"),
        ("خلي خطة الطاقة أداء عالي", "system_power", "power_plan_high"),
        ("افتح البيوس", "system_power", "reboot_bios"),
        ("وضع توسيع الشاشة", "window_control", "display_extend"),
        ("وضع تكرار الشاشة", "window_control", "display_duplicate"),
        ("افتح سجل الحافظة", "clipboard_tools", "history"),
        ("اظهار الملفات المخفية", "file_tools", "show_hidden"),
        ("اخفاء الملفات المخفية", "file_tools", "hide_hidden"),
        ("افتح اعدادات الشبكه", "open_settings_page", None),
        ("افتح إعدادات الخصوصية", "open_settings_page", None),
        ("اغلاق كل البرامج المفتوحة", "app_tools", "close_all_apps"),
        ("فحص حالة القرص الصلب", "disk_tools", "smart_status"),
        ("مفاتيح الاختصار المتاحة", "shell_tools", "list_shortcuts"),
        ("افراغ الرام", "maintenance_tools", "empty_ram"),
    ],
)
def test_resolve_new_capabilities(message: str, action: str, mode: str | None) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    if mode is not None:
        assert result.params.get("mode") == mode
    else:
        if "الخصوصية" in message:
            assert result.params.get("page") == "privacy"
        else:
            assert result.params.get("page") == "network"


def test_resolve_rename_pc_extracts_name() -> None:
    result = resolve_windows_intent("تغيير اسم الكمبيوتر إلى OFFICE-DEV")
    assert result.matched is True
    assert result.action == "system_power"
    assert result.params.get("mode") == "rename_pc"
    assert result.params.get("name") == "OFFICE-DEV"


def test_resolve_drag_drop_extracts_coordinates() -> None:
    result = resolve_windows_intent("drag and drop from 100 200 to 400 500")
    assert result.matched is True
    assert result.action == "automation_tools"
    assert result.params.get("mode") == "drag_drop"
    assert result.params.get("x") == 100
    assert result.params.get("y") == 200
    assert result.params.get("x2") == 400
    assert result.params.get("y2") == 500


def test_resolve_window_rename_title() -> None:
    result = resolve_windows_intent("rename window title to 'Focus Session'")
    assert result.matched is True
    assert result.action == "window_control"
    assert result.params.get("mode") == "rename_title"
    assert result.params.get("text") == "Focus Session"


def test_resolve_repeat_key_with_count() -> None:
    result = resolve_windows_intent("repeat key enter 7 times")
    assert result.matched is True
    assert result.action == "automation_tools"
    assert result.params.get("mode") == "repeat_key"
    assert result.params.get("key") == "enter"
    assert result.params.get("repeat_count") == 7


def test_resolve_type_current_date_and_time() -> None:
    date_result = resolve_windows_intent("type current date")
    time_result = resolve_windows_intent("type current time")
    assert date_result.matched is True
    assert time_result.matched is True
    assert date_result.action == "type_text"
    assert time_result.action == "type_text"
    assert isinstance(date_result.params.get("text"), str)
    assert isinstance(time_result.params.get("text"), str)
    assert len(date_result.params.get("text") or "") >= 8
    assert len(time_result.params.get("text") or "") >= 5


def test_resolve_unsupported_display_resolution() -> None:
    result = resolve_windows_intent("تغيير دقة الشاشة 1920x1080")
    assert result.matched is True
    assert result.unsupported is True
    assert "not implemented" in result.unsupported_reason.lower()


def test_resolve_arabic_brightness_and_battery_questions() -> None:
    bright = resolve_windows_intent("كم نسبة الاضاءة")
    battery = resolve_windows_intent("كم نسبة البطارية")
    assert bright.matched is True
    assert bright.action == "brightness"
    assert bright.params.get("mode") == "get"
    assert battery.matched is True
    assert battery.action == "system_info"
    assert battery.params.get("mode") == "battery"


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("مدير المهام", "app_tools", "open_task_manager"),
        ("لوحة التحكم", "app_tools", "open_control_panel"),
        ("متجر ميكروسوفت", "app_tools", "open_store"),
        ("الاعدادات السريعة", "shell_tools", "quick_settings"),
        ("مركز الاشعارات", "shell_tools", "notifications"),
        ("win+a", "shell_tools", "quick_settings"),
        ("win+n", "shell_tools", "notifications"),
        ("win+s", "shell_tools", "search"),
        ("win+r", "shell_tools", "run"),
    ],
)
def test_resolve_more_daily_arabic_aliases(message: str, action: str, mode: str) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("شغل الواي فاي", "network_tools", "wifi_on"),
        ("طفي الواي فاي", "network_tools", "wifi_off"),
        ("افراغ dns", "network_tools", "flush_dns"),
        ("افصل النت", "network_tools", "disconnect_current_network"),
        ("شغل وقف", "media_control", "play_pause"),
        ("المقطع التالي", "media_control", "next"),
        ("المقطع السابق", "media_control", "previous"),
        ("تصغير النافذة", "window_control", "minimize"),
        ("تكبير النافذة", "window_control", "maximize"),
        ("استعادة النافذة", "window_control", "restore"),
    ],
)
def test_resolve_more_common_arabic_phrases(message: str, action: str, mode: str) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("شغل خدمة Spooler", "service_tools", "start"),
        ("اعادة تشغيل خدمة Spooler", "service_tools", "restart"),
        ("حالة الجدار الناري", "security_tools", "firewall_status"),
        ("مسح الملفات المفتوحة مؤخرا", "security_tools", "recent_files_clear"),
        ("كشف محاولات الاختراق الفاشلة", "security_tools", "intrusion_summary"),
        ("اعادة تشغيل واجهة الويندوز", "process_tools", "restart_explorer"),
    ],
)
def test_resolve_services_security_and_process_arabic_aliases(
    message: str, action: str, mode: str
) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


def test_resolve_service_name_extraction_start() -> None:
    result = resolve_windows_intent("start service WinRM")
    assert result.matched is True
    assert result.action == "service_tools"
    assert result.params.get("mode") == "start"
    assert result.params.get("name") == "WinRM"


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("قائمة برامج بدء التشغيل", "startup_tools", "startup_list"),
        ("تعطيل برنامج من بدء التشغيل OneDrive", "startup_tools", "disable"),
        ("تفعيل برنامج في بدء التشغيل OneDrive", "startup_tools", "enable"),
        ("فحص امان برامج بدء التشغيل", "startup_tools", "signature_check"),
        ("تعداد التطبيقات المشغلة في الخلفية", "background_tools", "count_background"),
        ("اي تطبيق يستخدم الانترنت", "background_tools", "network_usage_per_app"),
        ("التطبيقات التي تمنع السكون", "background_tools", "wake_lock_apps"),
        ("اكثر 5 تطبيقات تستهلك المعالج", "performance_tools", "top_cpu"),
        ("اجمالي استهلاك الرام", "performance_tools", "total_ram_percent"),
    ],
)
def test_resolve_startup_background_performance_arabic_aliases(
    message: str, action: str, mode: str
) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


def test_resolve_startup_disable_extracts_name() -> None:
    result = resolve_windows_intent("disable startup OneDrive")
    assert result.matched is True
    assert result.action == "startup_tools"
    assert result.params.get("mode") == "disable"
    assert result.params.get("name") == "OneDrive"


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("فتح تبويب جديد", "browser_control", "new_tab"),
        ("اغلاق التبويب الحالي", "browser_control", "close_tab"),
        ("تحديث الصفحة", "browser_control", "reload"),
        ("قائمة المهام المجدولة", "task_tools", "list"),
        ("تشغيل مهمة مجدولة BackupTask", "task_tools", "run"),
        ("قائمة المستخدمين", "user_tools", "list"),
        ("حذف مستخدم TestUser", "user_tools", "delete"),
    ],
)
def test_resolve_browser_task_user_aliases(message: str, action: str, mode: str) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


def test_resolve_user_set_type_extracts_username_and_group() -> None:
    result = resolve_windows_intent("set user type Ahmed admin")
    assert result.matched is True
    assert result.action == "user_tools"
    assert result.params.get("mode") == "set_type"
    assert result.params.get("username") == "Ahmed"
    assert result.params.get("group") == "admin"


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("فحص التحديثات", "update_tools", "check_updates"),
        ("قائمة التحديثات", "update_tools", "list_updates"),
        ("تنظيف ملفات التحديثات القديمة", "update_tools", "winsxs_cleanup"),
        ("قطع اتصال vpn", "remote_tools", "vpn_disconnect"),
        ("تشغيل vpn OfficeVPN", "remote_tools", "vpn_connect"),
        ("تنظيف الملفات المؤقتة", "disk_tools", "temp_files_clean"),
        ("مسح prefetch", "disk_tools", "prefetch_clean"),
        ("استخدام القرص", "disk_tools", "disk_usage"),
        ("registry backup HKCU\\Software", "registry_tools", "backup"),
    ],
)
def test_resolve_update_remote_disk_registry_aliases(
    message: str, action: str, mode: str
) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


def test_resolve_install_kb_extracts_target() -> None:
    result = resolve_windows_intent("install kb KB5034123")
    assert result.matched is True
    assert result.action == "update_tools"
    assert result.params.get("mode") == "install_kb"
    assert result.params.get("target") == "KB5034123"


def test_resolve_registry_set_value_extracts_fields() -> None:
    result = resolve_windows_intent(
        'registry set value "HKCU\\Software\\MyApp" name Theme data Dark dword'
    )
    assert result.matched is True
    assert result.action == "registry_tools"
    assert result.params.get("mode") == "set_value"
    assert "HKCU\\Software\\MyApp" in str(result.params.get("key"))
    assert result.params.get("value_name") == "Theme"
    assert result.params.get("value_data") == "Dark"
    assert result.params.get("value_type") == "REG_DWORD"


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("المنافذ المفتوحة", "network_tools", "open_ports"),
        ("جدول التوجيه", "network_tools", "route_table"),
        ("الاجهزة المتصلة بالشبكة", "network_tools", "net_scan"),
        ("تشغيل مشاركة الملفات", "network_tools", "file_sharing_on"),
        ("المجلدات المشاركة", "network_tools", "shared_folders"),
        ("اغلاق منفذ 445", "security_tools", "block_port"),
        ("تعطيل منافذ usb", "security_tools", "disable_usb"),
        ("تفعيل الكاميرا", "security_tools", "enable_camera"),
        ("البحث عن نص داخل الملفات error", "search_tools", "search_text"),
        ("ملفات اكبر من 500", "search_tools", "files_larger_than"),
        ("ملفات تم تعديلها اليوم", "search_tools", "modified_today"),
        ("ايجاد جميع الصور", "search_tools", "find_images"),
        ("احصاء عدد الملفات", "search_tools", "count_files"),
        ("افتح رابط https://example.com", "web_tools", "open_url"),
        ("تحميل ملف من رابط https://example.com/a.zip", "web_tools", "download_file"),
        ("حالة الطقس مدينة Amman", "web_tools", "weather"),
        ("اسعار العملات", "api_tools", "currency"),
        ("ترجمة نص hello world", "api_tools", "translate_quick"),
    ],
)
def test_resolve_network_security_search_web_api_aliases(
    message: str, action: str, mode: str
) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


def test_resolve_port_owner_extracts_port() -> None:
    result = resolve_windows_intent("من يستخدم المنفذ 3389")
    assert result.matched is True
    assert result.action == "network_tools"
    assert result.params.get("mode") == "port_owner"
    assert result.params.get("port") == 3389


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("مسح الكاش لمتصفح chrome", "browser_deep_tools", "clear_chrome_cache"),
        ("مسح الكاش لمتصفح edge", "browser_deep_tools", "clear_edge_cache"),
        ("فتح مجموعة روابط https://a.com https://b.com", "browser_deep_tools", "multi_open"),
        ("فتح ملف word جديد", "office_tools", "open_word_new"),
        ("قائمة التعريفات المثبتة", "driver_tools", "drivers_list"),
        ("اخذ نسخة احتياطية من التعريفات", "driver_tools", "drivers_backup"),
        ("التعريفات التي فيها مشاكل", "driver_tools", "drivers_issues"),
        ("مفتاح تفعيل الويندوز", "info_tools", "windows_product_key"),
        ("موديل اللابتوب", "info_tools", "model_info"),
        ("لغة النظام الحالية", "info_tools", "system_language"),
        ("تاريخ تثبيت الويندوز", "info_tools", "windows_install_date"),
        ("سرعة استجابة الشاشة", "info_tools", "refresh_rate"),
    ],
)
def test_resolve_browserdeep_office_driver_info_aliases(
    message: str, action: str, mode: str
) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


@pytest.mark.parametrize(
    ("message", "action", "mode"),
    [
        ("control /name Microsoft.System", "app_tools", "open_control_panel"),
        ("control /name Microsoft.NetworkAndSharingCenter", "app_tools", "open_control_panel"),
        ("appwiz.cpl", "app_tools", "open_add_remove_programs"),
        ("افتح mmsys.cpl", "app_tools", "open_sound_cpl"),
        ("اتصالات الشبكة", "app_tools", "open_network_connections"),
        ("الوقت والتاريخ", "app_tools", "open_time_date"),
        ("خصائص النظام", "app_tools", "open_system_properties"),
        ("powercfg.cpl", "app_tools", "open_power_options"),
        ("firewall.cpl", "app_tools", "open_firewall_cpl"),
        ("devmgmt.msc", "dev_tools", "open_device_manager"),
        ("diskmgmt.msc", "dev_tools", "open_disk_management"),
        ("eventvwr.msc", "dev_tools", "open_event_viewer"),
        ("services.msc", "dev_tools", "open_services"),
        ("taskschd.msc", "dev_tools", "open_task_scheduler"),
        ("compmgmt.msc", "dev_tools", "open_computer_management"),
        ("lusrmgr.msc", "dev_tools", "open_local_users_groups"),
        ("secpol.msc", "dev_tools", "open_local_security_policy"),
        ("printmanagement.msc", "dev_tools", "open_print_management"),
    ],
)
def test_resolve_control_panel_and_mmc_aliases(message: str, action: str, mode: str) -> None:
    result = resolve_windows_intent(message)
    assert result.matched is True
    assert result.action == action
    assert result.params.get("mode") == mode


def test_resolve_docx_to_pdf_extracts_paths() -> None:
    result = resolve_windows_intent(
        'تحويل docx الى pdf "C:\\tmp\\a.docx" "C:\\tmp\\a.pdf"'
    )
    assert result.matched is True
    assert result.action == "office_tools"
    assert result.params.get("mode") == "docx_to_pdf"
    assert str(result.params.get("path", "")).endswith("a.docx")
    assert str(result.params.get("target", "")).endswith("a.pdf")
