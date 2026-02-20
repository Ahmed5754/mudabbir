"""Deterministic Windows intent mapping for DesktopTool actions."""

from __future__ import annotations

from dataclasses import dataclass, field
import re
from typing import Any


def _normalize_text(text: str) -> str:
    raw = (text or "").lower().strip()
    raw = raw.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    raw = raw.replace("ى", "ي").replace("ة", "ه")
    raw = re.sub(r"[\u0610-\u061a\u064b-\u065f\u0670\u06d6-\u06ed]", "", raw)
    raw = re.sub(r"\s+", " ", raw)
    return raw


def _contains_any(text: str, tokens: tuple[str, ...]) -> bool:
    return any(tok in text for tok in tokens)


def _extract_first_int(text: str) -> int | None:
    m = re.search(r"-?\d+", text)
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def _extract_url(text: str) -> str:
    m = re.search(r"https?://\S+", text, re.IGNORECASE)
    return (m.group(0).strip() if m else "")


def _extract_host(text: str) -> str:
    m = re.search(r"(?:ping|بينق|اختبار اتصال)\s+([a-zA-Z0-9\.\-]+)", text, re.IGNORECASE)
    return (m.group(1).strip() if m else "")


def _extract_minutes(text: str) -> int | None:
    m = re.search(r"(\d+)\s*(?:min|mins|minute|minutes|دقيقه|دقائق|دقيقه)", text, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return _extract_first_int(text)


def _extract_app_query(text: str) -> str:
    cleaned = re.sub(
        r"(?i)\b(open|launch|start|focus|switch|افتح|شغل|ركز|بدل|التبديل|to|الى|إلى|app|application|program|تطبيق|برنامج)\b",
        " ",
        text,
    )
    cleaned = re.sub(r"[\"'`]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:120]


def is_confirmation_message(text: str) -> bool:
    norm = _normalize_text(text)
    positive = (
        "yes",
        "yep",
        "confirm",
        "confirmed",
        "ok",
        "okay",
        "نفذ",
        "نعم",
        "ايوه",
        "اوافق",
        "موافق",
        "تمام",
        "نفذها",
    )
    negative = (
        "no",
        "cancel",
        "stop",
        "لا",
        "الغاء",
        "إلغاء",
        "وقف",
        "تراجع",
    )
    if _contains_any(norm, negative):
        return False
    return _contains_any(norm, positive)


@dataclass(frozen=True)
class IntentRule:
    capability_id: str
    action: str = ""
    mode: str = ""
    risk_level: str = "safe"  # safe | elevated | destructive
    aliases: tuple[str, ...] = ()
    unsupported_reason: str = ""
    params: tuple[str, ...] = ()


@dataclass
class IntentResolution:
    matched: bool
    capability_id: str = ""
    action: str = ""
    params: dict[str, Any] = field(default_factory=dict)
    risk_level: str = "safe"
    unsupported: bool = False
    unsupported_reason: str = ""


RULES: tuple[IntentRule, ...] = (
    IntentRule("system.shutdown", "system_power", "shutdown", "destructive", ("shutdown", "power off", "اطفي", "ايقاف التشغيل")),
    IntentRule("system.restart", "system_power", "restart", "destructive", ("restart", "reboot", "اعاده التشغيل", "إعادة التشغيل")),
    IntentRule("system.lock", "system_power", "lock", "safe", ("lock", "lock screen", "قفل الشاشه", "اقفل الشاشه")),
    IntentRule("system.sleep", "system_power", "sleep", "elevated", ("sleep", "sleep mode", "وضع السكون", "سكون")),
    IntentRule("system.hibernate", "system_power", "hibernate", "elevated", ("hibernate", "hibernation", "وضع السبات", "hibernate mode")),
    IntentRule("system.logoff", "system_power", "logoff", "elevated", ("logoff", "logout", "تسجيل الخروج")),
    IntentRule("system.screen_off", "system_power", "screen_off", "safe", ("screen off", "اغلاق الشاشه", "اطفاء الشاشه")),
    IntentRule("system.airplane_on", "power_user_tools", "airplane_on", "safe", ("airplane mode on", "turn on airplane", "تفعيل وضع الطيران")),
    IntentRule("system.airplane_off", "power_user_tools", "airplane_off", "safe", ("airplane mode off", "turn off airplane", "تعطيل وضع الطيران")),
    IntentRule("system.uptime", "system_info", "uptime", "safe", ("uptime", "وقت التشغيل", "مدة التشغيل")),
    IntentRule("system.windows_version", "system_info", "windows_version", "safe", ("windows version", "اصدار الويندوز", "إصدار الويندوز")),
    IntentRule("system.about", "system_info", "about", "safe", ("about device", "about page", "حول الجهاز")),
    IntentRule("system.schedule_shutdown", "shutdown_schedule", "set", "destructive", ("schedule shutdown", "shutdown after", "جدوله ايقاف", "جدولة ايقاف")),
    IntentRule("system.cancel_shutdown", "shutdown_schedule", "cancel", "safe", ("cancel shutdown", "الغاء جدوله", "إلغاء جدولة")),
    IntentRule("audio.mute", "volume", "mute", "safe", ("mute", "كتم الصوت", "اكتم الصوت")),
    IntentRule("audio.unmute", "volume", "unmute", "safe", ("unmute", "الغاء الكتم", "إلغاء الكتم")),
    IntentRule("audio.up", "volume", "up", "safe", ("volume up", "raise volume", "رفع الصوت", "اعلي الصوت")),
    IntentRule("audio.down", "volume", "down", "safe", ("volume down", "خفض الصوت", "وطي الصوت")),
    IntentRule("audio.get", "volume", "get", "safe", ("current volume", "volume level", "مستوى الصوت", "نسبه الصوت")),
    IntentRule("audio.play_pause", "media_control", "play_pause", "safe", ("play pause", "pause", "تشغيل", "ايقاف مؤقت")),
    IntentRule("audio.next", "media_control", "next", "safe", ("next song", "next track", "التالي", "الاغنيه التاليه")),
    IntentRule("audio.previous", "media_control", "previous", "safe", ("previous song", "previous track", "السابق", "الاغنيه السابقه")),
    IntentRule("audio.stop", "media_control", "stop", "safe", ("stop media", "ايقاف الوسائط", "stop playback")),
    IntentRule("audio.mixer", "app_tools", "open_volume_mixer", "safe", ("volume mixer", "sndvol", "خلط الصوت", "الميكسار")),
    IntentRule("audio.mic_settings", "app_tools", "open_mic_settings", "safe", ("microphone settings", "mic settings", "اعدادات الميكروفون")),
    IntentRule("audio.set_output", unsupported_reason="Changing audio output is not implemented in DesktopTool yet.", aliases=("audio output", "speaker headset", "تغيير مخرج الصوت")),
    IntentRule("display.brightness_up", "brightness", "up", "safe", ("brightness up", "raise brightness", "رفع الاضاءه", "رفع السطوع")),
    IntentRule("display.brightness_down", "brightness", "down", "safe", ("brightness down", "خفض الاضاءه", "خفض السطوع")),
    IntentRule("display.brightness_get", "brightness", "get", "safe", ("brightness level", "current brightness", "مستوى الاضاءه", "مستوى السطوع")),
    IntentRule("display.night_light_on", "ui_tools", "night_light_on", "safe", ("night light on", "تفعيل الوضع الليلي")),
    IntentRule("display.night_light_off", "ui_tools", "night_light_off", "safe", ("night light off", "تعطيل الوضع الليلي")),
    IntentRule("display.screenshot_full", "screenshot_tools", "full", "safe", ("screenshot", "لقطه شاشه", "سكرين شوت")),
    IntentRule("display.screenshot_window", "screenshot_tools", "window_active", "safe", ("window screenshot", "لقطه نافذه")),
    IntentRule("display.snipping_tool", "screenshot_tools", "snipping_tool", "safe", ("snipping tool", "اداه القص")),
    IntentRule("display.clipboard_clear", "clipboard_tools", "clear", "safe", ("clear clipboard", "مسح الحافظه")),
    IntentRule("display.desktop_icons_show", "ui_tools", "desktop_icons_show", "safe", ("show desktop icons", "اظهار ايقونات سطح المكتب")),
    IntentRule("display.desktop_icons_hide", "ui_tools", "desktop_icons_hide", "safe", ("اخفاء ايقونات سطح المكتب", "hide desktop icons")),
    IntentRule("files.open_documents", "file_tools", "open_documents", "safe", ("open documents", "افتح المستندات")),
    IntentRule("files.open_downloads", "file_tools", "open_downloads", "safe", ("open downloads", "افتح التنزيلات")),
    IntentRule("files.open_pictures", "file_tools", "open_pictures", "safe", ("open pictures", "افتح الصور")),
    IntentRule("files.open_videos", "file_tools", "open_videos", "safe", ("open videos", "افتح الفيديوهات")),
    IntentRule("files.create_folder", "file_tools", "create_folder", "safe", ("create folder", "انشاء مجلد", "إنشاء مجلد"), params=("name",)),
    IntentRule("files.delete", "file_tools", "delete", "destructive", ("delete file", "حذف ملف"), params=("path",)),
    IntentRule("files.delete_permanent", "file_tools", "delete", "destructive", ("permanent delete", "حذف نهائي"), params=("path", "permanent")),
    IntentRule("files.empty_recycle_bin", "file_tools", "empty_recycle_bin", "destructive", ("empty recycle bin", "افراغ سله المهملات")),
    IntentRule("files.copy", "file_tools", "copy", "safe", ("copy file", "نسخ ملف"), params=("path", "target")),
    IntentRule("files.move", "file_tools", "move", "safe", ("move file", "cut file", "قص ملف"), params=("path", "target")),
    IntentRule("files.rename", "file_tools", "rename", "safe", ("rename file", "اعاده تسميه", "إعادة تسمية"), params=("path", "name")),
    IntentRule("files.zip", "file_tools", "zip", "safe", ("zip file", "ضغط ملف"), params=("path", "target")),
    IntentRule("files.unzip", "file_tools", "unzip", "safe", ("unzip file", "فك ضغط"), params=("path", "target")),
    IntentRule("files.search_ext", "file_tools", "search_ext", "safe", ("search extension", "امتداد", "ابحث عن ملف"), params=("ext",)),
    IntentRule("files.folder_size", "file_tools", "folder_size", "safe", ("folder size", "حجم المجلد"), params=("path",)),
    IntentRule("files.open_cmd_here", "file_tools", "open_cmd_here", "safe", ("open cmd here", "فتح المسار في cmd"), params=("path",)),
    IntentRule("files.open_powershell_here", "file_tools", "open_powershell_here", "safe", ("open powershell here", "فتح المسار في powershell"), params=("path",)),
    IntentRule("network.wifi_on", "network_tools", "wifi_on", "safe", ("wifi on", "تشغيل الواي فاي")),
    IntentRule("network.wifi_off", "network_tools", "wifi_off", "safe", ("wifi off", "ايقاف الواي فاي")),
    IntentRule("network.wifi_passwords", "network_tools", "wifi_passwords", "elevated", ("wifi passwords", "كلمات سر الواي فاي")),
    IntentRule("network.ip_internal", "network_tools", "ip_internal", "safe", ("internal ip", "local ip", "ip الداخلي")),
    IntentRule("network.ip_external", "network_tools", "ip_external", "safe", ("external ip", "public ip", "ip الخارجي")),
    IntentRule("network.renew_ip", "network_tools", "renew_ip", "elevated", ("renew ip", "release renew", "تجديد الip")),
    IntentRule("network.flush_dns", "network_tools", "flush_dns", "safe", ("flush dns", "مسح dns")),
    IntentRule("network.ping", "network_tools", "ping", "safe", ("ping", "اختبار اتصال", "بينق"), params=("host",)),
    IntentRule("network.bluetooth_on", "bluetooth_control", "on", "safe", ("bluetooth on", "تشغيل البلوتوث")),
    IntentRule("network.bluetooth_off", "bluetooth_control", "off", "safe", ("bluetooth off", "ايقاف البلوتوث")),
    IntentRule("network.hotspot_on", "network_tools", "hotspot_on", "elevated", ("hotspot on", "تشغيل نقطه الاتصال")),
    IntentRule("network.hotspot_off", "network_tools", "hotspot_off", "elevated", ("hotspot off", "ايقاف نقطه الاتصال")),
    IntentRule("network.disconnect", "network_tools", "disconnect_current_network", "safe", ("disconnect network", "قطع الاتصال بالشبكه")),
    IntentRule("network.connect_named", "network_tools", "connect_wifi", "safe", ("connect wifi", "الاتصال بشبكه"), params=("host",)),
    IntentRule("apps.open_browser", "app_tools", "open_default_browser", "safe", ("open browser", "افتح المتصفح")),
    IntentRule("apps.open_chrome", "app_tools", "open_chrome", "safe", ("open chrome", "افتح كروم")),
    IntentRule("apps.open_notepad", "app_tools", "open_notepad", "safe", ("open notepad", "افتح المفكره")),
    IntentRule("apps.open_calc", "app_tools", "open_calc", "safe", ("open calculator", "افتح الحاسبه")),
    IntentRule("apps.open_paint", "app_tools", "open_paint", "safe", ("open paint", "افتح الرسام")),
    IntentRule("apps.open_task_manager", "app_tools", "open_task_manager", "safe", ("open task manager", "افتح مدير المهام")),
    IntentRule("apps.close_app", "close_app", "", "elevated", ("close app", "اغلاق برنامج", "اغلق التطبيق"), params=("process_name",)),
    IntentRule("apps.close_all", "app_tools", "close_all_apps", "elevated", ("close all apps", "اغلاق كل البرامج")),
    IntentRule("apps.open_control_panel", "app_tools", "open_control_panel", "safe", ("control panel", "لوحه التحكم")),
    IntentRule("apps.open_store", "app_tools", "open_store", "safe", ("microsoft store", "متجر مايكروسوفت")),
    IntentRule("apps.open_registry", "app_tools", "open_registry", "destructive", ("registry editor", "regedit", "محرر السجل")),
    IntentRule("apps.open_add_remove", "app_tools", "open_add_remove_programs", "safe", ("add remove programs", "اضافه او ازاله البرامج")),
    IntentRule("apps.open_camera", "app_tools", "open_camera", "safe", ("open camera", "تشغيل الكاميرا")),
    IntentRule("apps.open_calendar", "app_tools", "open_calendar", "safe", ("open calendar", "فتح التقويم")),
    IntentRule("apps.open_mail", "app_tools", "open_mail", "safe", ("open mail", "فتح البريد")),
    IntentRule("dev.open_cmd_admin", "dev_tools", "open_cmd_admin", "elevated", ("cmd as admin", "فتح cmd كمسؤول")),
    IntentRule("dev.open_powershell_admin", "dev_tools", "open_powershell_admin", "elevated", ("powershell as admin", "فتح powershell كمسؤول")),
    IntentRule("dev.top_cpu", "process_tools", "top_cpu", "safe", ("top cpu", "اكثر العمليات استهلاكا للمعالج")),
    IntentRule("dev.top_ram", "process_tools", "top_ram", "safe", ("top ram", "اكثر العمليات استهلاكا للرام")),
    IntentRule("dev.sfc_scan", "dev_tools", "sfc_scan", "elevated", ("sfc scan", "فحص ملفات النظام")),
    IntentRule("dev.chkdsk", "dev_tools", "chkdsk", "elevated", ("chkdsk", "فحص القرص")),
    IntentRule("dev.disk_management", "dev_tools", "open_disk_management", "safe", ("disk management", "اداره الاقراص")),
    IntentRule("dev.device_manager", "dev_tools", "open_device_manager", "safe", ("device manager", "اداره الاجهزه")),
    IntentRule("dev.perfmon", "dev_tools", "open_perfmon", "safe", ("performance monitor", "مراقب الاداء")),
    IntentRule("dev.event_viewer", "dev_tools", "open_event_viewer", "safe", ("event viewer", "سجل الاحداث")),
    IntentRule("dev.rdp", "remote_tools", "rdp_open", "elevated", ("remote desktop", "تشغيل remote desktop")),
    IntentRule("services.stop", "service_tools", "stop", "destructive", ("stop service", "ايقاف خدمه"), params=("name",)),
    IntentRule("window.minimize", "window_control", "minimize", "safe", ("minimize window", "تصغير النافذه")),
    IntentRule("window.maximize", "window_control", "maximize", "safe", ("maximize window", "تكبير النافذه")),
    IntentRule("window.restore", "window_control", "restore", "safe", ("restore window", "استعاده النافذه")),
    IntentRule("window.close_current", "window_control", "close_current", "safe", ("close current window", "اغلاق النافذه الحاليه")),
    IntentRule("window.show_desktop", "window_control", "show_desktop", "safe", ("show desktop", "تصغير كل النوافذ")),
    IntentRule("window.undo_show_desktop", "window_control", "undo_show_desktop", "safe", ("undo show desktop", "اظهار النوافذ المصغره")),
    IntentRule("window.always_on_top_on", "window_control", "always_on_top_on", "safe", ("always on top", "دائما في المقدمه")),
    IntentRule("window.always_on_top_off", "window_control", "always_on_top_off", "safe", ("remove always on top", "الغاء دائما في المقدمه")),
    IntentRule("window.split_right", "window_control", "split_right", "safe", ("split right", "يمين الشاشه")),
    IntentRule("window.split_left", "window_control", "split_left", "safe", ("split left", "يسار الشاشه")),
    IntentRule("window.move_next_monitor_right", "window_control", "move_next_monitor_right", "safe", ("next monitor", "الشاشه الثانيه")),
    IntentRule("window.alt_tab", "window_control", "alt_tab", "safe", ("alt tab", "تبديل النوافذ")),
    IntentRule("window.task_view", "window_control", "task_view", "safe", ("task view", "عرض المهام")),
    IntentRule("window.transparency", "window_control", "transparency", "elevated", ("window opacity", "شفافيه النافذه", "transparency"), params=("opacity",)),
    IntentRule("mouse.move", "mouse_move", "", "safe", ("move mouse", "حرك الماوس"), params=("x", "y")),
    IntentRule("mouse.click_left", "click", "", "safe", ("left click", "نقره يسار", "ضغطة يسار")),
    IntentRule("mouse.click_right", "click", "", "safe", ("right click", "نقره يمين", "ضغطة يمين")),
    IntentRule("mouse.double_click", "click", "", "safe", ("double click", "نقره مزدوجه", "ضغطة مزدوجة")),
    IntentRule("mouse.scroll_up", "press_key", "", "safe", ("scroll up", "تمرير للاعلى"), params=("key",)),
    IntentRule("mouse.scroll_down", "press_key", "", "safe", ("scroll down", "تمرير للاسفل"), params=("key",)),
    IntentRule("keyboard.type", "type_text", "", "safe", ("type text", "اكتب نص", "كتابه"), params=("text",)),
    IntentRule("keyboard.enter", "press_key", "", "safe", ("press enter", "اضغط enter"), params=("key",)),
    IntentRule("keyboard.space", "press_key", "", "safe", ("press space", "اضغط مسافه"), params=("key",)),
    IntentRule("keyboard.backspace", "press_key", "", "safe", ("press backspace", "اضغط backspace"), params=("key",)),
    IntentRule("keyboard.escape", "press_key", "", "safe", ("press escape", "اضغط escape"), params=("key",)),
    IntentRule("keyboard.tab", "press_key", "", "safe", ("press tab", "اضغط tab"), params=("key",)),
    IntentRule("keyboard.copy", "hotkey", "", "safe", ("ctrl c", "نسخ"), params=("keys",)),
    IntentRule("keyboard.paste", "hotkey", "", "safe", ("ctrl v", "لصق"), params=("keys",)),
    IntentRule("keyboard.undo", "hotkey", "", "safe", ("ctrl z", "تراجع"), params=("keys",)),
    IntentRule("keyboard.select_all", "hotkey", "", "safe", ("ctrl a", "تحديد الكل"), params=("keys",)),
    IntentRule("keyboard.save", "hotkey", "", "safe", ("ctrl s", "حفظ"), params=("keys",)),
    IntentRule("keyboard.emoji_panel", "shell_tools", "emoji_panel", "safe", ("emoji panel", "لوحه الايموجي")),
    IntentRule("keyboard.start_menu", "shell_tools", "start_menu", "safe", ("windows key", "start menu", "قائمه ابدا")),
    IntentRule("shell.new_virtual_desktop", "shell_tools", "new_virtual_desktop", "safe", ("new virtual desktop", "سطح مكتب افتراضي جديد")),
    IntentRule("shell.next_virtual_desktop", "shell_tools", "next_virtual_desktop", "safe", ("next virtual desktop", "التنقل بين الاسطح")),
    IntentRule("shell.close_virtual_desktop", "shell_tools", "close_virtual_desktop", "safe", ("close virtual desktop", "اغلاق سطح المكتب الافتراضي")),
    IntentRule("shell.quick_settings", "shell_tools", "quick_settings", "safe", ("quick settings", "الاعدادات السريعه")),
    IntentRule("shell.notifications", "shell_tools", "notifications", "safe", ("notification center", "مركز الاشعارات")),
    IntentRule("shell.search", "shell_tools", "search", "safe", ("windows search", "بحث ويندوز")),
    IntentRule("shell.run", "shell_tools", "run", "safe", ("run dialog", "نافذه run")),
    IntentRule("shell.magnifier_open", "shell_tools", "magnifier_open", "safe", ("magnifier", "تكبير منطقه")),
    IntentRule("shell.magnifier_close", "shell_tools", "magnifier_close", "safe", ("close magnifier", "اغلاق المكبر")),
    IntentRule("shell.file_explorer", "shell_tools", "file_explorer", "safe", ("file explorer", "مستكشف الملفات")),
    IntentRule("shell.refresh", "shell_tools", "refresh", "safe", ("refresh desktop", "تحديث سطح المكتب")),
    IntentRule("shell.quick_link_menu", "shell_tools", "quick_link_menu", "safe", ("win x", "quick link menu", "قائمه الارتباط السريع")),
    IntentRule("shell.narrator_toggle", "shell_tools", "narrator_toggle", "safe", ("narrator", "الراوي")),
)


def _build_params(rule: IntentRule, raw_text: str, normalized: str) -> dict[str, Any]:
    params: dict[str, Any] = {}
    if rule.mode:
        params["mode"] = rule.mode

    value = _extract_first_int(raw_text)
    if rule.action in {"volume", "brightness"}:
        if rule.mode == "set" and value is not None:
            params["level"] = max(0, min(100, value))
        elif rule.mode in {"up", "down"}:
            params["delta"] = max(1, min(100, abs(value) if value is not None else 10))

    if rule.action == "shutdown_schedule" and rule.mode == "set":
        mins = _extract_minutes(raw_text)
        if mins is not None:
            params["minutes"] = max(1, min(1440, mins))

    if "host" in rule.params:
        host = _extract_host(raw_text) or _extract_app_query(raw_text)
        if host:
            params["host"] = host
    if "query" in rule.params:
        q = _extract_app_query(raw_text)
        if q:
            params["query"] = q
    if "process_name" in rule.params:
        q = _extract_app_query(raw_text)
        if q:
            params["process_name"] = q
    if "name" in rule.params:
        q = _extract_app_query(raw_text)
        if q:
            params["name"] = q
    if "opacity" in rule.params and value is not None:
        params["opacity"] = max(20, min(100, value))
    if "x" in rule.params and "y" in rule.params:
        nums = [int(x) for x in re.findall(r"-?\d+", raw_text)[:2]]
        if len(nums) == 2:
            params["x"], params["y"] = nums[0], nums[1]
    if rule.capability_id.startswith("mouse.click"):
        params["button"] = "left"
        params["clicks"] = 1
        if rule.capability_id.endswith("right"):
            params["button"] = "right"
        if "double" in rule.capability_id:
            params["clicks"] = 2
    if rule.capability_id == "mouse.scroll_up":
        params["key"] = "pageup"
    if rule.capability_id == "mouse.scroll_down":
        params["key"] = "pagedown"
    if rule.capability_id in {"keyboard.enter", "keyboard.space", "keyboard.backspace", "keyboard.escape", "keyboard.tab"}:
        key_lookup = {
            "keyboard.enter": "enter",
            "keyboard.space": "space",
            "keyboard.backspace": "backspace",
            "keyboard.escape": "esc",
            "keyboard.tab": "tab",
        }
        params["key"] = key_lookup[rule.capability_id]
    if rule.capability_id in {"keyboard.copy", "keyboard.paste", "keyboard.undo", "keyboard.select_all", "keyboard.save"}:
        keys_lookup = {
            "keyboard.copy": ["ctrl", "c"],
            "keyboard.paste": ["ctrl", "v"],
            "keyboard.undo": ["ctrl", "z"],
            "keyboard.select_all": ["ctrl", "a"],
            "keyboard.save": ["ctrl", "s"],
        }
        params["keys"] = keys_lookup[rule.capability_id]
    if rule.capability_id == "keyboard.type":
        quoted = re.search(r"[\"“](.+?)[\"”]|'(.+?)'", raw_text or "")
        if quoted:
            text_val = (quoted.group(1) or quoted.group(2) or "").strip()
            if text_val:
                params["text"] = text_val
    if rule.capability_id == "network.connect_named":
        named = _extract_app_query(raw_text)
        if named:
            params["host"] = named
    if rule.capability_id == "web.open_url":
        url = _extract_url(raw_text)
        if url:
            params["url"] = url
    if rule.capability_id == "apps.close_app" and not params.get("process_name"):
        params["process_name"] = "notepad"
    if rule.capability_id == "files.create_folder" and not params.get("name"):
        params["name"] = "New Folder"
    return params


def resolve_windows_intent(message: str) -> IntentResolution:
    raw = message or ""
    normalized = _normalize_text(raw)
    if not normalized:
        return IntentResolution(matched=False)

    # Contextual override for percentage-based audio/brightness set.
    if _contains_any(normalized, ("volume", "الصوت", "الاضاءه", "السطوع", "brightness")):
        if any(token in normalized for token in ("set", "اجعل", "خلي", "اعمل", "to ", "الى", "إلى")):
            value = _extract_first_int(raw)
            if value is not None:
                if _contains_any(normalized, ("volume", "الصوت")):
                    return IntentResolution(
                        matched=True,
                        capability_id="audio.set",
                        action="volume",
                        params={"mode": "set", "level": max(0, min(100, value))},
                        risk_level="safe",
                    )
                if _contains_any(normalized, ("brightness", "الاضاءه", "السطوع")):
                    return IntentResolution(
                        matched=True,
                        capability_id="display.set_brightness",
                        action="brightness",
                        params={"mode": "set", "level": max(0, min(100, value))},
                        risk_level="safe",
                    )

    for rule in RULES:
        if _contains_any(normalized, tuple(_normalize_text(a) for a in rule.aliases)):
            if rule.unsupported_reason:
                return IntentResolution(
                    matched=True,
                    capability_id=rule.capability_id,
                    risk_level=rule.risk_level,
                    unsupported=True,
                    unsupported_reason=rule.unsupported_reason,
                )
            params = _build_params(rule, raw, normalized)
            return IntentResolution(
                matched=True,
                capability_id=rule.capability_id,
                action=rule.action,
                params=params,
                risk_level=rule.risk_level,
            )

    # Unsupported but explicitly known asks from requested catalog.
    if _contains_any(normalized, ("change audio output", "speaker headset", "تغيير مخرج الصوت", "spatial sound", "الصوت المحيطي")):
        return IntentResolution(
            matched=True,
            capability_id="audio.unsupported.output_routing",
            risk_level="safe",
            unsupported=True,
            unsupported_reason="Audio output routing/spatial sound automation is not implemented yet.",
        )
    if _contains_any(normalized, ("mute microphone", "unmute microphone", "كتم الميكروفون", "الغاء كتم الميكروفون")):
        return IntentResolution(
            matched=True,
            capability_id="audio.unsupported.mic_toggle",
            risk_level="safe",
            unsupported=True,
            unsupported_reason="Microphone mute/unmute direct toggle is not implemented yet.",
        )
    return IntentResolution(matched=False)

