"""Deterministic Windows intent mapping for DesktopTool actions."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
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
    return any(tok and tok in text for tok in tokens)


def _extract_first_int(text: str) -> int | None:
    m = re.search(r"-?\d+", text)
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def _extract_ints(text: str, limit: int = 6) -> list[int]:
    values: list[int] = []
    for raw in re.findall(r"-?\d+", text or ""):
        try:
            values.append(int(raw))
        except Exception:
            continue
        if len(values) >= limit:
            break
    return values


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
        r"(?i)\b(open|launch|start|focus|switch|close|kill|stop|افتح|شغل|ركز|بدل|اغلق|سكر|اقفل|ايقاف|تشغيل|التبديل|to|الى|إلى|app|application|program|service|تطبيق|برنامج|خدمه|خدمة)\b",
        " ",
        text,
    )
    cleaned = re.sub(r"[\"'`]", " ", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:120]


def _extract_quoted_chunks(text: str) -> list[str]:
    chunks: list[str] = []
    for match in re.finditer(r"[\"“”']([^\"“”']+)[\"“”']", text or ""):
        value = (match.group(1) or "").strip()
        if value:
            chunks.append(value)
    return chunks


def _extract_paths(text: str) -> list[str]:
    quoted = _extract_quoted_chunks(text)
    path_like = [q for q in quoted if ("\\" in q or "/" in q or re.match(r"^[A-Za-z]:", q))]
    if path_like:
        return path_like
    candidates: list[str] = []
    for token in re.findall(r"[A-Za-z]:\\[^\s]+", text or ""):
        if token:
            candidates.append(token)
    return candidates


def _extract_extension(text: str) -> str:
    m = re.search(r"\.[a-zA-Z0-9]{1,8}", text or "")
    if m:
        return m.group(0).lower()
    m = re.search(r"(?:امتداد|extension)\s+([a-zA-Z0-9]{1,8})", text or "", re.IGNORECASE)
    if m:
        return f".{m.group(1).lower()}"
    return ""


def _extract_corner(text: str) -> str:
    norm = _normalize_text(text)
    if _contains_any(norm, ("top right", "فوق يمين", "اعلى يمين", "top_right")):
        return "top_right"
    if _contains_any(norm, ("bottom left", "تحت يسار", "اسفل يسار", "bottom_left")):
        return "bottom_left"
    if _contains_any(norm, ("bottom right", "تحت يمين", "اسفل يمين", "bottom_right")):
        return "bottom_right"
    return "top_left"


def _extract_named_value(text: str, patterns: tuple[str, ...]) -> str:
    quoted = _extract_quoted_chunks(text)
    if quoted:
        return quoted[0]
    for pat in patterns:
        m = re.search(pat, text or "", re.IGNORECASE)
        if not m:
            continue
        value = (m.group(1) or "").strip(" .")
        if value:
            return value
    return ""


def _extract_key_name(text: str) -> str:
    norm = _normalize_text(text)
    quoted = _extract_quoted_chunks(text)
    if quoted and len(quoted[0]) <= 16:
        return quoted[0].lower()

    key_map = {
        "enter": ("enter", "انتر"),
        "space": ("space", "مسافه", "مسافة"),
        "tab": ("tab", "تاب"),
        "esc": ("esc", "escape"),
        "up": ("up", "اعلى", "فوق"),
        "down": ("down", "اسفل", "تحت"),
        "left": ("left", "يسار"),
        "right": ("right", "يمين"),
        "f5": ("f5",),
    }
    for key, tokens in key_map.items():
        if _contains_any(norm, tokens):
            return key
    return ""


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
    IntentRule("system.power_plan_balanced", "system_power", "power_plan_balanced", "safe", ("balanced power", "متوازن", "خطة الطاقة المتوازنة")),
    IntentRule("system.power_plan_saver", "system_power", "power_plan_saver", "safe", ("power saver", "battery saver on", "توفير الطاقة", "تفعيل توفير البطارية")),
    IntentRule("system.power_plan_high", "system_power", "power_plan_high", "safe", ("high performance", "اداء عالي", "أداء عالي")),
    IntentRule("system.battery_status", "system_info", "battery", "safe", ("battery percentage", "battery status", "نسبة البطارية", "نسبه البطاريه", "كم نسبة البطارية", "شو نسبة البطارية")),
    IntentRule("system.battery_saver_off", "system_power", "power_plan_balanced", "safe", ("disable battery saver", "تعطيل توفير البطارية")),
    IntentRule("system.bios", "system_power", "reboot_bios", "destructive", ("bios", "فتح البيوس", "فتح الـ bios", "reboot bios")),
    IntentRule("system.rename_pc", "system_power", "rename_pc", "elevated", ("rename computer", "rename pc", "تغيير اسم الكمبيوتر"), params=("name",)),
    IntentRule("system.schedule_shutdown", "shutdown_schedule", "set", "destructive", ("schedule shutdown", "shutdown after", "جدوله ايقاف", "جدولة ايقاف")),
    IntentRule("system.cancel_shutdown", "shutdown_schedule", "cancel", "safe", ("cancel shutdown", "الغاء جدوله", "إلغاء جدولة")),
    IntentRule("audio.mute", "volume", "mute", "safe", ("mute", "كتم الصوت", "اكتم الصوت")),
    IntentRule("audio.unmute", "volume", "unmute", "safe", ("unmute", "الغاء الكتم", "إلغاء الكتم")),
    IntentRule("audio.up", "volume", "up", "safe", ("volume up", "raise volume", "رفع الصوت", "اعلي الصوت")),
    IntentRule("audio.down", "volume", "down", "safe", ("volume down", "خفض الصوت", "وطي الصوت")),
    IntentRule("audio.get", "volume", "get", "safe", ("current volume", "volume level", "مستوى الصوت", "نسبه الصوت", "نسبة الصوت", "كم نسبة الصوت", "شو نسبة الصوت")),
    IntentRule("audio.play_pause", "media_control", "play_pause", "safe", ("play pause", "pause media", "ايقاف مؤقت", "تشغيل او ايقاف مؤقت", "شغل وقف", "شغل/وقف")),
    IntentRule("audio.next", "media_control", "next", "safe", ("next song", "next track", "التالي", "الاغنيه التاليه", "المقطع التالي")),
    IntentRule("audio.previous", "media_control", "previous", "safe", ("previous song", "previous track", "السابق", "الاغنيه السابقه", "المقطع السابق")),
    IntentRule("audio.stop", "media_tools", "stop_all_media", "safe", ("stop media", "ايقاف الوسائط", "stop playback", "ايقاف كل الوسائط", "إيقاف كل الوسائط")),
    IntentRule("audio.mixer", "app_tools", "open_volume_mixer", "safe", ("volume mixer", "sndvol", "خلط الصوت", "الميكسار")),
    IntentRule("audio.mic_settings", "app_tools", "open_mic_settings", "safe", ("microphone settings", "mic settings", "اعدادات الميكروفون", "إعدادات الميكروفون")),
    IntentRule("audio.voice_record", "microphone_record", "", "safe", ("start voice recorder", "تسجيل صوت سريع", "ابدأ تسجيل صوت"), params=("seconds",)),
    IntentRule("audio.set_output", unsupported_reason="Changing audio output is not implemented in DesktopTool yet.", aliases=("audio output", "speaker headset", "تغيير مخرج الصوت")),
    IntentRule("audio.spatial_sound", unsupported_reason="Spatial sound toggle is not implemented in DesktopTool yet.", aliases=("spatial sound", "الصوت المحيطي")),
    IntentRule("display.brightness_up", "brightness", "up", "safe", ("brightness up", "raise brightness", "رفع الاضاءه", "رفع السطوع")),
    IntentRule("display.brightness_down", "brightness", "down", "safe", ("brightness down", "خفض الاضاءه", "خفض السطوع")),
    IntentRule("display.brightness_get", "brightness", "get", "safe", ("brightness level", "current brightness", "مستوى الاضاءه", "مستوى السطوع", "نسبة الاضاءة", "كم نسبة الاضاءة", "كم السطوع")),
    IntentRule("display.night_light_on", "ui_tools", "night_light_on", "safe", ("night light on", "تفعيل الوضع الليلي")),
    IntentRule("display.night_light_off", "ui_tools", "night_light_off", "safe", ("night light off", "تعطيل الوضع الليلي")),
    IntentRule("display.project_panel", "window_control", "project_panel", "safe", ("project", "العرض على شاشة أخرى", "العرض على شاشه اخرى")),
    IntentRule("display.extend", "window_control", "display_extend", "safe", ("extend", "وضع توسيع الشاشة", "وضع توسيع الشاشه")),
    IntentRule("display.duplicate", "window_control", "display_duplicate", "safe", ("duplicate", "وضع تكرار الشاشة", "وضع تكرار الشاشه")),
    IntentRule("display.resolution", unsupported_reason="Changing display resolution from intent map is not implemented yet.", aliases=("change resolution", "تغيير دقة الشاشة", "تغيير دقه الشاشه")),
    IntentRule("display.rotate", unsupported_reason="Display rotation from intent map is not implemented yet.", aliases=("rotate screen", "تدوير الشاشة", "تدوير الشاشه")),
    IntentRule("display.screenshot_window", "screenshot_tools", "window_active", "safe", ("window screenshot", "لقطه نافذه")),
    IntentRule("display.screenshot_full", "screenshot_tools", "full", "safe", ("full screenshot", "screenshot", "لقطه شاشه", "سكرين شوت")),
    IntentRule("display.snipping_tool", "screenshot_tools", "snipping_tool", "safe", ("snipping tool", "اداه القص")),
    IntentRule("display.clipboard_history", "clipboard_tools", "history", "safe", ("clipboard history", "سجل الحافظة", "سجل الحافظه", "الحافظة السحابية", "افتح سجل الحافظه", "افتح سجل الحافظة")),
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
    IntentRule("files.paste", "hotkey", "", "safe", ("paste file", "لصق ملف"), params=("keys",)),
    IntentRule("files.rename", "file_tools", "rename", "safe", ("rename file", "اعاده تسميه", "إعادة تسمية"), params=("path", "name")),
    IntentRule("files.zip", "file_tools", "zip", "safe", ("zip file", "ضغط ملف"), params=("path", "target")),
    IntentRule("files.unzip", "file_tools", "unzip", "safe", ("unzip file", "فك ضغط"), params=("path", "target")),
    IntentRule("files.search_ext", "file_tools", "search_ext", "safe", ("search extension", "امتداد", "ابحث عن ملف"), params=("ext",)),
    IntentRule("files.show_hidden", "file_tools", "show_hidden", "safe", ("show hidden files", "اظهار الملفات المخفية", "إظهار الملفات المخفية")),
    IntentRule("files.hide_hidden", "file_tools", "hide_hidden", "safe", ("hide hidden files", "اخفاء الملفات المخفية", "إخفاء الملفات المخفية")),
    IntentRule("files.folder_size", "file_tools", "folder_size", "safe", ("folder size", "حجم المجلد"), params=("path",)),
    IntentRule("files.open_cmd_here", "file_tools", "open_cmd_here", "safe", ("open cmd here", "فتح المسار في cmd"), params=("path",)),
    IntentRule("files.open_powershell_here", "file_tools", "open_powershell_here", "safe", ("open powershell here", "فتح المسار في powershell"), params=("path",)),
    IntentRule("network.wifi_on", "network_tools", "wifi_on", "safe", ("wifi on", "تشغيل الواي فاي", "شغل الواي فاي")),
    IntentRule("network.wifi_off", "network_tools", "wifi_off", "safe", ("wifi off", "ايقاف الواي فاي", "طفي الواي فاي")),
    IntentRule("network.wifi_passwords", "network_tools", "wifi_passwords", "elevated", ("wifi passwords", "كلمات سر الواي فاي")),
    IntentRule("network.ip_internal", "network_tools", "ip_internal", "safe", ("internal ip", "local ip", "ip الداخلي")),
    IntentRule("network.ip_external", "network_tools", "ip_external", "safe", ("external ip", "public ip", "ip الخارجي")),
    IntentRule("network.renew_ip", "network_tools", "renew_ip", "elevated", ("renew ip", "release renew", "تجديد الip")),
    IntentRule("network.flush_dns", "network_tools", "flush_dns", "safe", ("flush dns", "مسح dns", "افراغ dns", "إفراغ dns")),
    IntentRule("network.ping", "network_tools", "ping", "safe", ("ping", "اختبار اتصال", "بينق"), params=("host",)),
    IntentRule("network.bluetooth_on", "bluetooth_control", "on", "safe", ("bluetooth on", "تشغيل البلوتوث")),
    IntentRule("network.bluetooth_off", "bluetooth_control", "off", "safe", ("bluetooth off", "ايقاف البلوتوث")),
    IntentRule("network.hotspot_on", "network_tools", "hotspot_on", "elevated", ("hotspot on", "تشغيل نقطه الاتصال")),
    IntentRule("network.hotspot_off", "network_tools", "hotspot_off", "elevated", ("hotspot off", "ايقاف نقطه الاتصال")),
    IntentRule("network.settings_open", "open_settings_page", "network", "safe", ("open network settings", "فتح إعدادات الشبكة", "فتح اعدادات الشبكه", "افتح الشبكة", "اعدادات النت", "إعدادات النت"), params=("page",)),
    IntentRule("settings.privacy_open", "open_settings_page", "privacy", "safe", ("privacy settings", "open privacy settings", "اعدادات الخصوصية", "إعدادات الخصوصية", "افتح الخصوصية"), params=("page",)),
    IntentRule("network.disconnect", "network_tools", "disconnect_current_network", "safe", ("disconnect network", "قطع الاتصال بالشبكه", "افصل الشبكه", "افصل النت")),
    IntentRule("network.connect_named", "network_tools", "connect_wifi", "safe", ("connect wifi", "الاتصال بشبكه"), params=("host",)),
    IntentRule("network.open_ports", "network_tools", "open_ports", "safe", ("open ports", "المنافذ المفتوحة", "اظهار المنافذ المفتوحة")),
    IntentRule("network.port_owner", "network_tools", "port_owner", "safe", ("port owner", "من يستخدم المنفذ", "البرنامج الذي يستخدم منفذ"), params=("port",)),
    IntentRule("network.route_table", "network_tools", "route_table", "safe", ("route table", "جدول التوجيه")),
    IntentRule("network.net_scan", "network_tools", "net_scan", "safe", ("net scan", "الاجهزة المتصلة بالشبكة", "الاجهزة المتصله بالشبكه")),
    IntentRule("network.file_sharing_on", "network_tools", "file_sharing_on", "elevated", ("file sharing on", "تشغيل مشاركة الملفات")),
    IntentRule("network.file_sharing_off", "network_tools", "file_sharing_off", "elevated", ("file sharing off", "ايقاف مشاركة الملفات")),
    IntentRule("network.shared_folders", "network_tools", "shared_folders", "safe", ("shared folders", "المجلدات المشاركة", "قائمة المجلدات المشاركة")),
    IntentRule("network.server_online", "network_tools", "server_online", "safe", ("server online", "فحص توافر خادم", "هل السيرفر شغال"), params=("host",)),
    IntentRule("network.last_login_events", "network_tools", "last_login_events", "safe", ("last login attempts", "اخر محاولات الدخول", "آخر محاولات الدخول")),
    IntentRule("apps.open_browser", "app_tools", "open_default_browser", "safe", ("open browser", "افتح المتصفح")),
    IntentRule("apps.open_chrome", "app_tools", "open_chrome", "safe", ("open chrome", "افتح كروم")),
    IntentRule("apps.open_notepad", "app_tools", "open_notepad", "safe", ("open notepad", "افتح المفكره", "افتح المفكرة")),
    IntentRule("apps.open_calc", "app_tools", "open_calc", "safe", ("open calculator", "افتح الحاسبه", "افتح الحاسبة")),
    IntentRule("apps.open_paint", "app_tools", "open_paint", "safe", ("open paint", "افتح الرسام")),
    IntentRule("apps.open_task_manager", "app_tools", "open_task_manager", "safe", ("open task manager", "افتح مدير المهام", "taskmgr", "مدير المهام")),
    IntentRule("apps.close_app", "close_app", "", "elevated", ("close app", "اغلاق برنامج", "اغلق التطبيق"), params=("process_name",)),
    IntentRule("apps.close_all", "app_tools", "close_all_apps", "elevated", ("close all apps", "اغلاق كل البرامج", "اغلاق كل البرامج المفتوحة")),
    IntentRule(
        "apps.open_control_panel",
        "app_tools",
        "open_control_panel",
        "safe",
        (
            "control panel",
            "control.exe",
            "control",
            "control /name microsoft.networkandsharingcenter",
            "control /name microsoft.system",
            "control /name microsoft.sound",
            "control /name microsoft.useraccounts",
            "لوحه التحكم",
            "لوحة التحكم",
            "مركز الشبكة والمشاركة",
        ),
    ),
    IntentRule("apps.open_store", "app_tools", "open_store", "safe", ("microsoft store", "متجر مايكروسوفت", "متجر ميكروسوفت")),
    IntentRule("apps.open_registry", "app_tools", "open_registry", "destructive", ("registry editor", "regedit", "محرر السجل")),
    IntentRule("apps.open_add_remove", "app_tools", "open_add_remove_programs", "safe", ("add remove programs", "appwiz.cpl", "اضافه او ازاله البرامج")),
    IntentRule("apps.open_sound_cpl", "app_tools", "open_sound_cpl", "safe", ("mmsys.cpl", "sound settings classic", "اعدادات الصوت الكلاسيكية", "اعدادات الصوت الكلاسيكيه")),
    IntentRule("apps.open_network_connections", "app_tools", "open_network_connections", "safe", ("ncpa.cpl", "network connections", "اتصالات الشبكة", "اتصالات الشبكه")),
    IntentRule("apps.open_time_date", "app_tools", "open_time_date", "safe", ("timedate.cpl", "time and date", "الوقت والتاريخ")),
    IntentRule("apps.open_system_properties", "app_tools", "open_system_properties", "safe", ("sysdm.cpl", "system properties", "خصائص النظام")),
    IntentRule("apps.open_power_options", "app_tools", "open_power_options", "safe", ("powercfg.cpl", "power options", "خيارات الطاقة")),
    IntentRule("apps.open_firewall_cpl", "app_tools", "open_firewall_cpl", "safe", ("firewall.cpl", "windows firewall", "جدار الحماية", "جدار الحمايه")),
    IntentRule("apps.open_mouse_cpl", "app_tools", "open_mouse_cpl", "safe", ("main.cpl", "mouse properties", "خصائص الفاره", "خصائص الفأرة")),
    IntentRule("apps.open_keyboard_cpl", "app_tools", "open_keyboard_cpl", "safe", ("control keyboard", "keyboard settings classic", "اعدادات الكيبورد الكلاسيكية")),
    IntentRule("apps.open_fonts_cpl", "app_tools", "open_fonts_cpl", "safe", ("control fonts", "fonts control panel", "لوحة الخطوط", "لوحه الخطوط")),
    IntentRule("apps.open_region_cpl", "app_tools", "open_region_cpl", "safe", ("intl.cpl", "region settings", "الاقليم", "الإقليم")),
    IntentRule("apps.open_camera", "app_tools", "open_camera", "safe", ("open camera", "تشغيل الكاميرا")),
    IntentRule("apps.open_calendar", "app_tools", "open_calendar", "safe", ("open calendar", "فتح التقويم")),
    IntentRule("apps.open_mail", "app_tools", "open_mail", "safe", ("open mail", "فتح البريد")),
    IntentRule("dev.open_cmd_admin", "dev_tools", "open_cmd_admin", "elevated", ("cmd as admin", "فتح cmd كمسؤول")),
    IntentRule("dev.open_powershell_admin", "dev_tools", "open_powershell_admin", "elevated", ("powershell as admin", "فتح powershell كمسؤول")),
    IntentRule("dev.top_cpu", "process_tools", "top_cpu", "safe", ("top cpu", "اكثر العمليات استهلاكا للمعالج", "اعلى استهلاك cpu", "اعملي اعلى cpu")),
    IntentRule("dev.top_ram", "process_tools", "top_ram", "safe", ("top ram", "اكثر العمليات استهلاكا للرام", "اعلى استهلاك رام", "اعلى استهلاك ذاكره")),
    IntentRule("dev.sfc_scan", "dev_tools", "sfc_scan", "elevated", ("sfc scan", "فحص ملفات النظام")),
    IntentRule("dev.chkdsk", "dev_tools", "chkdsk", "elevated", ("chkdsk", "فحص القرص")),
    IntentRule("dev.disk_management", "dev_tools", "open_disk_management", "safe", ("disk management", "diskmgmt.msc", "اداره الاقراص")),
    IntentRule("dev.device_manager", "dev_tools", "open_device_manager", "safe", ("device manager", "devmgmt.msc", "اداره الاجهزه")),
    IntentRule("dev.perfmon", "dev_tools", "open_perfmon", "safe", ("performance monitor", "perfmon.msc", "مراقب الاداء")),
    IntentRule("dev.event_viewer", "dev_tools", "open_event_viewer", "safe", ("event viewer", "eventvwr.msc", "سجل الاحداث")),
    IntentRule("dev.task_scheduler", "dev_tools", "open_task_scheduler", "safe", ("task scheduler", "taskschd.msc", "جدولة المهام", "جدول المهام")),
    IntentRule("dev.computer_management", "dev_tools", "open_computer_management", "safe", ("computer management", "compmgmt.msc", "ادارة الكمبيوتر", "إدارة الكمبيوتر")),
    IntentRule("dev.local_users_groups", "dev_tools", "open_local_users_groups", "safe", ("lusrmgr.msc", "local users and groups", "مستخدمون ومجموعات محلية")),
    IntentRule("dev.local_security_policy", "dev_tools", "open_local_security_policy", "safe", ("secpol.msc", "local security policy", "سياسة الأمان المحلية", "سياسه الامان المحليه")),
    IntentRule("dev.print_management", "dev_tools", "open_print_management", "safe", ("printmanagement.msc", "print management", "ادارة الطباعة", "إدارة الطباعة")),
    IntentRule("dev.text_to_file", "text_tools", "text_to_file", "safe", ("text to file", "تحويل نص الى ملف", "تحويل نص لملف"), params=("text", "path")),
    IntentRule("dev.disk_health", "disk_tools", "smart_status", "safe", ("disk health", "health check", "فحص حالة القرص الصلب", "فحص حاله القرص الصلب")),
    IntentRule("dev.shortcuts", "shell_tools", "list_shortcuts", "safe", ("all shortcuts", "مفاتيح الاختصار المتاحة", "مفاتيح الاختصار المتاحه")),
    IntentRule("dev.rdp", "remote_tools", "rdp_open", "elevated", ("remote desktop", "تشغيل remote desktop")),
    IntentRule("services.stop", "service_tools", "stop", "destructive", ("stop service", "ايقاف خدمه", "ايقاف خدمة"), params=("name",)),
    IntentRule("services.restart", "service_tools", "restart", "elevated", ("restart service", "اعادة تشغيل خدمة", "اعاده تشغيل خدمة", "إعادة تشغيل خدمة"), params=("name",)),
    IntentRule("services.start", "service_tools", "start", "elevated", ("start service", "تشغيل خدمة", "تشغيل خدمه", "شغل خدمة", "شغل خدمه"), params=("name",)),
    IntentRule("services.open", "dev_tools", "open_services", "safe", ("open services", "services.msc", "فتح الخدمات")),
    IntentRule("process.restart_explorer", "process_tools", "restart_explorer", "elevated", ("restart explorer", "restart explorer.exe", "اعادة تشغيل explorer", "إعادة تشغيل واجهة الويندوز")),
    IntentRule("security.firewall_status", "security_tools", "firewall_status", "safe", ("firewall status", "حالة الجدار الناري", "هل الجدار الناري شغال", "هل جدار الحمايه شغال")),
    IntentRule("security.firewall_enable", "security_tools", "firewall_enable", "elevated", ("enable firewall", "تفعيل الجدار الناري", "تشغيل الجدار الناري")),
    IntentRule("security.firewall_disable", "security_tools", "firewall_disable", "destructive", ("disable firewall", "تعطيل الجدار الناري", "ايقاف الجدار الناري")),
    IntentRule("security.block_port", "security_tools", "block_port", "destructive", ("block port", "اغلاق منفذ", "إغلاق منفذ"), params=("port", "rule_name")),
    IntentRule("security.unblock_rule", "security_tools", "unblock_rule", "elevated", ("remove firewall rule", "حذف قاعدة جدار الحماية", "مسح قاعدة جدار الحماية"), params=("rule_name",)),
    IntentRule("security.disable_usb", "security_tools", "disable_usb", "destructive", ("disable usb", "تعطيل منافذ usb")),
    IntentRule("security.enable_usb", "security_tools", "enable_usb", "elevated", ("enable usb", "تفعيل منافذ usb")),
    IntentRule("security.disable_camera", "security_tools", "disable_camera", "destructive", ("disable camera", "تعطيل الكاميرا")),
    IntentRule("security.enable_camera", "security_tools", "enable_camera", "elevated", ("enable camera", "تفعيل الكاميرا")),
    IntentRule("security.clear_recent_files", "security_tools", "recent_files_clear", "safe", ("clear recent files", "مسح الملفات المفتوحة مؤخرا", "مسح الملفات المفتوحه مؤخرا")),
    IntentRule("security.recent_files", "security_tools", "recent_files_list", "safe", ("recent files", "الملفات المفتوحة مؤخرا", "الملفات المفتوحه مؤخرا")),
    IntentRule("security.close_remote_sessions", "security_tools", "close_remote_sessions", "elevated", ("close remote sessions", "اغلاق الجلسات عن بعد", "إغلاق الجلسات عن بعد")),
    IntentRule("security.intrusion_summary", "security_tools", "intrusion_summary", "safe", ("intrusion summary", "ملخص محاولات الاختراق", "كشف محاولات الاختراق الفاشلة")),
    IntentRule("background.count", "background_tools", "count_background", "safe", ("count background processes", "تعداد التطبيقات المشغلة في الخلفية", "عدد تطبيقات الخلفية")),
    IntentRule("background.visible_windows", "background_tools", "list_visible_windows", "safe", ("list visible windows", "قائمة التطبيقات المرئية")),
    IntentRule("background.minimized_windows", "background_tools", "list_minimized_windows", "safe", ("list minimized windows", "قائمة التطبيقات المصغرة")),
    IntentRule("background.ghost_apps", "background_tools", "ghost_apps", "safe", ("ghost apps", "التطبيقات التي لا تملك نافذة")),
    IntentRule("background.network_usage", "background_tools", "network_usage_per_app", "safe", ("network usage per app", "اي تطبيق يستخدم الانترنت", "من يستخدم النت الان")),
    IntentRule("background.camera_usage", "background_tools", "camera_usage_now", "safe", ("camera usage now", "اي تطبيق يستخدم الكاميرا", "من يستخدم الكاميرا")),
    IntentRule("background.mic_usage", "background_tools", "mic_usage_now", "safe", ("mic usage now", "اي تطبيق يستخدم الميكروفون", "من يستخدم الميكروفون")),
    IntentRule("background.wake_lock", "background_tools", "wake_lock_apps", "safe", ("wake lock apps", "التطبيقات التي تمنع السكون", "مين مانع السكون")),
    IntentRule("background.process_paths", "background_tools", "process_paths", "safe", ("process paths", "مسار التطبيقات الشغالة", "مسار التطبيق الشغال")),
    IntentRule("startup.signature_check", "startup_tools", "signature_check", "safe", ("startup signature check", "فحص امان برامج بدء التشغيل", "فحص توقيع برامج بدء التشغيل")),
    IntentRule("startup.list", "startup_tools", "startup_list", "safe", ("startup list", "قائمة برامج بدء التشغيل", "startup apps list")),
    IntentRule("startup.disable", "startup_tools", "disable", "elevated", ("disable startup", "تعطيل برنامج من بدء التشغيل"), params=("name",)),
    IntentRule("startup.enable", "startup_tools", "enable", "elevated", ("enable startup", "تفعيل برنامج في بدء التشغيل"), params=("name",)),
    IntentRule("startup.registry", "startup_tools", "registry_startups", "safe", ("registry startups", "برامج بدء التشغيل من السجل")),
    IntentRule("startup.folder", "startup_tools", "folder_startups", "safe", ("startup folder list", "برامج بدء التشغيل من مجلد startup")),
    IntentRule("startup.impact", "startup_tools", "startup_impact_time", "safe", ("startup impact time", "وقت تحميل برامج بدء التشغيل")),
    IntentRule("perf.top_cpu5", "performance_tools", "top_cpu", "safe", ("top 5 cpu", "اكثر 5 تطبيقات تستهلك المعالج", "اعلى 5 cpu")),
    IntentRule("perf.top_ram5", "performance_tools", "top_ram", "safe", ("top 5 ram", "اكثر 5 تطبيقات تستهلك الرام", "اعلى 5 ram")),
    IntentRule("perf.top_disk5", "performance_tools", "top_disk", "safe", ("top 5 disk", "اكثر 5 تطبيقات تستهلك القرص", "اعلى 5 disk")),
    IntentRule("perf.total_ram", "performance_tools", "total_ram_percent", "safe", ("total ram percent", "اجمالي استهلاك الرام", "نسبة استهلاك الرام")),
    IntentRule("perf.total_cpu", "performance_tools", "total_cpu_percent", "safe", ("total cpu percent", "اجمالي استهلاك المعالج", "نسبة استهلاك المعالج")),
    IntentRule("perf.cpu_clock", "performance_tools", "cpu_clock", "safe", ("cpu clock", "سرعة المعالج الحالية")),
    IntentRule("perf.available_ram", "performance_tools", "available_ram", "safe", ("available ram", "حجم الذاكرة المتاحة", "الرام المتاح")),
    IntentRule("perf.pagefile", "performance_tools", "pagefile_used", "safe", ("page file used", "حجم ملف التبادل", "استهلاك page file")),
    IntentRule("window.minimize", "window_control", "minimize", "safe", ("minimize window", "تصغير النافذه", "تصغير النافذة")),
    IntentRule("window.maximize", "window_control", "maximize", "safe", ("maximize window", "تكبير النافذه", "تكبير النافذة")),
    IntentRule("window.restore", "window_control", "restore", "safe", ("restore window", "استعاده النافذه", "استعادة النافذة")),
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
    IntentRule("window.hide", "window_control", "hide", "safe", ("hide window", "اخفاء نافذه", "إخفاء نافذة")),
    IntentRule("window.show", "window_control", "show", "safe", ("show hidden window", "اظهار النافذه المخفيه", "إظهار النافذة المخفية")),
    IntentRule("window.bring_to_front", "window_control", "bring_to_front", "safe", ("bring to front", "جلب نافذه للمقدمه"), params=("query",)),
    IntentRule("window.aero_shake", "window_control", "aero_shake", "safe", ("aero shake", "هز النافذه")),
    IntentRule("window.rename_title", "window_control", "rename_title", "elevated", ("rename window title", "اعادة تسمية عنوان النافذة", "اعاده تسميه عنوان النافذه"), params=("text",)),
    IntentRule("window.coords", "window_control", "coords", "safe", ("window coordinates", "احداثيات النافذة", "احداثيات النافذه")),
    IntentRule("mouse.move", "mouse_move", "", "safe", ("move mouse", "حرك الماوس"), params=("x", "y")),
    IntentRule("mouse.click_left", "click", "", "safe", ("left click", "نقره يسار", "ضغطة يسار")),
    IntentRule("mouse.click_right", "click", "", "safe", ("right click", "نقره يمين", "ضغطة يمين")),
    IntentRule("mouse.double_click", "click", "", "safe", ("double click", "نقره مزدوجه", "ضغطة مزدوجة")),
    IntentRule("mouse.down", "automation_tools", "mouse_down", "safe", ("mouse down", "الضغط المستمر"), params=("key",)),
    IntentRule("mouse.up", "automation_tools", "mouse_up", "safe", ("mouse up", "الإفلات", "الافلات"), params=("key",)),
    IntentRule("mouse.drag_drop", "automation_tools", "drag_drop", "safe", ("drag and drop", "السحب والإفلات", "السحب والافلات"), params=("x", "y", "x2", "y2")),
    IntentRule("mouse.scroll_up", "automation_tools", "scroll_up", "safe", ("scroll up", "تمرير للاعلى"), params=("repeat_count",)),
    IntentRule("mouse.scroll_down", "automation_tools", "scroll_down", "safe", ("scroll down", "تمرير للاسفل"), params=("repeat_count",)),
    IntentRule("mouse.move_corner", "automation_tools", "move_corner", "safe", ("move mouse corner", "زاوية الشاشة", "زاويه الشاشه"), params=("key",)),
    IntentRule("mouse.clip_cursor", "automation_tools", "mouse_lock_on", "safe", ("clip cursor", "حجز الماوس")),
    IntentRule("mouse.slow_move", "mouse_move", "", "safe", ("slow motion mouse", "تحريك الماوس ببطء"), params=("x", "y")),
    IntentRule("mouse.click_center", "automation_tools", "click_center", "safe", ("click center", "منتصف الشاشة", "منتصف الشاشه")),
    IntentRule("mouse.speed_up", "automation_tools", "mouse_speed_up", "safe", ("increase mouse speed", "زيادة سرعة مؤشر الماوس", "زياده سرعه مؤشر الماوس")),
    IntentRule("mouse.sonar_on", "automation_tools", "mouse_sonar_on", "safe", ("sonar effect", "دائرة حول الماوس", "دائره حول الماوس")),
    IntentRule("keyboard.type", "type_text", "", "safe", ("type text", "اكتب نص", "كتابه"), params=("text",)),
    IntentRule("keyboard.enter", "press_key", "", "safe", ("press enter", "اضغط enter"), params=("key",)),
    IntentRule("keyboard.space", "press_key", "", "safe", ("press space", "اضغط مسافه"), params=("key",)),
    IntentRule("keyboard.backspace", "press_key", "", "safe", ("press backspace", "اضغط backspace"), params=("key",)),
    IntentRule("keyboard.escape", "press_key", "", "safe", ("press escape", "اضغط escape"), params=("key",)),
    IntentRule("keyboard.tab", "press_key", "", "safe", ("press tab", "اضغط tab"), params=("key",)),
    IntentRule("keyboard.arrow_up", "press_key", "", "safe", ("arrow up", "سهم لاعلى"), params=("key",)),
    IntentRule("keyboard.arrow_down", "press_key", "", "safe", ("arrow down", "سهم لاسفل"), params=("key",)),
    IntentRule("keyboard.arrow_left", "press_key", "", "safe", ("arrow left", "سهم لليسار"), params=("key",)),
    IntentRule("keyboard.arrow_right", "press_key", "", "safe", ("arrow right", "سهم لليمين"), params=("key",)),
    IntentRule("keyboard.copy", "hotkey", "", "safe", ("ctrl c", "اختصار نسخ"), params=("keys",)),
    IntentRule("keyboard.paste", "hotkey", "", "safe", ("ctrl v", "اختصار لصق"), params=("keys",)),
    IntentRule("keyboard.undo", "hotkey", "", "safe", ("ctrl z", "اختصار تراجع"), params=("keys",)),
    IntentRule("keyboard.select_all", "hotkey", "", "safe", ("ctrl a", "اختصار تحديد الكل"), params=("keys",)),
    IntentRule("keyboard.save", "hotkey", "", "safe", ("ctrl s", "اختصار حفظ"), params=("keys",)),
    IntentRule("keyboard.caps_lock", "press_key", "", "safe", ("caps lock", "تفعيل caps lock"), params=("key",)),
    IntentRule("keyboard.num_lock", "press_key", "", "safe", ("num lock", "تفعيل num lock"), params=("key",)),
    IntentRule("keyboard.type_date", "type_text", "", "safe", ("type current date", "كتابة التاريخ الحالي")),
    IntentRule("keyboard.type_time", "type_text", "", "safe", ("type current time", "كتابة الوقت الحالي")),
    IntentRule("keyboard.repeat_key", "automation_tools", "repeat_key", "safe", ("repeat key", "تكرار ضغطة زر", "تكرار ضغطة مفتاح"), params=("key", "repeat_count")),
    IntentRule("keyboard.mouse_keys", "automation_tools", "mouse_keys_toggle", "safe", ("mouse keys", "الماوس بالكيبورد")),
    IntentRule("keyboard.emoji_panel", "shell_tools", "emoji_panel", "safe", ("emoji panel", "لوحه الايموجي")),
    IntentRule("keyboard.start_menu", "shell_tools", "start_menu", "safe", ("windows key", "start menu", "قائمه ابدا")),
    IntentRule("shell.new_virtual_desktop", "shell_tools", "new_virtual_desktop", "safe", ("new virtual desktop", "سطح مكتب افتراضي جديد")),
    IntentRule("shell.next_virtual_desktop", "shell_tools", "next_virtual_desktop", "safe", ("next virtual desktop", "التنقل بين الاسطح")),
    IntentRule("shell.prev_virtual_desktop", "shell_tools", "prev_virtual_desktop", "safe", ("previous virtual desktop", "سطح المكتب السابق")),
    IntentRule("shell.close_virtual_desktop", "shell_tools", "close_virtual_desktop", "safe", ("close virtual desktop", "اغلاق سطح المكتب الافتراضي")),
    IntentRule("shell.quick_settings", "shell_tools", "quick_settings", "safe", ("quick settings", "الاعدادات السريعه", "الاعدادات السريعة")),
    IntentRule("shell.notifications", "shell_tools", "notifications", "safe", ("notification center", "مركز الاشعارات", "مركز الاشعارات")),
    IntentRule("shell.search", "shell_tools", "search", "safe", ("windows search", "بحث ويندوز")),
    IntentRule("shell.run", "shell_tools", "run", "safe", ("run dialog", "نافذه run")),
    IntentRule("shell.magnifier_open", "shell_tools", "magnifier_open", "safe", ("magnifier", "تكبير منطقه")),
    IntentRule("shell.magnifier_zoom_out", "shell_tools", "magnifier_zoom_out", "safe", ("zoom out magnifier", "تصغير منطقة المكبر", "تصغير منطقه المكبر")),
    IntentRule("shell.magnifier_close", "shell_tools", "magnifier_close", "safe", ("close magnifier", "اغلاق المكبر")),
    IntentRule("shell.file_explorer", "shell_tools", "file_explorer", "safe", ("file explorer", "مستكشف الملفات")),
    IntentRule("shell.empty_ram", "maintenance_tools", "empty_ram", "elevated", ("empty ram", "إفراغ الرام", "افراغ الرام", "تفريغ الرام")),
    IntentRule("shell.refresh", "shell_tools", "refresh", "safe", ("refresh desktop", "تحديث سطح المكتب")),
    IntentRule("shell.quick_link_menu", "shell_tools", "quick_link_menu", "safe", ("win x", "quick link menu", "قائمه الارتباط السريع")),
    IntentRule("shell.narrator_toggle", "shell_tools", "narrator_toggle", "safe", ("narrator", "الراوي")),
    IntentRule("browser.new_tab", "browser_control", "new_tab", "safe", ("new tab", "فتح تبويب جديد")),
    IntentRule("browser.close_tab", "browser_control", "close_tab", "safe", ("close tab", "اغلاق التبويب الحالي", "إغلاق التبويب الحالي")),
    IntentRule("browser.reopen_tab", "browser_control", "reopen_tab", "safe", ("reopen closed tab", "اعادة فتح التبويب المغلق", "إعادة فتح التبويب المغلق")),
    IntentRule("browser.next_tab", "browser_control", "next_tab", "safe", ("next tab", "التبويب التالي")),
    IntentRule("browser.prev_tab", "browser_control", "prev_tab", "safe", ("previous tab", "التبويب السابق")),
    IntentRule("browser.reload", "browser_control", "reload", "safe", ("reload page", "تحديث الصفحة")),
    IntentRule("browser.incognito", "browser_control", "incognito", "safe", ("incognito", "private window", "التصفح الخفي")),
    IntentRule("browser.home", "browser_control", "home", "safe", ("browser home", "صفحة البداية")),
    IntentRule("browser.history", "browser_control", "history", "safe", ("browser history", "سجل التصفح")),
    IntentRule("browser.downloads", "browser_control", "downloads", "safe", ("browser downloads", "تنزيلات المتصفح")),
    IntentRule("browser.find", "browser_control", "find", "safe", ("find in page", "البحث داخل الصفحة")),
    IntentRule("browser.zoom_in", "browser_control", "zoom_in", "safe", ("zoom in", "تكبير الصفحة")),
    IntentRule("browser.zoom_out", "browser_control", "zoom_out", "safe", ("zoom out", "تصغير الصفحة")),
    IntentRule("browser.zoom_reset", "browser_control", "zoom_reset", "safe", ("zoom 100", "الحجم الطبيعي", "ارجاع الزوم 100")),
    IntentRule("browser.save_pdf", "browser_control", "save_pdf", "safe", ("save page pdf", "حفظ الصفحة pdf")),
    IntentRule("tasks.list", "task_tools", "list", "safe", ("task scheduler list", "قائمة المهام المجدولة")),
    IntentRule("tasks.run", "task_tools", "run", "safe", ("run scheduled task", "تشغيل مهمة مجدولة"), params=("name",)),
    IntentRule("tasks.delete", "task_tools", "delete", "destructive", ("delete scheduled task", "حذف مهمة مجدولة"), params=("name",)),
    IntentRule("tasks.create", "task_tools", "create", "elevated", ("create scheduled task", "انشاء مهمة مجدولة", "إنشاء مهمة مجدولة"), params=("name", "command", "trigger")),
    IntentRule("users.list", "user_tools", "list", "safe", ("list users", "قائمة المستخدمين")),
    IntentRule("users.create", "user_tools", "create", "destructive", ("create user", "انشاء مستخدم", "إنشاء مستخدم"), params=("username", "password")),
    IntentRule("users.delete", "user_tools", "delete", "destructive", ("delete user", "حذف مستخدم"), params=("username",)),
    IntentRule("users.set_password", "user_tools", "set_password", "destructive", ("change user password", "تغيير كلمة سر مستخدم"), params=("username", "password")),
    IntentRule("users.set_type", "user_tools", "set_type", "elevated", ("set user type", "تغيير نوع المستخدم", "admin standard"), params=("username", "group")),
    IntentRule("updates.list", "update_tools", "list_updates", "safe", ("list updates", "قائمة التحديثات", "التحديثات المثبتة")),
    IntentRule("updates.last_time", "update_tools", "last_update_time", "safe", ("last update time", "اخر تحديث للنظام", "آخر تحديث للنظام")),
    IntentRule("updates.check", "update_tools", "check_updates", "safe", ("check updates", "فحص التحديثات", "تحقق من التحديثات", "التحقق من وجود تحديثات لويندوز")),
    IntentRule("updates.install_kb", "update_tools", "install_kb", "destructive", ("install kb", "تثبيت تحديث kb"), params=("target",)),
    IntentRule("updates.cleanup_winsxs", "update_tools", "winsxs_cleanup", "elevated", ("winsxs cleanup", "تنظيف winsxs", "تنظيف ملفات التحديثات القديمة")),
    IntentRule("updates.stop_background", "update_tools", "stop_background_updates", "destructive", ("stop windows updates", "ايقاف تحديثات ويندوز", "إيقاف تحديثات ويندوز الجارية")),
    IntentRule("remote.vpn_connect", "remote_tools", "vpn_connect", "elevated", ("vpn connect", "تفعيل vpn", "تشغيل vpn"), params=("host",)),
    IntentRule("remote.vpn_disconnect", "remote_tools", "vpn_disconnect", "safe", ("vpn disconnect", "قطع اتصال vpn", "ايقاف vpn")),
    IntentRule("disk.clean_temp", "disk_tools", "temp_files_clean", "safe", ("clean temp files", "تنظيف الملفات المؤقتة", "مسح temp")),
    IntentRule("disk.clean_prefetch", "disk_tools", "prefetch_clean", "elevated", ("clean prefetch", "مسح prefetch")),
    IntentRule("disk.clean_logs", "disk_tools", "logs_clean", "elevated", ("clean windows logs", "مسح ملفات سجلات الويندوز", "تنظيف سجل النظام")),
    IntentRule("disk.usage", "disk_tools", "disk_usage", "safe", ("disk usage", "مساحة الاقراص", "استخدام القرص")),
    IntentRule("disk.defrag", "disk_tools", "defrag", "elevated", ("defrag", "الغاء تجزئة", "إلغاء تجزئة"), params=("drive",)),
    IntentRule("disk.chkdsk_scan", "disk_tools", "chkdsk_scan", "elevated", ("chkdsk scan", "فحص الباد سيكتور", "فحص القرص"), params=("drive",)),
    IntentRule("registry.query", "registry_tools", "query", "safe", ("registry query", "استعلام السجل"), params=("key",)),
    IntentRule("registry.add_key", "registry_tools", "add_key", "destructive", ("registry add key", "اضافة مفتاح للسجل", "إضافة مفتاح للسجل"), params=("key",)),
    IntentRule("registry.delete_key", "registry_tools", "delete_key", "destructive", ("registry delete key", "حذف مفتاح من السجل"), params=("key",)),
    IntentRule("registry.set_value", "registry_tools", "set_value", "destructive", ("registry set value", "تعديل قيمة في السجل", "registry dword"), params=("key", "value_name", "value_data", "value_type")),
    IntentRule("registry.backup", "registry_tools", "backup", "safe", ("registry backup", "نسخة احتياطية للسجل"), params=("key",)),
    IntentRule("registry.restore", "registry_tools", "restore", "destructive", ("registry restore", "استعادة نسخة السجل"), params=("key", "value_data")),
    IntentRule("search.text", "search_tools", "search_text", "safe", ("search text in files", "البحث عن نص داخل الملفات"), params=("folder", "pattern")),
    IntentRule("search.big_files", "search_tools", "files_larger_than", "safe", ("files larger than", "ملفات اكبر من", "ملفات أكبر من"), params=("folder", "size_mb")),
    IntentRule("search.modified_today", "search_tools", "modified_today", "safe", ("files modified today", "ملفات تم تعديلها اليوم"), params=("folder",)),
    IntentRule("search.images", "search_tools", "find_images", "safe", ("find all images", "ايجاد جميع الصور", "إيجاد جميع الصور"), params=("folder",)),
    IntentRule("search.videos", "search_tools", "find_videos", "safe", ("find all videos", "ايجاد جميع الفيديوهات", "إيجاد جميع الفيديوهات"), params=("folder",)),
    IntentRule("search.count_files", "search_tools", "count_files", "safe", ("count files", "احصاء عدد الملفات", "إحصاء عدد الملفات"), params=("folder",)),
    IntentRule("search.windows_text", "search_tools", "search_windows_content", "safe", ("search in open windows", "البحث عن كلمة في النوافذ المفتوحة"), params=("pattern",)),
    IntentRule("web.open_url", "web_tools", "open_url", "safe", ("open url", "افتح رابط", "فتح رابط"), params=("url",)),
    IntentRule("web.download_file", "web_tools", "download_file", "safe", ("download file", "تحميل ملف من رابط"), params=("url",)),
    IntentRule("web.weather", "web_tools", "weather", "safe", ("weather city", "حالة الطقس", "طقس"), params=("city",)),
    IntentRule("api.currency", "api_tools", "currency", "safe", ("currency prices", "اسعار العملات", "أسعار العملات"), params=("target",)),
    IntentRule("api.translate", "api_tools", "translate_quick", "safe", ("translate text", "ترجمة كلمة", "ترجمة نص"), params=("text",)),
    IntentRule("browserdeep.multi_open", "browser_deep_tools", "multi_open", "safe", ("open multiple links", "فتح مجموعة روابط"), params=("urls",)),
    IntentRule("browserdeep.clear_chrome_cache", "browser_deep_tools", "clear_chrome_cache", "safe", ("clear chrome cache", "مسح الكاش لمتصفح chrome")),
    IntentRule("browserdeep.clear_edge_cache", "browser_deep_tools", "clear_edge_cache", "safe", ("clear edge cache", "مسح الكاش لمتصفح edge")),
    IntentRule("office.open_word_new", "office_tools", "open_word_new", "safe", ("open word new", "فتح ملف word جديد")),
    IntentRule("office.docx_to_pdf", "office_tools", "docx_to_pdf", "safe", ("docx to pdf", "تحويل docx الى pdf", "تحويل docx إلى pdf"), params=("path", "target")),
    IntentRule("office.silent_print", "office_tools", "silent_print", "elevated", ("silent print", "طباعة ملف فورا", "طباعة ملف فوراً"), params=("path",)),
    IntentRule("drivers.list", "driver_tools", "drivers_list", "safe", ("drivers list", "قائمة التعريفات المثبتة")),
    IntentRule("drivers.backup", "driver_tools", "drivers_backup", "elevated", ("backup drivers", "نسخة احتياطية للتعريفات", "اخذ نسخة احتياطية من التعريفات")),
    IntentRule("drivers.pending_updates", "driver_tools", "updates_pending", "safe", ("pending updates", "التحديثات المعلقة")),
    IntentRule("drivers.issues", "driver_tools", "drivers_issues", "safe", ("drivers issues", "تعريفات فيها مشاكل", "التعريفات التي فيها مشاكل")),
    IntentRule("info.product_key", "info_tools", "windows_product_key", "safe", ("windows product key", "مفتاح تفعيل الويندوز")),
    IntentRule("info.model", "info_tools", "model_info", "safe", ("laptop model", "موديل اللابتوب", "الشركة المصنعة")),
    IntentRule("info.system_language", "info_tools", "system_language", "safe", ("system language", "لغة النظام الحالية")),
    IntentRule("info.timezone_get", "info_tools", "timezone_get", "safe", ("current timezone", "المنطقة الزمنية الحالية")),
    IntentRule("info.timezone_set", "info_tools", "timezone_set", "elevated", ("set timezone", "تعديل المنطقة الزمنية"), params=("timezone",)),
    IntentRule("info.install_date", "info_tools", "windows_install_date", "safe", ("windows install date", "تاريخ تثبيت الويندوز")),
    IntentRule("info.refresh_rate", "info_tools", "refresh_rate", "safe", ("refresh rate", "سرعة استجابة الشاشة")),
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
        host = _extract_host(raw_text)
        if not host:
            host = _extract_named_value(
                raw_text,
                (
                    r"(?:connect(?: to)? wifi|الاتصال بشبكه|الاتصال بشبكة|اتصل بشبكه|اتصل بشبكة)\s+(.+)$",
                    r"(?:network|شبكه|شبكة)\s+(.+)$",
                ),
            )
        if not host:
            host = _extract_app_query(raw_text)
        if host:
            params["host"] = host
    if "query" in rule.params:
        q = _extract_app_query(raw_text)
        if q:
            params["query"] = q
    if "process_name" in rule.params:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:close app|اغلاق برنامج|اغلق برنامج|سكر برنامج|kill)\s+(.+)$",
                r"(?:app|program|برنامج|تطبيق)\s+(.+)$",
            ),
        )
        if not q:
            q = _extract_app_query(raw_text)
        if q:
            params["process_name"] = q
    if "name" in rule.params:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:rename computer|rename pc|تغيير اسم الكمبيوتر)\s*(?:to|الى|إلى)?\s*[:\-]?\s*(.+)$",
                r"(?:name|named|اسم|باسم)\s*[:\-]?\s*(.+)$",
            ),
        )
        if not q:
            q = _extract_app_query(raw_text)
        if q:
            params["name"] = q
    if "username" in rule.params:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:create user|delete user|change user password|set user type|انشاء مستخدم|إنشاء مستخدم|حذف مستخدم|تغيير نوع المستخدم)\s+([A-Za-z0-9._-]+)",
                r"(?:user|username|المستخدم)\s+([A-Za-z0-9._-]+)",
            ),
        )
        if q:
            params["username"] = q
    if "password" in rule.params:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:password|كلمه السر|كلمة السر)\s*[:=]?\s*([^\s]+)",
            ),
        )
        if q:
            params["password"] = q
    if "group" in rule.params:
        if _contains_any(normalized, ("admin", "administrator", "ادمن", "مسؤول")):
            params["group"] = "admin"
        elif _contains_any(normalized, ("standard", "user", "عادي")):
            params["group"] = "users"
    if "command" in rule.params:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:command|الامر|الأمر)\s*[:=]?\s*(.+)$",
            ),
        )
        if q:
            params["command"] = q
    if "trigger" in rule.params:
        q = _extract_named_value(raw_text, (r"(?:trigger|الجدوله|الجدولة)\s*[:=]?\s*(.+)$",))
        if q:
            params["trigger"] = q
    if "opacity" in rule.params and value is not None:
        params["opacity"] = max(20, min(100, value))
    if "x" in rule.params and "y" in rule.params:
        nums = _extract_ints(raw_text, limit=4)
        if len(nums) >= 2:
            params["x"], params["y"] = nums[0], nums[1]
    if "x2" in rule.params and "y2" in rule.params:
        nums = _extract_ints(raw_text, limit=4)
        if len(nums) >= 4:
            params["x"], params["y"], params["x2"], params["y2"] = nums[0], nums[1], nums[2], nums[3]

    if "path" in rule.params or "target" in rule.params:
        found_paths = _extract_paths(raw_text)
        if "path" in rule.params and found_paths:
            params["path"] = found_paths[0]
    if "target" in rule.params and len(found_paths) >= 2:
            params["target"] = found_paths[1]
    if "target" in rule.params and "target" not in params:
        kb_match = re.search(r"\bKB\d{4,8}\b", raw_text or "", re.IGNORECASE)
        if kb_match:
            params["target"] = kb_match.group(0).upper()
        else:
            q = _extract_named_value(raw_text, (r"(?:target|kb)\s*[:=]?\s*([^\s]+)",))
            if q:
                params["target"] = q
    if "ext" in rule.params:
        ext = _extract_extension(raw_text)
        if ext:
            params["ext"] = ext
    if "url" in rule.params and "url" not in params:
        u = _extract_url(raw_text)
        if u:
            params["url"] = u
    if "city" in rule.params and "city" not in params:
        q = _extract_named_value(raw_text, (r"(?:city|مدينة|مدينه)\s*[:=]?\s*(.+)$",))
        if q:
            params["city"] = q
    if "folder" in rule.params and "folder" not in params:
        found_paths = _extract_paths(raw_text)
        if found_paths:
            params["folder"] = found_paths[0]
    if "pattern" in rule.params and "pattern" not in params:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:search(?: text)?|ابحث(?: عن)?|البحث عن)\s+(.+)$",
                r"(?:pattern|كلمة|نص)\s*[:=]?\s*(.+)$",
            ),
        )
        if q:
            params["pattern"] = q
    if "size_mb" in rule.params and "size_mb" not in params and value is not None:
        params["size_mb"] = max(1, min(1024 * 100, abs(value)))
    if "timezone" in rule.params and "timezone" not in params:
        q = _extract_named_value(raw_text, (r"(?:timezone|المنطقة الزمنية)\s*[:=]?\s*(.+)$",))
        if q:
            params["timezone"] = q
    if "urls" in rule.params and "urls" not in params:
        found = re.findall(r"https?://\S+", raw_text or "", re.IGNORECASE)
        cleaned = [u.strip(" ,)") for u in found if u]
        if cleaned:
            params["urls"] = cleaned[:20]
    if "text" in rule.params:
        quoted = _extract_quoted_chunks(raw_text)
        if quoted:
            params["text"] = quoted[0]
        else:
            q = _extract_named_value(
                raw_text,
                (
                    r"(?:type string|type text|اكتب نص)\s+(.+)$",
                    r"(?:rename window title|اعاده تسميه عنوان النافذه|إعادة تسمية عنوان النافذة)\s*(?:to|الى|إلى)?\s*(.+)$",
                    r"(?:text to file|تحويل نص)\s*(?:الى|to)?\s*(.+)$",
                ),
            )
            if q:
                params["text"] = q
    if "key" in rule.params and (
        rule.capability_id.startswith("keyboard.") or rule.capability_id.startswith("mouse.")
    ):
        k = _extract_key_name(raw_text)
        if k:
            params["key"] = k
    if "repeat_count" in rule.params:
        cnt = _extract_first_int(raw_text)
        if cnt is not None:
            params["repeat_count"] = max(1, min(200, abs(cnt)))
    if "page" in rule.params:
        params["page"] = rule.mode or "network"
        params.pop("mode", None)
    if "seconds" in rule.params and value is not None:
        params["seconds"] = max(1, min(120, abs(value)))
    if rule.capability_id.startswith("mouse.click"):
        params["button"] = "left"
        params["clicks"] = 1
        if rule.capability_id.endswith("right"):
            params["button"] = "right"
        if "double" in rule.capability_id:
            params["clicks"] = 2
    if rule.capability_id == "mouse.down" and "key" not in params:
        params["key"] = "left"
    if rule.capability_id == "mouse.up" and "key" not in params:
        params["key"] = "left"
    if rule.capability_id in {"mouse.scroll_up", "mouse.scroll_down"} and "repeat_count" not in params:
        params["repeat_count"] = 4
    if rule.capability_id == "mouse.move_corner":
        params["key"] = _extract_corner(raw_text)
    if rule.capability_id == "mouse.slow_move":
        params["duration"] = 1.5
        nums = _extract_ints(raw_text, limit=2)
        if len(nums) >= 2:
            params["x"], params["y"] = nums[0], nums[1]
    if rule.capability_id in {"keyboard.enter", "keyboard.space", "keyboard.backspace", "keyboard.escape", "keyboard.tab"}:
        key_lookup = {
            "keyboard.enter": "enter",
            "keyboard.space": "space",
            "keyboard.backspace": "backspace",
            "keyboard.escape": "esc",
            "keyboard.tab": "tab",
        }
        params["key"] = key_lookup[rule.capability_id]
    if rule.capability_id in {"keyboard.arrow_up", "keyboard.arrow_down", "keyboard.arrow_left", "keyboard.arrow_right"}:
        arrow_lookup = {
            "keyboard.arrow_up": "up",
            "keyboard.arrow_down": "down",
            "keyboard.arrow_left": "left",
            "keyboard.arrow_right": "right",
        }
        params["key"] = arrow_lookup[rule.capability_id]
    if rule.capability_id in {"keyboard.caps_lock", "keyboard.num_lock"}:
        params["key"] = "capslock" if rule.capability_id == "keyboard.caps_lock" else "numlock"
    if rule.capability_id in {"keyboard.copy", "keyboard.paste", "keyboard.undo", "keyboard.select_all", "keyboard.save"}:
        keys_lookup = {
            "keyboard.copy": ["ctrl", "c"],
            "keyboard.paste": ["ctrl", "v"],
            "keyboard.undo": ["ctrl", "z"],
            "keyboard.select_all": ["ctrl", "a"],
            "keyboard.save": ["ctrl", "s"],
        }
        params["keys"] = keys_lookup[rule.capability_id]
    if rule.capability_id == "files.paste":
        params["keys"] = ["ctrl", "v"]
    if rule.capability_id == "keyboard.type":
        if "text" not in params:
            quoted = re.search(r"[\"“](.+?)[\"”]|'(.+?)'", raw_text or "")
            if quoted:
                text_val = (quoted.group(1) or quoted.group(2) or "").strip()
                if text_val:
                    params["text"] = text_val
    if rule.capability_id == "keyboard.type_date":
        params["text"] = datetime.now().strftime("%Y-%m-%d")
    if rule.capability_id == "keyboard.type_time":
        params["text"] = datetime.now().strftime("%H:%M:%S")
    if rule.capability_id == "keyboard.repeat_key":
        if "key" not in params:
            params["key"] = _extract_key_name(raw_text) or "enter"
        if "repeat_count" not in params:
            params["repeat_count"] = 3
    if rule.capability_id == "network.connect_named":
        named = _extract_app_query(raw_text)
        if named and not params.get("host"):
            params["host"] = named
    if rule.capability_id == "remote.vpn_connect" and params.get("host"):
        host = str(params.get("host") or "")
        host = re.sub(r"(?i)^\s*(?:vpn|تشغيل\s+vpn|تفعيل\s+vpn)\s+", "", host).strip()
        if host:
            params["host"] = host
    if rule.capability_id == "web.open_url":
        url = _extract_url(raw_text)
        if url:
            params["url"] = url
    if rule.capability_id == "dev.text_to_file":
        if "text" not in params:
            params["text"] = raw_text.strip()
        for chunk in _extract_quoted_chunks(raw_text):
            if chunk.lower().endswith(".txt"):
                params["path"] = chunk
                break
    if rule.capability_id.startswith("registry."):
        if "key" in rule.params and not params.get("key"):
            q = _extract_named_value(
                raw_text,
                (
                    r"((?:HKLM|HKCU|HKCR|HKU|HKCC):?\\[^\s]+)",
                    r"(HKEY_[A-Z_]+\\[^\s]+)",
                ),
            )
            if q:
                params["key"] = q
        if "value_name" in rule.params and not params.get("value_name"):
            m = re.search(r"(?:name|value[_ ]name|اسم القيمه|اسم القيمة)\s*[:=]?\s*([^\s]+)", raw_text or "", re.IGNORECASE)
            q = (m.group(1).strip() if m else "")
            if q:
                params["value_name"] = q
        if "value_data" in rule.params and not params.get("value_data"):
            m = re.search(r"(?:data|value[_ ]data|القيمه|القيمة)\s*[:=]?\s*([^\s]+)", raw_text or "", re.IGNORECASE)
            q = (m.group(1).strip() if m else "")
            if q:
                params["value_data"] = q
        if "value_type" in rule.params and not params.get("value_type"):
            if _contains_any(normalized, ("dword", "reg_dword")):
                params["value_type"] = "REG_DWORD"
            elif _contains_any(normalized, ("qword", "reg_qword")):
                params["value_type"] = "REG_QWORD"
            else:
                params["value_type"] = "REG_SZ"
    if "drive" in rule.params and not params.get("drive"):
        m = re.search(r"\b([A-Za-z]:)\b", raw_text or "")
        if m:
            params["drive"] = m.group(1).upper()
    if "rule_name" in rule.params and not params.get("rule_name"):
        q = _extract_named_value(raw_text, (r"(?:rule|rule_name|اسم القاعده|اسم القاعدة)\s*[:=]?\s*(.+)$",))
        if q:
            params["rule_name"] = q
    if "port" in rule.params and "port" not in params and value is not None:
        params["port"] = max(1, min(65535, abs(value)))
    if rule.capability_id == "services.stop" and not params.get("name"):
        q = _extract_named_value(raw_text, (r"(?:stop service|ايقاف خدمه|إيقاف خدمة)\s+(.+)$",))
        if q:
            params["name"] = q
    if rule.capability_id in {"services.start", "services.restart"}:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:restart service|اعاده تشغيل خدمة|إعادة تشغيل خدمة)\s+(.+)$",
                r"(?:start service|تشغيل خدمه|تشغيل خدمة|شغل خدمة|شغل خدمه)\s+(.+)$",
            ),
        )
        if q:
            params["name"] = q
    if rule.capability_id in {"startup.disable", "startup.enable"}:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:disable startup|تعطيل برنامج من بدء التشغيل)\s+(.+)$",
                r"(?:enable startup|تفعيل برنامج في بدء التشغيل)\s+(.+)$",
            ),
        )
        if q:
            params["name"] = q
    if rule.capability_id in {"tasks.run", "tasks.delete", "tasks.create"}:
        q = _extract_named_value(
            raw_text,
            (
                r"(?:run scheduled task|تشغيل مهمة مجدولة)\s+(.+)$",
                r"(?:delete scheduled task|حذف مهمة مجدولة)\s+(.+)$",
                r"(?:create scheduled task|انشاء مهمة مجدولة|إنشاء مهمة مجدولة)\s+(.+)$",
            ),
        )
        if q:
            params["name"] = q
    if rule.capability_id == "window.rename_title" and not params.get("text"):
        q = _extract_named_value(
            raw_text,
            (
                r"(?:rename window title|اعاده تسميه عنوان النافذه|إعادة تسمية عنوان النافذة)\s*(?:to|الى|إلى)?\s*(.+)$",
                r"(?:window title|عنوان النافذه|عنوان النافذة)\s*(?:to|الى|إلى)?\s*(.+)$",
            ),
        )
        if q:
            params["text"] = q
    if rule.capability_id == "apps.close_app" and not params.get("process_name"):
        params["process_name"] = "notepad"
    if rule.capability_id == "files.create_folder" and not params.get("name"):
        params["name"] = "New Folder"
    if rule.capability_id == "files.delete_permanent":
        params["permanent"] = True
    return params


def resolve_windows_intent(message: str) -> IntentResolution:
    raw = message or ""
    normalized = _normalize_text(raw)
    if not normalized:
        return IntentResolution(matched=False)

    # Contextual override for percentage-based audio/brightness set.
    if _contains_any(normalized, ("تمنع", "منع", "blocking", "wake lock")) and _contains_any(
        normalized, ("السكون", "sleep")
    ):
        return IntentResolution(
            matched=True,
            capability_id="background.wake_lock",
            action="background_tools",
            params={"mode": "wake_lock_apps"},
            risk_level="safe",
        )

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

