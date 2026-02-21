"""Unified Agent Loop.
Created: 2026-02-02
Changes:
  - Added BrowserTool registration
  - 2026-02-05: Refactored to use AgentRouter for all backends.
                Now properly emits system_event for tool_use/tool_result.

This is the core "brain" of Mudabbir. It integrates:
1. MessageBus (Input/Output)
2. MemoryManager (Short-term & Long-term memory)
3. AgentRouter (Backend selection: claude_agent_sdk, Mudabbir_native, open_interpreter)
4. AgentContextBuilder (Identity & System Prompt)

It replaces the old highly-coupled bot loops.
"""

import asyncio
import json
import logging
import re
from typing import Any

from Mudabbir.agents.router import AgentRouter
from Mudabbir.bootstrap import AgentContextBuilder
from Mudabbir.bus import InboundMessage, OutboundMessage, SystemEvent, get_message_bus
from Mudabbir.bus.commands import get_command_handler
from Mudabbir.bus.events import Channel
from Mudabbir.config import Settings, get_settings
from Mudabbir.memory import get_memory_manager
from Mudabbir.security.injection_scanner import ThreatLevel, get_injection_scanner

logger = logging.getLogger(__name__)

FIRST_RESPONSE_TIMEOUT_SECONDS = 90
STREAM_CHUNK_TIMEOUT_SECONDS = 120


class StreamTimeoutError(TimeoutError):
    """Raised when stream iteration times out with phase metadata."""

    def __init__(self, phase: str, timeout_seconds: float):
        self.phase = phase
        self.timeout_seconds = timeout_seconds
        super().__init__(f"{phase} chunk timed out after {timeout_seconds}s")


async def _iter_with_timeout(
    aiter,
    first_timeout=FIRST_RESPONSE_TIMEOUT_SECONDS,
    timeout=STREAM_CHUNK_TIMEOUT_SECONDS,
):
    """Yield items from an async iterator with per-item timeouts.

    Uses a shorter timeout for the first item (to detect dead/hung backends
    quickly) and a longer timeout for subsequent items (to allow for tool
    execution, file operations, etc.).
    """
    ait = aiter.__aiter__()
    first = True
    while True:
        try:
            t = first_timeout if first else timeout
            yield await asyncio.wait_for(ait.__anext__(), timeout=t)
            first = False
        except TimeoutError:
            phase = "first" if first else "stream"
            raise StreamTimeoutError(phase=phase, timeout_seconds=t) from None
        except StopAsyncIteration:
            break


class AgentLoop:
    """
    Main agent execution loop.

    Orchestrates the flow of data between Bus, Memory, and AgentRouter.
    Uses AgentRouter to delegate to the selected backend (claude_agent_sdk,
    Mudabbir_native, or open_interpreter).
    """

    def __init__(self):
        self.settings = get_settings()
        self.bus = get_message_bus()
        self.memory = get_memory_manager()
        self.context_builder = AgentContextBuilder(memory_manager=self.memory)

        # Agent Router handles backend selection
        self._router: AgentRouter | None = None

        # Concurrency controls
        self._session_locks: dict[str, asyncio.Lock] = {}
        self._session_tasks: dict[str, asyncio.Task] = {}
        self._global_semaphore = asyncio.Semaphore(self.settings.max_concurrent_conversations)
        self._background_tasks: set[asyncio.Task] = set()

        self._running = False
        self._pending_windows_dangerous: dict[str, dict[str, Any]] = {}
        self._windows_session_state: dict[str, dict[str, Any]] = {}
        get_command_handler().set_on_settings_changed(self._on_settings_changed)

    async def _llm_one_shot_text(
        self,
        *,
        system_prompt: str,
        user_prompt: str,
        max_tokens: int = 300,
        temperature: float = 0.2,
    ) -> str | None:
        """Run a single non-streaming completion using current provider settings."""
        try:
            from Mudabbir.llm.client import resolve_llm_client

            llm = resolve_llm_client(self.settings)

            if llm.is_ollama:
                import httpx

                payload = {
                    "model": llm.model,
                    "messages": [
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    "stream": False,
                    "options": {
                        "temperature": max(0.0, min(1.0, float(temperature))),
                    },
                }
                async with httpx.AsyncClient(timeout=30.0) as client:
                    resp = await client.post(f"{llm.ollama_host}/api/chat", json=payload)
                    resp.raise_for_status()
                    data = resp.json()
                    content = (
                        (data.get("message") or {}).get("content", "")
                        if isinstance(data, dict)
                        else ""
                    )
                    return str(content or "").strip() or None

            if llm.provider in {"openai", "openai_compatible", "gemini"}:
                if llm.provider == "openai":
                    from openai import AsyncOpenAI

                    client = AsyncOpenAI(
                        api_key=llm.api_key,
                        timeout=30.0,
                        max_retries=1,
                    )
                else:
                    client = llm.create_openai_client(timeout=30.0, max_retries=1)

                request_temperature = max(0.0, min(1.0, float(temperature)))
                if llm.is_gemini and str(llm.model).lower().startswith("gemini-3"):
                    request_temperature = 1.0

                response = await client.chat.completions.create(
                    model=llm.model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    temperature=request_temperature,
                    max_tokens=int(max(80, min(1200, max_tokens))),
                )
                message = response.choices[0].message if response and response.choices else None
                return str(getattr(message, "content", "") or "").strip() or None

            client = llm.create_anthropic_client(timeout=30.0, max_retries=1)
            response = await client.messages.create(
                model=llm.model,
                max_tokens=int(max(80, min(1200, max_tokens))),
                system=system_prompt,
                messages=[{"role": "user", "content": user_prompt}],
            )
            parts: list[str] = []
            for block in getattr(response, "content", []) or []:
                if getattr(block, "type", "") == "text":
                    parts.append(str(getattr(block, "text", "") or ""))
            text = "".join(parts).strip()
            return text or None
        except Exception as exc:
            logger.debug("Response composer completion failed: %s", exc)
            return None

    async def _compose_response(
        self,
        *,
        user_query: str,
        events: list[dict],
        fallback_text: str,
    ) -> str:
        """Compose a flexible factual response from execution events."""
        style = str(getattr(self.settings, "ai_response_style", "flex_factual") or "flex_factual")
        max_tokens = int(getattr(self.settings, "ai_response_max_tokens", 320) or 320)

        system_prompt = (
            "You are Mudabbir response composer.\n"
            "Write a natural, elegant user-facing reply in the same language as the user.\n"
            "Be factual. Keep concrete numbers/paths/status exactly when available.\n"
            "Never output raw tool payload JSON, code, or markdown fences.\n"
            "Do not invent actions that were not executed.\n"
            f"Style mode: {style}."
        )
        user_prompt = (
            f"User request:\n{user_query}\n\n"
            f"Execution events (JSON):\n{json.dumps(events[-24:], ensure_ascii=False)}\n\n"
            f"Fallback plain text:\n{fallback_text[:2200]}\n\n"
            "Now produce the final answer."
        )
        composed = await self._llm_one_shot_text(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            max_tokens=max_tokens,
            temperature=0.25,
        )
        if composed:
            return composed
        return fallback_text

    def _on_settings_changed(self) -> None:
        """Reload settings-sensitive runtime pieces after slash command updates."""
        self.settings = Settings.load()
        self.reset_router()

    def _get_router(self) -> AgentRouter:
        """Get or create the agent router (lazy initialization)."""
        if self._router is None:
            # Reload settings to pick up any changes
            settings = Settings.load()
            self._router = AgentRouter(settings)
        return self._router

    @staticmethod
    def _sanitize_stream_chunk(content: str) -> str:
        """Hide malformed execute payload noise from end-user chat."""
        text = str(content or "")
        stripped = text.strip()
        if not stripped:
            return text

        compact = stripped.strip("`").strip()
        lowered = compact.lower()
        if "mcp_sequential-thinking" in lowered:
            return ""
        markers = (
            "execute",
            "arguments",
            "language",
            "code",
            "start-process",
            "get-process",
            "stop-process",
            "set-volume",
            "set-volumelevel",
            "set-culture",
            "set-win",
            "set-bluetoothstate",
            "write-output",
            "ms-settings:",
            "telegram:",
            "shell:appsfolder",
            "http://",
            "https://",
            "explorer.exe",
            "pyautogui",
            "powershell",
            "powers ",
            "python",
            "mcp_sequential-thinking",
            "tool▁sep",
            "tool_call_end",
        )
        broken = ("namepowers", "argumentspowers", "namepython", "argumentspython")

        if re.search(r'^\s*(?:powers|powershell)\b', lowered) and any(
            t in lowered for t in ("-volume", "start-process", "get-process", "set-", "pyautogui", "}}")
        ):
            return ""
        if (lowered.startswith("import pyautogui") or lowered.startswith("\\nimport pyautogui")) and (
            "\\n" in compact or "pyautogui." in lowered
        ):
            return ""

        looks_jsonish = (
            compact.startswith("{")
            or compact.startswith("[")
            or bool(re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*:', lowered))
        )
        if looks_jsonish and any(m in lowered for m in markers):
            return ""
        if any(b in lowered for b in broken) and any(m in lowered for m in markers):
            return ""
        if re.search(r'^\s*"?\s*:\s*"(powershell|python|pwsh)"', lowered):
            return ""
        if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*,\s*"\s*:\s*', lowered):
            return ""
        if '"language"' in lowered and '"code"' in lowered and any(m in lowered for m in markers):
            return ""
        if compact in {"{", "}", "[", "]", "```", "`"}:
            return ""
        return text

    @staticmethod
    def _contains_arabic(text: str) -> bool:
        return bool(re.search(r"[\u0600-\u06FF]", str(text or "")))

    @staticmethod
    def _normalize_intent_text(text: str) -> str:
        normalized = str(text or "").strip().lower()
        normalized = normalized.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
        normalized = normalized.replace("ة", "ه").replace("ى", "ي")
        normalized = re.sub(r"\s+", " ", normalized)
        return normalized

    @staticmethod
    def _is_confirmation_message(text: str) -> bool:
        normalized = AgentLoop._normalize_intent_text(text)
        return normalized in {"yes", "y", "ok", "confirm", "نعم", "اي", "أجل", "اجل"}

    async def _try_global_windows_fastpath(
        self, *, text: str, session_key: str
    ) -> tuple[bool, str | None]:
        """Deterministic Windows desktop execution path before any backend call."""
        try:
            from Mudabbir.tools.builtin.desktop import DesktopTool
            from Mudabbir.tools.capabilities.windows_intent_map import resolve_windows_intent
        except Exception:
            return False, None

        arabic = self._contains_arabic(text)
        normalized = self._normalize_intent_text(text)
        cancel_tokens = ("cancel", "stop", "no", "لا", "الغاء", "إلغاء", "وقف")

        pending = self._pending_windows_dangerous.get(session_key)
        if pending is not None:
            if self._is_confirmation_message(text):
                resolution = pending
                self._pending_windows_dangerous.pop(session_key, None)
            elif any(tok in normalized for tok in cancel_tokens):
                self._pending_windows_dangerous.pop(session_key, None)
                return True, ("تم إلغاء العملية الخطرة." if arabic else "Canceled the pending dangerous operation.")
            else:
                return True, (
                    "لدي عملية خطرة بانتظار التأكيد. اكتب 'نعم' للتنفيذ أو 'إلغاء' للإلغاء."
                    if arabic
                    else "A dangerous operation is pending. Reply 'yes' to execute or 'cancel' to abort."
                )
        else:
            resolved = resolve_windows_intent(text)
            if not resolved.matched:
                return False, None
            resolution = {
                "capability_id": resolved.capability_id,
                "action": resolved.action,
                "params": dict(resolved.params or {}),
                "risk_level": str(resolved.risk_level or "safe"),
                "unsupported": bool(resolved.unsupported),
                "unsupported_reason": resolved.unsupported_reason,
            }

        if resolution.get("unsupported"):
            return True, str(
                resolution.get("unsupported_reason")
                or ("هذه المهارة غير مدعومة حالياً." if arabic else "This capability is not implemented yet.")
            )

        risk_level = str(resolution.get("risk_level", "safe"))
        if risk_level == "destructive" and not self._is_confirmation_message(text):
            self._pending_windows_dangerous[session_key] = resolution
            return True, (
                "هذا أمر خطِر. للتأكيد اكتب: نعم. للإلغاء اكتب: إلغاء."
                if arabic
                else "This is a destructive command. Reply 'yes' to confirm or 'cancel' to abort."
            )

        action = str(resolution.get("action", "")).strip()
        params = resolution.get("params", {}) if isinstance(resolution.get("params"), dict) else {}
        if not action:
            return False, None
        session_state = self._windows_session_state.setdefault(session_key, {})
        mode = str(params.get("mode", "")).lower()
        if (
            action == "process_tools"
            and mode.startswith("app_")
            and not str(params.get("name", "")).strip()
            and str(session_state.get("last_app", "")).strip()
        ):
            params["name"] = str(session_state.get("last_app", "")).strip()
        if (
            action == "service_tools"
            and mode in {"start", "stop", "restart", "describe", "dependencies", "startup"}
            and not str(params.get("name", "")).strip()
            and str(session_state.get("last_service", "")).strip()
        ):
            params["name"] = str(session_state.get("last_service", "")).strip()

        raw = await DesktopTool().execute(action=action, **params)
        raw_text = str(raw or "")
        if raw_text.lower().startswith("error:"):
            return True, raw_text

        parsed: Any = raw
        if isinstance(raw, str):
            try:
                parsed = json.loads(raw)
            except Exception:
                parsed = raw
        if isinstance(parsed, dict) and parsed.get("ok") is False:
            err = str(parsed.get("error") or parsed.get("message") or "").strip()
            if arabic:
                return True, (f"فشل التنفيذ: {err}" if err else "فشل التنفيذ.")
            return True, (f"Execution failed: {err}" if err else "Execution failed.")
        if isinstance(parsed, dict):
            remembered_app = str(parsed.get("query") or params.get("name") or "").strip()
            if remembered_app:
                session_state["last_app"] = remembered_app
            if action == "service_tools":
                remembered_service = str(parsed.get("name") or params.get("name") or "").strip()
                if remembered_service:
                    session_state["last_service"] = remembered_service
            top_app = str(parsed.get("top_app") or "").strip()
            if top_app:
                session_state["last_app"] = top_app

        if action == "volume" and str(params.get("mode", "")).lower() == "get" and isinstance(parsed, dict):
            level = parsed.get("level_percent")
            muted = bool(parsed.get("muted", False))
            if level is not None:
                if arabic:
                    return True, f"مستوى الصوت الحالي: {int(level)}% {'(مكتوم)' if muted else ''}".strip()
                return True, f"Current volume is {int(level)}%{' (muted)' if muted else ''}."
        if action == "brightness" and str(params.get("mode", "")).lower() == "get" and isinstance(parsed, dict):
            level = parsed.get("brightness_percent")
            if level is not None:
                if arabic:
                    return True, f"مستوى السطوع الحالي: {int(level)}%."
                return True, f"Current brightness is {int(level)}%."

        if action == "system_info" and str(params.get("mode", "")).lower() == "battery" and isinstance(parsed, dict):
            available = bool(parsed.get("available", False))
            percent = parsed.get("percent")
            plugged = parsed.get("plugged")
            if available and percent is not None:
                if arabic:
                    state = "موصول بالشاحن" if plugged else "على البطارية"
                    return True, f"نسبة البطارية الحالية: {int(float(percent))}% ({state})."
                state = "plugged in" if plugged else "on battery"
                return True, f"Current battery is {int(float(percent))}% ({state})."
            if arabic:
                return True, "لا يمكن قراءة معلومات البطارية على هذا الجهاز حالياً."
            return True, "Battery information is not available on this machine right now."

        if action == "clipboard_tools" and str(params.get("mode", "")).lower() in {"history", "clipboard_history"}:
            return True, ("تم فتح سجل الحافظة (Win+V)." if arabic else "Opened Clipboard History (Win+V).")

        if action == "network_tools" and str(params.get("mode", "")).lower() in {"open_network_settings", "settings"}:
            return True, ("تم فتح إعدادات الشبكة." if arabic else "Opened network settings.")
        if action == "open_settings_page":
            page = str(params.get("page", "")).strip().lower()
            page_msgs_ar = {
                "network": "تم فتح إعدادات الشبكة.",
                "privacy": "تم فتح إعدادات الخصوصية.",
                "sound": "تم فتح إعدادات الصوت.",
                "windowsupdate": "تم فتح إعدادات تحديثات ويندوز.",
                "update": "تم فتح إعدادات تحديثات ويندوز.",
                "appsfeatures": "تم فتح إعدادات التطبيقات.",
            }
            page_msgs_en = {
                "network": "Opened network settings.",
                "privacy": "Opened privacy settings.",
                "sound": "Opened sound settings.",
                "windowsupdate": "Opened Windows Update settings.",
                "update": "Opened Windows Update settings.",
                "appsfeatures": "Opened apps settings.",
            }
            msg = page_msgs_ar.get(page) if arabic else page_msgs_en.get(page)
            if msg:
                return True, msg
            return True, ("تم فتح صفحة الإعدادات." if arabic else "Opened Settings page.")
        if action == "network_tools":
            mode = str(params.get("mode", "")).lower()
            if mode == "connect_wifi" and isinstance(parsed, dict):
                requested = str(parsed.get("requested_ssid") or params.get("host") or "").strip()
                connected = parsed.get("connected")
                actual = str(parsed.get("connected_ssid") or "").strip()
                if arabic:
                    if connected is True:
                        return True, f"📶 تم الاتصال بالشبكة: {actual or requested}."
                    if connected is False:
                        return True, f"⚠️ ما تم الاتصال بـ {requested}."
                    return True, f"📶 تم إرسال طلب الاتصال بـ {requested}."
                if connected is True:
                    return True, f"📶 Connected to Wi-Fi: {actual or requested}."
                if connected is False:
                    return True, f"⚠️ Could not connect to Wi-Fi: {requested}."
                return True, f"📶 Sent Wi-Fi connect request: {requested}."
            network_msgs_ar = {
                "wifi_on": "تم تشغيل الواي فاي.",
                "wifi_off": "تم إيقاف الواي فاي.",
                "flush_dns": "تم مسح ذاكرة DNS.",
                "display_dns": "تم عرض ذاكرة DNS الحالية.",
                "renew_ip": "تم تجديد عنوان IP.",
                "disconnect_current_network": "تم قطع الاتصال بالشبكة الحالية.",
                "connect_wifi": "تم إرسال طلب الاتصال بالشبكة.",
                "ip_internal": "تم جلب عنوان IP الداخلي.",
                "ip_external": "تم جلب عنوان IP الخارجي.",
                "ipconfig_all": "تم جلب معلومات الشبكة التفصيلية.",
                "ping": "تم تنفيذ اختبار الاتصال (Ping).",
                "open_ports": "تم جلب المنافذ المفتوحة.",
                "port_owner": "تم جلب البرنامج الذي يستخدم المنفذ.",
                "route_table": "تم جلب جدول التوجيه.",
                "tracert": "تم تنفيذ تتبع المسار.",
                "pathping": "تم تنفيذ فحص المسار وفقدان الحزم.",
                "nslookup": "تم تنفيذ استعلام DNS.",
                "netstat_active": "تم جلب الاتصالات النشطة.",
                "getmac": "تم جلب عناوين MAC للأجهزة.",
                "arp_table": "تم جلب جدول ARP.",
                "nbtstat_cache": "تم جلب كاش NetBIOS.",
                "nbtstat_host": "تم تنفيذ استعلام NetBIOS للجهاز.",
                "net_view": "تم جلب قائمة أجهزة الشبكة.",
                "netstat_binary": "تم جلب الاتصالات مع البرامج المرتبطة بها.",
                "wifi_profiles": "تم جلب قائمة شبكات الواي فاي المحفوظة.",
                "net_scan": "تم جلب الأجهزة المتصلة على الشبكة المحلية.",
                "file_sharing_on": "تم تشغيل مشاركة الملفات.",
                "file_sharing_off": "تم إيقاف مشاركة الملفات.",
                "shared_folders": "تم جلب قائمة المجلدات المشاركة.",
                "server_online": "تم فحص توافر الخادم.",
                "last_login_events": "تم جلب سجل آخر محاولات الدخول.",
            }
            network_msgs_en = {
                "wifi_on": "Wi-Fi turned on.",
                "wifi_off": "Wi-Fi turned off.",
                "flush_dns": "DNS cache flushed.",
                "display_dns": "Displayed DNS cache.",
                "renew_ip": "IP renewed.",
                "disconnect_current_network": "Disconnected from current network.",
                "connect_wifi": "Sent Wi-Fi connection request.",
                "ip_internal": "Fetched internal IP.",
                "ip_external": "Fetched external IP.",
                "ipconfig_all": "Fetched detailed network configuration.",
                "ping": "Ping test executed.",
                "open_ports": "Fetched open ports.",
                "port_owner": "Fetched process using the selected port.",
                "route_table": "Fetched route table.",
                "tracert": "Ran trace route.",
                "pathping": "Ran path ping/packet-loss diagnostics.",
                "nslookup": "Ran DNS lookup.",
                "netstat_active": "Fetched active network connections.",
                "getmac": "Fetched MAC addresses.",
                "arp_table": "Fetched ARP table.",
                "nbtstat_cache": "Fetched NetBIOS cache.",
                "nbtstat_host": "Ran NetBIOS host query.",
                "net_view": "Fetched visible network computers.",
                "netstat_binary": "Fetched network connections with executable names.",
                "wifi_profiles": "Fetched saved Wi-Fi profiles.",
                "net_scan": "Fetched local network scan results.",
                "file_sharing_on": "Enabled file sharing.",
                "file_sharing_off": "Disabled file sharing.",
                "shared_folders": "Fetched shared folders.",
                "server_online": "Checked server availability.",
                "last_login_events": "Fetched latest login attempt events.",
            }
            msg = network_msgs_ar.get(mode) if arabic else network_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "power_user_tools":
            mode = str(params.get("mode", "")).lower()
            power_msgs_ar = {
                "airplane_on": "✈️ تم تفعيل وضع الطيران.",
                "airplane_off": "✈️ تم تعطيل وضع الطيران.",
                "airplane_toggle": "✈️ تم تبديل وضع الطيران.",
                "god_mode": "🧩 تم إنشاء مجلد God Mode على سطح المكتب.",
                "invert_colors": "🎨 تم تبديل عكس الألوان.",
            }
            power_msgs_en = {
                "airplane_on": "✈️ Airplane mode enabled.",
                "airplane_off": "✈️ Airplane mode disabled.",
                "airplane_toggle": "✈️ Airplane mode toggled.",
                "god_mode": "🧩 God Mode folder created on desktop.",
                "invert_colors": "🎨 Color inversion toggled.",
            }
            msg = power_msgs_ar.get(mode) if arabic else power_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "system_power":
            mode = str(params.get("mode", "")).lower()
            sys_msgs_ar = {
                "hibernate_on": "🌙 تم تفعيل السبات.",
                "hibernate_off": "🌙 تم تعطيل السبات.",
            }
            sys_msgs_en = {
                "hibernate_on": "🌙 Hibernate enabled.",
                "hibernate_off": "🌙 Hibernate disabled.",
            }
            msg = sys_msgs_ar.get(mode) if arabic else sys_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "media_control":
            mode = str(params.get("mode", "")).lower()
            media_msgs_ar = {
                "play_pause": "تم تنفيذ تشغيل/إيقاف مؤقت.",
                "next": "تم الانتقال للمقطع التالي.",
                "previous": "تم الرجوع للمقطع السابق.",
                "stop": "تم إيقاف الوسائط.",
            }
            media_msgs_en = {
                "play_pause": "Play/Pause executed.",
                "next": "Skipped to next track.",
                "previous": "Went back to previous track.",
                "stop": "Stopped media playback.",
            }
            msg = media_msgs_ar.get(mode) if arabic else media_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "microphone_control":
            mode = str(params.get("mode", "")).lower()
            muted = bool(parsed.get("muted")) if isinstance(parsed, dict) else None
            if arabic:
                if mode in {"get", "status"}:
                    if muted is True:
                        return True, "🎤 حالة الميكروفون: مكتوم."
                    if muted is False:
                        return True, "🎤 حالة الميكروفون: غير مكتوم."
                    return True, "🎤 تم فحص حالة الميكروفون."
                if muted is True:
                    return True, "🎤 تم كتم الميكروفون."
                if muted is False:
                    return True, "🎤 تم إلغاء كتم الميكروفون."
                return True, "🎤 تم تنفيذ أمر الميكروفون."
            if mode in {"get", "status"}:
                if muted is True:
                    return True, "🎤 Microphone status: muted."
                if muted is False:
                    return True, "🎤 Microphone status: unmuted."
                return True, "🎤 Checked microphone status."
            if muted is True:
                return True, "🎤 Microphone muted."
            if muted is False:
                return True, "🎤 Microphone unmuted."
            return True, "🎤 Microphone command executed."

        if action == "process_tools":
            mode = str(params.get("mode", "")).lower()
            if mode in {"top_cpu", "top_ram"} and isinstance(parsed, dict):
                items = parsed.get("items") if isinstance(parsed.get("items"), list) else []
                top = items[0] if items else {}
                name = str(top.get("name") or "").strip()
                pid = top.get("pid")
                value = top.get("cpu" if mode == "top_cpu" else "ram_mb")
                if name and value is not None:
                    if arabic:
                        metric = "CPU" if mode == "top_cpu" else "RAM"
                        unit = "%" if mode == "top_cpu" else " MB"
                        return True, f"أعلى عملية حالياً: {name} (PID: {pid}) - {metric}: {value}{unit}."
                    metric = "CPU" if mode == "top_cpu" else "RAM"
                    unit = "%" if mode == "top_cpu" else " MB"
                    return True, f"Top process now: {name} (PID: {pid}) - {metric}: {value}{unit}."
            if mode == "app_memory_total" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                process_count = int(parsed.get("process_count") or 0)
                total_ram_mb = parsed.get("total_ram_mb")
                if query and total_ram_mb is not None:
                    if arabic:
                        return True, f"{query}: {process_count} عملية، الإجمالي {total_ram_mb} MB."
                    return True, f"{query}: {process_count} processes, total {total_ram_mb} MB."
            if mode == "app_process_count_total" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                process_count = int(parsed.get("process_count") or 0)
                total_ram_mb = parsed.get("total_ram_mb")
                if query and total_ram_mb is not None:
                    if arabic:
                        return True, f"{query}: عدد العمليات {process_count}، والمجموع {total_ram_mb} MB."
                    return True, f"{query}: {process_count} processes, combined memory {total_ram_mb} MB."
            if mode == "app_cpu_total" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                process_count = int(parsed.get("process_count") or 0)
                total_cpu = parsed.get("total_cpu_percent")
                if query and total_cpu is not None:
                    if arabic:
                        return True, f"إجمالي CPU لتطبيق {query}: {total_cpu}% عبر {process_count} عملية."
                    return True, f"Total CPU for app {query}: {total_cpu}% across {process_count} processes."
            if mode == "app_disk_total" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                process_count = int(parsed.get("process_count") or 0)
                total_disk_mb = parsed.get("total_disk_mb")
                if query and total_disk_mb is not None:
                    if arabic:
                        return True, f"إجمالي نشاط القرص لتطبيق {query}: {total_disk_mb} MB عبر {process_count} عملية."
                    return True, f"Total disk activity for app {query}: {total_disk_mb} MB across {process_count} processes."
            if mode == "app_network_total" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                process_count = int(parsed.get("process_count") or 0)
                total_connections = int(parsed.get("total_connections") or 0)
                established_connections = int(parsed.get("established_connections") or 0)
                unique_remote_ips = int(parsed.get("unique_remote_ips") or 0)
                if query:
                    if arabic:
                        return True, (
                            f"نشاط الشبكة لتطبيق {query}: {total_connections} اتصال "
                            f"(منها {established_connections} نشط) عبر {process_count} عملية، "
                            f"ومع {unique_remote_ips} عناوين IP بعيدة."
                        )
                    return True, (
                        f"Network activity for app {query}: {total_connections} connections "
                        f"({established_connections} established) across {process_count} processes, "
                        f"with {unique_remote_ips} unique remote IPs."
                    )
            if mode == "app_resource_summary" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                if query:
                    process_count = int(parsed.get("process_count") or 0)
                    total_ram_mb = parsed.get("total_ram_mb")
                    total_cpu = parsed.get("total_cpu_percent")
                    total_disk_mb = parsed.get("total_disk_mb")
                    total_connections = int(parsed.get("total_connections") or 0)
                    if arabic:
                        return True, (
                            f"ملخص {query}: RAM={total_ram_mb} MB، CPU={total_cpu}%، "
                            f"Disk={total_disk_mb} MB، Network={total_connections} اتصال، "
                            f"وعبر {process_count} عملية."
                        )
                    return True, (
                        f"Summary for {query}: RAM={total_ram_mb} MB, CPU={total_cpu}%, "
                        f"Disk={total_disk_mb} MB, Network={total_connections} connections, "
                        f"across {process_count} processes."
                    )
            if mode == "app_compare" and isinstance(parsed, dict):
                left = parsed.get("left") if isinstance(parsed.get("left"), dict) else {}
                right = parsed.get("right") if isinstance(parsed.get("right"), dict) else {}
                winners = parsed.get("winners") if isinstance(parsed.get("winners"), dict) else {}
                recommendations = parsed.get("recommendations") if isinstance(parsed.get("recommendations"), list) else []
                lq = str(left.get("query") or params.get("name") or "").strip()
                rq = str(right.get("query") or params.get("target") or "").strip()
                if lq and rq:
                    lram = left.get("total_ram_mb")
                    rram = right.get("total_ram_mb")
                    lcpu = left.get("total_cpu_percent")
                    rcpu = right.get("total_cpu_percent")
                    ldisk = left.get("total_disk_mb")
                    rdisk = right.get("total_disk_mb")
                    lnet = left.get("total_connections")
                    rnet = right.get("total_connections")
                    ram_w = str(winners.get("ram") or "equal")
                    cpu_w = str(winners.get("cpu") or "equal")
                    disk_w = str(winners.get("disk") or "equal")
                    net_w = str(winners.get("network") or "equal")
                    rec_line = ", ".join(str(r) for r in recommendations[:4]) if recommendations else ""
                    practical_tips: list[str] = []
                    for rec in recommendations:
                        token = str(rec or "").strip().lower()
                        if token.startswith("ram_hotspot="):
                            appn = rec.split("=", 1)[1] if "=" in rec else "this app"
                            practical_tips.append(
                                f"خفّف RAM بإغلاق نوافذ/تبويبات {appn} غير الضرورية"
                                if arabic
                                else f"Reduce RAM first by closing unnecessary {appn} windows/tabs"
                            )
                        elif token.startswith("cpu_hotspot="):
                            appn = rec.split("=", 1)[1] if "=" in rec else "this app"
                            practical_tips.append(
                                f"خفّف CPU بإيقاف مهام الخلفية داخل {appn}"
                                if arabic
                                else f"Reduce CPU by stopping heavy background tasks in {appn}"
                            )
                        elif token.startswith("disk_hotspot="):
                            appn = rec.split("=", 1)[1] if "=" in rec else "this app"
                            practical_tips.append(
                                f"خفّف Disk بإيقاف تنزيلات/فهرسة {appn} مؤقتاً"
                                if arabic
                                else f"Reduce Disk by pausing downloads/indexing in {appn}"
                            )
                        elif token.startswith("network_hotspot="):
                            appn = rec.split("=", 1)[1] if "=" in rec else "this app"
                            practical_tips.append(
                                f"خفّف Network بتقليل مزامنة أو اتصالات {appn}"
                                if arabic
                                else f"Reduce Network by limiting sync/connections in {appn}"
                            )
                    practical = " | ".join(practical_tips[:2])
                    if arabic:
                        return True, (
                            f"مقارنة {lq} vs {rq}: RAM {lram}/{rram} MB، CPU {lcpu}/{rcpu}%، "
                            f"Disk {ldisk}/{rdisk} MB، Network {lnet}/{rnet}. "
                            f"الأثقل: RAM={ram_w}، CPU={cpu_w}، Disk={disk_w}، Network={net_w}."
                            + (f" التوصية: {rec_line}." if rec_line else "")
                            + (f" إجراء عملي: {practical}." if practical else "")
                        )
                    return True, (
                        f"Comparison {lq} vs {rq}: RAM {lram}/{rram} MB, CPU {lcpu}/{rcpu}%, "
                        f"Disk {ldisk}/{rdisk} MB, Network {lnet}/{rnet}. "
                        f"Heavier: RAM={ram_w}, CPU={cpu_w}, Disk={disk_w}, Network={net_w}."
                        + (f" Recommendation: {rec_line}." if rec_line else "")
                        + (f" Practical action: {practical}." if practical else "")
                    )
            if mode == "app_reduce_ram_plan" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                total_ram_mb = parsed.get("total_ram_mb")
                process_count = int(parsed.get("process_count") or 0)
                reclaimable = parsed.get("reclaimable_mb_estimate")
                top = parsed.get("top_processes") if isinstance(parsed.get("top_processes"), list) else []
                top_hint = ""
                if top:
                    p0 = top[0] if isinstance(top[0], dict) else {}
                    p1 = top[1] if len(top) > 1 and isinstance(top[1], dict) else {}
                    n0 = str(p0.get("name") or "")
                    m0 = p0.get("ram_mb")
                    n1 = str(p1.get("name") or "")
                    m1 = p1.get("ram_mb")
                    if n0 and m0 is not None:
                        top_hint = f" أكبر عملية: {n0} ({m0} MB)." if arabic else f" Top process: {n0} ({m0} MB)."
                    if n1 and m1 is not None:
                        top_hint += (
                            f" ثاني أكبر: {n1} ({m1} MB)."
                            if arabic
                            else f" Second largest: {n1} ({m1} MB)."
                        )
                if query and total_ram_mb is not None:
                    if arabic:
                        return True, (
                            f"خطة تخفيف RAM لتطبيق {query}: الإجمالي {total_ram_mb} MB عبر {process_count} عملية. "
                            f"تقدير الاسترجاع السريع: {reclaimable} MB."
                            f"{top_hint} الإجراءات: اغلاق النوافذ/التبويبات الزائدة ثم إعادة تشغيل التطبيق إذا لزم."
                        )
                    return True, (
                        f"RAM reduction plan for {query}: total {total_ram_mb} MB across {process_count} processes. "
                        f"Quick reclaim estimate: {reclaimable} MB."
                        f"{top_hint} Actions: close extra windows/tabs, then restart app if needed."
                    )
            if mode == "app_reduce_ram_execute" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                killed_count = int(parsed.get("killed_count") or 0)
                protected_name = str(parsed.get("protected_name") or "")
                protected_pid = parsed.get("protected_pid")
                dry_run = bool(parsed.get("dry_run"))
                max_kill = parsed.get("max_kill")
                if query:
                    if arabic:
                        return True, (
                            (f"معاينة تخفيف RAM لتطبيق {query}. " if dry_run else f"تم تنفيذ تخفيف RAM لتطبيق {query}. ")
                            + f"تم {'تحديد' if dry_run else 'إغلاق'} {killed_count} عملية ثانوية"
                            + (f" (حد أقصى {max_kill})" if max_kill is not None else "")
                            + f"، مع إبقاء العملية الرئيسية {protected_name} (PID: {protected_pid})."
                        )
                    return True, (
                        (f"Previewed RAM reduction for {query}. " if dry_run else f"Executed RAM reduction for {query}. ")
                        + f"{'Selected' if dry_run else 'Closed'} {killed_count} secondary processes"
                        + (f" (max_kill={max_kill})" if max_kill is not None else "")
                        + f", while keeping main process {protected_name} (PID: {protected_pid})."
                    )
            if mode == "app_reduce_cpu_plan" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                total_cpu = parsed.get("total_cpu_percent")
                process_count = int(parsed.get("process_count") or 0)
                reclaimable = parsed.get("reclaimable_cpu_estimate")
                if query and total_cpu is not None:
                    if arabic:
                        return True, (
                            f"خطة تخفيف CPU لتطبيق {query}: الإجمالي {total_cpu}% عبر {process_count} عملية. "
                            f"تقدير التخفيض السريع: {reclaimable}%."
                        )
                    return True, (
                        f"CPU reduction plan for {query}: total {total_cpu}% across {process_count} processes. "
                        f"Quick reduction estimate: {reclaimable}%."
                    )
            if mode == "app_reduce_cpu_execute" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                killed_count = int(parsed.get("killed_count") or 0)
                protected_name = str(parsed.get("protected_name") or "")
                protected_pid = parsed.get("protected_pid")
                threshold = parsed.get("threshold")
                dry_run = bool(parsed.get("dry_run"))
                if query:
                    if arabic:
                        return True, (
                            (f"معاينة تخفيف CPU لتطبيق {query} (عتبة {threshold}%): " if dry_run else f"تم تنفيذ تخفيف CPU لتطبيق {query} (عتبة {threshold}%): ")
                            + f"{'سيتم' if dry_run else 'تم'} التعامل مع {killed_count} عملية ثانوية "
                            + f"مع إبقاء العملية الرئيسية {protected_name} (PID: {protected_pid})."
                        )
                    return True, (
                        (f"Previewed CPU reduction for {query} (threshold {threshold}%): " if dry_run else f"Executed CPU reduction for {query} (threshold {threshold}%): ")
                        + f"{'would handle' if dry_run else 'handled'} {killed_count} secondary processes, "
                        + f"while keeping main process {protected_name} (PID: {protected_pid})."
                    )
            if mode == "app_reduce_disk_plan" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                total_disk = parsed.get("total_disk_mb")
                process_count = int(parsed.get("process_count") or 0)
                reclaimable = parsed.get("reclaimable_disk_estimate")
                if query and total_disk is not None:
                    if arabic:
                        return True, (
                            f"خطة تخفيف Disk لتطبيق {query}: الإجمالي {total_disk} MB عبر {process_count} عملية. "
                            f"تقدير التخفيض السريع: {reclaimable} MB."
                        )
                    return True, (
                        f"Disk reduction plan for {query}: total {total_disk} MB across {process_count} processes. "
                        f"Quick reduction estimate: {reclaimable} MB."
                    )
            if mode == "app_reduce_disk_execute" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                killed_count = int(parsed.get("killed_count") or 0)
                protected_name = str(parsed.get("protected_name") or "")
                protected_pid = parsed.get("protected_pid")
                threshold = parsed.get("threshold")
                dry_run = bool(parsed.get("dry_run"))
                if query:
                    if arabic:
                        return True, (
                            (f"معاينة تخفيف Disk لتطبيق {query} (عتبة {threshold} MB): " if dry_run else f"تم تنفيذ تخفيف Disk لتطبيق {query} (عتبة {threshold} MB): ")
                            + f"{'سيتم' if dry_run else 'تم'} التعامل مع {killed_count} عملية ثانوية "
                            + f"مع إبقاء العملية الرئيسية {protected_name} (PID: {protected_pid})."
                        )
                    return True, (
                        (f"Previewed Disk reduction for {query} (threshold {threshold} MB): " if dry_run else f"Executed Disk reduction for {query} (threshold {threshold} MB): ")
                        + f"{'would handle' if dry_run else 'handled'} {killed_count} secondary processes, "
                        + f"while keeping main process {protected_name} (PID: {protected_pid})."
                    )
            if mode == "app_reduce_network_plan" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                total_conn = int(parsed.get("total_connections") or 0)
                process_count = int(parsed.get("process_count") or 0)
                reclaimable = int(parsed.get("reclaimable_network_estimate") or 0)
                if query:
                    if arabic:
                        return True, (
                            f"خطة تخفيف Network لتطبيق {query}: الإجمالي {total_conn} اتصال عبر {process_count} عملية. "
                            f"تقدير التخفيض السريع: {reclaimable} اتصال."
                        )
                    return True, (
                        f"Network reduction plan for {query}: total {total_conn} connections across {process_count} processes. "
                        f"Quick reduction estimate: {reclaimable} connections."
                    )
            if mode == "app_reduce_network_execute" and isinstance(parsed, dict):
                query = str(parsed.get("query") or params.get("name") or "").strip()
                killed_count = int(parsed.get("killed_count") or 0)
                protected_name = str(parsed.get("protected_name") or "")
                protected_pid = parsed.get("protected_pid")
                threshold = parsed.get("threshold")
                dry_run = bool(parsed.get("dry_run"))
                if query:
                    if arabic:
                        return True, (
                            (f"معاينة تخفيف Network لتطبيق {query} (عتبة {threshold} اتصالات): " if dry_run else f"تم تنفيذ تخفيف Network لتطبيق {query} (عتبة {threshold} اتصالات): ")
                            + f"{'سيتم' if dry_run else 'تم'} التعامل مع {killed_count} عملية ثانوية "
                            + f"مع إبقاء العملية الرئيسية {protected_name} (PID: {protected_pid})."
                        )
                    return True, (
                        (f"Previewed Network reduction for {query} (threshold {threshold} connections): " if dry_run else f"Executed Network reduction for {query} (threshold {threshold} connections): ")
                        + f"{'would handle' if dry_run else 'handled'} {killed_count} secondary processes, "
                        + f"while keeping main process {protected_name} (PID: {protected_pid})."
                    )
            if mode == "kill_high_cpu" and isinstance(parsed, dict):
                count = int(parsed.get("count") or 0)
                threshold = parsed.get("threshold")
                dry_run = bool(parsed.get("dry_run"))
                max_kill = parsed.get("max_kill")
                if arabic:
                    return True, (
                        (f"معاينة kill_high_cpu عند عتبة {threshold}%: " if dry_run else f"تم تنفيذ kill_high_cpu عند عتبة {threshold}%: ")
                        + f"{'سيتم' if dry_run else 'تم'} التعامل مع {count} عملية"
                        + (f" (حد أقصى {max_kill})" if max_kill is not None else "")
                        + "."
                    )
                return True, (
                    (f"Previewed kill_high_cpu at threshold {threshold}%: " if dry_run else f"Executed kill_high_cpu at threshold {threshold}%: ")
                    + f"{'would handle' if dry_run else 'handled'} {count} process(es)"
                    + (f" (max_kill={max_kill})" if max_kill is not None else "")
                    + "."
                )
            if mode == "app_reduce" and isinstance(parsed, dict):
                stage = str(parsed.get("stage") or params.get("stage") or "").strip().lower()
                resource = str(parsed.get("resource") or params.get("resource") or "resource").strip().lower()
                if arabic:
                    if stage == "plan":
                        return True, f"🧠 خطة تخفيف {resource} جاهزة."
                    return True, f"✅ تم تنفيذ تخفيف {resource}."
                if stage == "plan":
                    return True, f"🧠 {resource} reduction plan is ready."
                return True, f"✅ {resource} reduction executed."
            process_msgs_ar = {
                "restart_explorer": "تمت إعادة تشغيل واجهة ويندوز (Explorer).",
            }
            process_msgs_en = {
                "restart_explorer": "Windows Explorer has been restarted.",
            }
            msg = process_msgs_ar.get(mode) if arabic else process_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "service_tools":
            mode = str(params.get("mode", "")).lower()
            svc = str(params.get("name", "") or "").strip()
            startup = str(params.get("startup", "") or "").strip()
            service_msgs_ar = {
                "start": f"تم تشغيل الخدمة: {svc or 'المحددة'}.",
                "stop": f"تم إيقاف الخدمة: {svc or 'المحددة'}.",
                "restart": f"تمت إعادة تشغيل الخدمة: {svc or 'المحددة'}.",
                "list": "تم جلب قائمة الخدمات.",
                "describe": f"تم جلب وصف الخدمة: {svc or 'المحددة'}.",
                "dependencies": f"تم جلب تبعيات الخدمة: {svc or 'المحددة'}.",
                "startup": f"تم تعديل نوع تشغيل الخدمة: {svc or 'المحددة'} إلى {startup or 'المطلوب'}.",
            }
            service_msgs_en = {
                "start": f"Service started: {svc or 'target service'}.",
                "stop": f"Service stopped: {svc or 'target service'}.",
                "restart": f"Service restarted: {svc or 'target service'}.",
                "list": "Fetched services list.",
                "describe": f"Fetched service description: {svc or 'target service'}.",
                "dependencies": f"Fetched service dependencies: {svc or 'target service'}.",
                "startup": f"Updated service startup type: {svc or 'target service'} -> {startup or 'target mode'}.",
            }
            msg = service_msgs_ar.get(mode) if arabic else service_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "security_tools":
            mode = str(params.get("mode", "")).lower()
            security_msgs_ar = {
                "firewall_status": "تم جلب حالة جدار الحماية.",
                "firewall_enable": "تم تفعيل جدار الحماية.",
                "firewall_disable": "تم تعطيل جدار الحماية.",
                "block_port": "تم إضافة قاعدة حظر منفذ في الجدار الناري.",
                "unblock_rule": "تم حذف قاعدة من الجدار الناري.",
                "disable_usb": "تم تعطيل منافذ USB.",
                "enable_usb": "تم تفعيل منافذ USB.",
                "disable_camera": "تم تعطيل الكاميرا.",
                "enable_camera": "تم تفعيل الكاميرا.",
                "recent_files_list": "تم جلب قائمة الملفات المفتوحة مؤخراً.",
                "recent_files_clear": "تم مسح قائمة الملفات المفتوحة مؤخراً.",
                "close_remote_sessions": "تم تنفيذ إغلاق الجلسات البعيدة.",
                "intrusion_summary": "تم تجهيز ملخص محاولات الدخول الفاشلة.",
            }
            security_msgs_en = {
                "firewall_status": "Fetched firewall status.",
                "firewall_enable": "Firewall enabled.",
                "firewall_disable": "Firewall disabled.",
                "block_port": "Added firewall block-port rule.",
                "unblock_rule": "Removed firewall rule.",
                "disable_usb": "USB ports disabled.",
                "enable_usb": "USB ports enabled.",
                "disable_camera": "Camera disabled.",
                "enable_camera": "Camera enabled.",
                "recent_files_list": "Fetched recent files list.",
                "recent_files_clear": "Cleared recent files list.",
                "close_remote_sessions": "Executed remote sessions close.",
                "intrusion_summary": "Prepared failed-login intrusion summary.",
            }
            msg = security_msgs_ar.get(mode) if arabic else security_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "background_tools":
            mode = str(params.get("mode", "")).lower()
            background_msgs_ar = {
                "count_background": "تم جلب عدد تطبيقات الخلفية.",
                "list_visible_windows": "تم جلب قائمة التطبيقات المرئية.",
                "list_minimized_windows": "تم جلب قائمة التطبيقات المصغرة.",
                "ghost_apps": "تم جلب التطبيقات الخلفية الثقيلة.",
                "network_usage_per_app": "تم جلب التطبيقات التي تستخدم الشبكة الآن.",
                "camera_usage_now": "تم فحص التطبيقات التي تستخدم الكاميرا.",
                "mic_usage_now": "تم فحص التطبيقات التي تستخدم الميكروفون.",
                "wake_lock_apps": "تم جلب التطبيقات التي تمنع السكون.",
                "process_paths": "تم جلب مسارات التطبيقات الشغالة.",
            }
            background_msgs_en = {
                "count_background": "Fetched background processes count.",
                "list_visible_windows": "Fetched visible windows list.",
                "list_minimized_windows": "Fetched minimized windows list.",
                "ghost_apps": "Fetched heavy headless/background apps.",
                "network_usage_per_app": "Fetched apps currently using network.",
                "camera_usage_now": "Checked apps currently using camera.",
                "mic_usage_now": "Checked apps currently using microphone.",
                "wake_lock_apps": "Fetched apps blocking sleep.",
                "process_paths": "Fetched running app paths.",
            }
            msg = background_msgs_ar.get(mode) if arabic else background_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "startup_tools":
            mode = str(params.get("mode", "")).lower()
            item = str(params.get("name", "") or "").strip()
            startup_msgs_ar = {
                "startup_list": "تم جلب قائمة برامج بدء التشغيل.",
                "startup_impact_time": "تم جلب وقت تأثير بدء التشغيل.",
                "registry_startups": "تم جلب برامج بدء التشغيل من السجل.",
                "folder_startups": "تم جلب برامج بدء التشغيل من مجلد Startup.",
                "signature_check": "تم فحص أمان/توقيع برامج بدء التشغيل.",
                "disable": f"تم تعطيل برنامج بدء التشغيل: {item or 'المحدد'}.",
                "enable": f"تم تفعيل برنامج بدء التشغيل: {item or 'المحدد'}.",
            }
            startup_msgs_en = {
                "startup_list": "Fetched startup apps list.",
                "startup_impact_time": "Fetched startup impact time.",
                "registry_startups": "Fetched registry startup entries.",
                "folder_startups": "Fetched startup folder entries.",
                "signature_check": "Checked startup apps signatures/security.",
                "disable": f"Disabled startup app: {item or 'target item'}.",
                "enable": f"Enabled startup app: {item or 'target item'}.",
            }
            msg = startup_msgs_ar.get(mode) if arabic else startup_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "performance_tools":
            mode = str(params.get("mode", "")).lower()
            if mode in {"total_cpu_percent", "total_ram_percent"} and isinstance(parsed, dict):
                percent = parsed.get("percent")
                if percent is not None:
                    if arabic:
                        label = "المعالج" if mode == "total_cpu_percent" else "الرام"
                        return True, f"نسبة الاستهلاك الحالية ({label}): {percent}%."
                    label = "CPU" if mode == "total_cpu_percent" else "RAM"
                    return True, f"Current {label} usage: {percent}%."
            perf_msgs_ar = {
                "top_cpu": "تم جلب أعلى 5 عمليات استهلاكاً للمعالج.",
                "top_ram": "تم جلب أعلى 5 عمليات استهلاكاً للرام.",
                "top_disk": "تم جلب أعلى 5 عمليات استهلاكاً للقرص.",
                "cpu_clock": "تم جلب سرعة المعالج الحالية.",
                "available_ram": "تم جلب حجم الذاكرة المتاحة.",
                "pagefile_used": "تم جلب استهلاك ملف التبادل (Page File).",
            }
            perf_msgs_en = {
                "top_cpu": "Fetched top 5 CPU-consuming processes.",
                "top_ram": "Fetched top 5 RAM-consuming processes.",
                "top_disk": "Fetched top 5 disk-consuming processes.",
                "cpu_clock": "Fetched current CPU clock speed.",
                "available_ram": "Fetched available RAM.",
                "pagefile_used": "Fetched page file usage.",
            }
            msg = perf_msgs_ar.get(mode) if arabic else perf_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "vision_tools":
            mode = str(params.get("mode", "")).lower()
            if mode == "describe_screen":
                preview = ""
                summary = ""
                top_app = ""
                if isinstance(parsed, dict):
                    preview = str(parsed.get("ocr_text_preview") or "").strip()
                    summary = str(parsed.get("ui_summary") or "").strip()
                    top_app = str(parsed.get("top_app") or "").strip()
                if arabic:
                    if summary or top_app:
                        msg = f"تحليل الشاشة جاهز{f'، الأعلى: {top_app}' if top_app else ''}."
                        if summary:
                            msg += f" {summary[:180]}"
                        return True, msg
                    if preview:
                        return True, f"تحليل الشاشة جاهز. مقتطف: {preview[:180]}"
                    return True, "تحليل الشاشة جاهز."
                if summary or top_app:
                    msg = f"Screen analyzed{f', top app: {top_app}' if top_app else ''}."
                    if summary:
                        msg += f" {summary[:180]}"
                    return True, msg
                if preview:
                    return True, f"Screen analyzed. Preview: {preview[:180]}"
                return True, "Screen analyzed."

        if action == "screenshot_tools":
            mode = str(params.get("mode", "")).lower()
            out_path = str(parsed.get("path") or "").strip() if isinstance(parsed, dict) else ""
            if arabic:
                if mode == "window_active":
                    return True, ("📸 تم أخذ لقطة للنافذة الحالية." + (f"\n{out_path}" if out_path else ""))
                if mode == "full":
                    return True, ("📸 تم أخذ لقطة للشاشة." + (f"\n{out_path}" if out_path else ""))
                if mode == "region":
                    return True, ("📸 تم أخذ لقطة للمنطقة المحددة." + (f"\n{out_path}" if out_path else ""))
                if mode == "snipping_tool":
                    return True, "✂️ تم فتح أداة القص."
            else:
                if mode == "window_active":
                    return True, ("📸 Captured the active window." + (f"\n{out_path}" if out_path else ""))
                if mode == "full":
                    return True, ("📸 Captured the screen." + (f"\n{out_path}" if out_path else ""))
                if mode == "region":
                    return True, ("📸 Captured the selected region." + (f"\n{out_path}" if out_path else ""))
                if mode == "snipping_tool":
                    return True, "✂️ Opened Snipping Tool."

        if action == "ui_tools":
            mode = str(params.get("mode", "")).lower()
            ui_msgs_ar = {
                "night_light_on": "تم تفعيل الوضع الليلي.",
                "night_light_off": "تم تعطيل الوضع الليلي.",
                "desktop_icons_show": "تم إظهار أيقونات سطح المكتب.",
                "desktop_icons_hide": "تم إخفاء أيقونات سطح المكتب.",
                "open_display_resolution": "تم فتح إعدادات دقة الشاشة.",
                "open_display_rotation": "تم فتح إعدادات تدوير الشاشة.",
            }
            ui_msgs_en = {
                "night_light_on": "Night light enabled.",
                "night_light_off": "Night light disabled.",
                "desktop_icons_show": "Desktop icons shown.",
                "desktop_icons_hide": "Desktop icons hidden.",
                "open_display_resolution": "Opened display resolution settings.",
                "open_display_rotation": "Opened display rotation settings.",
            }
            msg = ui_msgs_ar.get(mode) if arabic else ui_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "browser_control":
            mode = str(params.get("mode", "")).lower()
            browser_msgs_ar = {
                "new_tab": "تم فتح تبويب جديد.",
                "close_tab": "تم إغلاق التبويب الحالي.",
                "reopen_tab": "تمت إعادة فتح آخر تبويب مغلق.",
                "next_tab": "تم الانتقال للتبويب التالي.",
                "prev_tab": "تم الانتقال للتبويب السابق.",
                "reload": "تم تحديث الصفحة.",
                "incognito": "تم فتح نافذة التصفح الخفي.",
                "home": "تم الذهاب إلى صفحة البداية.",
                "history": "تم فتح سجل التصفح.",
                "downloads": "تم فتح تنزيلات المتصفح.",
                "find": "تم فتح البحث داخل الصفحة.",
                "zoom_in": "تم تكبير الصفحة.",
                "zoom_out": "تم تصغير الصفحة.",
                "zoom_reset": "تمت إعادة الزوم إلى 100%.",
                "save_pdf": "تم فتح نافذة حفظ الصفحة PDF.",
            }
            browser_msgs_en = {
                "new_tab": "Opened a new browser tab.",
                "close_tab": "Closed current browser tab.",
                "reopen_tab": "Reopened last closed tab.",
                "next_tab": "Moved to next tab.",
                "prev_tab": "Moved to previous tab.",
                "reload": "Reloaded page.",
                "incognito": "Opened incognito/private window.",
                "home": "Opened browser home page.",
                "history": "Opened browser history.",
                "downloads": "Opened browser downloads.",
                "find": "Opened Find in page.",
                "zoom_in": "Zoomed in.",
                "zoom_out": "Zoomed out.",
                "zoom_reset": "Reset zoom to 100%.",
                "save_pdf": "Opened save as PDF flow.",
            }
            msg = browser_msgs_ar.get(mode) if arabic else browser_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "task_tools":
            mode = str(params.get("mode", "")).lower()
            name = str(params.get("name", "") or "").strip()
            task_msgs_ar = {
                "list": "تم جلب قائمة المهام المجدولة.",
                "running": "تم جلب المهام المجدولة الجارية.",
                "last_run": "تم جلب آخر وقت تشغيل للمهام المجدولة.",
                "run": f"تم تشغيل المهمة المجدولة: {name or 'المحددة'}.",
                "end": f"تم إنهاء المهمة المجدولة: {name or 'المحددة'}.",
                "enable": f"تم تمكين المهمة المجدولة: {name or 'المحددة'}.",
                "disable": f"تم تعطيل المهمة المجدولة: {name or 'المحددة'}.",
                "delete": f"تم حذف المهمة المجدولة: {name or 'المحددة'}.",
                "create": f"تم إنشاء مهمة مجدولة: {name or 'جديدة'}.",
            }
            task_msgs_en = {
                "list": "Fetched scheduled tasks list.",
                "running": "Fetched running scheduled tasks.",
                "last_run": "Fetched scheduled tasks last run times.",
                "run": f"Ran scheduled task: {name or 'target task'}.",
                "end": f"Stopped scheduled task: {name or 'target task'}.",
                "enable": f"Enabled scheduled task: {name or 'target task'}.",
                "disable": f"Disabled scheduled task: {name or 'target task'}.",
                "delete": f"Deleted scheduled task: {name or 'target task'}.",
                "create": f"Created scheduled task: {name or 'new task'}.",
            }
            msg = task_msgs_ar.get(mode) if arabic else task_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "user_tools":
            mode = str(params.get("mode", "")).lower()
            uname = str(params.get("username", "") or "").strip()
            user_msgs_ar = {
                "list": "تم جلب قائمة المستخدمين.",
                "create": f"تم إنشاء المستخدم: {uname or 'الجديد'}.",
                "delete": f"تم حذف المستخدم: {uname or 'المحدد'}.",
                "set_password": f"تم تغيير كلمة مرور المستخدم: {uname or 'المحدد'}.",
                "set_type": f"تم تحديث نوع المستخدم: {uname or 'المحدد'}.",
            }
            user_msgs_en = {
                "list": "Fetched users list.",
                "create": f"Created user: {uname or 'new user'}.",
                "delete": f"Deleted user: {uname or 'target user'}.",
                "set_password": f"Updated password for user: {uname or 'target user'}.",
                "set_type": f"Updated user type for: {uname or 'target user'}.",
            }
            msg = user_msgs_ar.get(mode) if arabic else user_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "update_tools":
            mode = str(params.get("mode", "")).lower()
            kb = str(params.get("target", "") or "").strip()
            update_msgs_ar = {
                "list_updates": "تم جلب قائمة تحديثات ويندوز.",
                "last_update_time": "تم جلب وقت آخر تحديث للنظام.",
                "check_updates": "تم بدء فحص تحديثات ويندوز.",
                "winsxs_cleanup": "تم تشغيل تنظيف ملفات WinSxS.",
                "stop_background_updates": "تم إيقاف خدمات تحديثات ويندوز في الخلفية.",
                "install_kb": f"تم إرسال طلب تثبيت التحديث: {kb or 'KB المطلوب'}.",
            }
            update_msgs_en = {
                "list_updates": "Fetched Windows updates list.",
                "last_update_time": "Fetched last system update time.",
                "check_updates": "Started Windows Update scan.",
                "winsxs_cleanup": "Started WinSxS cleanup.",
                "stop_background_updates": "Stopped background Windows Update services.",
                "install_kb": f"Sent install request for update: {kb or 'target KB'}.",
            }
            msg = update_msgs_ar.get(mode) if arabic else update_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "remote_tools":
            mode = str(params.get("mode", "")).lower()
            host = str(params.get("host", "") or "").strip()
            remote_msgs_ar = {
                "rdp_open": "تم فتح Remote Desktop.",
                "vpn_connect": f"تمت محاولة الاتصال بـ VPN: {host or 'المحدد'}.",
                "vpn_disconnect": "تم تنفيذ قطع اتصال VPN.",
            }
            remote_msgs_en = {
                "rdp_open": "Opened Remote Desktop.",
                "vpn_connect": f"Attempted VPN connect: {host or 'target connection'}.",
                "vpn_disconnect": "Executed VPN disconnect.",
            }
            msg = remote_msgs_ar.get(mode) if arabic else remote_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "disk_tools":
            mode = str(params.get("mode", "")).lower()
            drv = str(params.get("drive", "") or "").strip()
            disk_msgs_ar = {
                "smart_status": "تم جلب حالة صحة الأقراص (SMART).",
                "temp_files_clean": "تم تنظيف الملفات المؤقتة.",
                "prefetch_clean": "تم تنظيف ملفات Prefetch.",
                "logs_clean": "تم تنظيف سجلات ويندوز.",
                "disk_usage": "تم جلب استهلاك ومساحة الأقراص.",
                "defrag": f"تم بدء إلغاء التجزئة للقرص: {drv or 'C:'}.",
                "chkdsk_scan": f"تم بدء فحص القرص: {drv or 'C:'}.",
            }
            disk_msgs_en = {
                "smart_status": "Fetched disk SMART/health status.",
                "temp_files_clean": "Cleaned temp files.",
                "prefetch_clean": "Cleaned Prefetch files.",
                "logs_clean": "Cleaned Windows logs.",
                "disk_usage": "Fetched disk usage and free space.",
                "defrag": f"Started defrag on drive: {drv or 'C:'}.",
                "chkdsk_scan": f"Started disk check on drive: {drv or 'C:'}.",
            }
            msg = disk_msgs_ar.get(mode) if arabic else disk_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "registry_tools":
            mode = str(params.get("mode", "")).lower()
            key = str(params.get("key", "") or "").strip()
            reg_msgs_ar = {
                "query": f"تم الاستعلام عن السجل: {key or 'المفتاح المحدد'}.",
                "add_key": f"تمت إضافة مفتاح سجل: {key or 'المفتاح المحدد'}.",
                "delete_key": f"تم حذف مفتاح سجل: {key or 'المفتاح المحدد'}.",
                "set_value": f"تم تحديث قيمة في السجل: {key or 'المفتاح المحدد'}.",
                "backup": "تم إنشاء نسخة احتياطية للسجل.",
                "restore": "تم تنفيذ استعادة نسخة السجل.",
            }
            reg_msgs_en = {
                "query": f"Queried registry key: {key or 'target key'}.",
                "add_key": f"Added registry key: {key or 'target key'}.",
                "delete_key": f"Deleted registry key: {key or 'target key'}.",
                "set_value": f"Updated registry value under: {key or 'target key'}.",
                "backup": "Created registry backup.",
                "restore": "Executed registry restore.",
            }
            msg = reg_msgs_ar.get(mode) if arabic else reg_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "search_tools":
            mode = str(params.get("mode", "")).lower()
            search_msgs_ar = {
                "search_text": "تم البحث عن النص داخل الملفات.",
                "files_larger_than": "تم جلب الملفات الأكبر من الحجم المطلوب.",
                "modified_today": "تم جلب الملفات المعدلة اليوم.",
                "find_images": "تم جلب الصور الموجودة في المسار.",
                "find_videos": "تم جلب الفيديوهات الموجودة في المسار.",
                "count_files": "تم حساب عدد الملفات.",
                "search_windows_content": "تم البحث في محتوى النوافذ المفتوحة.",
            }
            search_msgs_en = {
                "search_text": "Searched text inside files.",
                "files_larger_than": "Fetched files larger than requested size.",
                "modified_today": "Fetched files modified today.",
                "find_images": "Fetched image files in target path.",
                "find_videos": "Fetched video files in target path.",
                "count_files": "Counted files in target path.",
                "search_windows_content": "Searched in open windows content.",
            }
            msg = search_msgs_ar.get(mode) if arabic else search_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "web_tools":
            mode = str(params.get("mode", "")).lower()
            web_msgs_ar = {
                "open_url": "تم فتح الرابط.",
                "download_file": "تم بدء تنزيل الملف.",
                "weather": "تم جلب حالة الطقس.",
            }
            web_msgs_en = {
                "open_url": "Opened URL.",
                "download_file": "Started file download.",
                "weather": "Fetched weather information.",
            }
            msg = web_msgs_ar.get(mode) if arabic else web_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "api_tools":
            mode = str(params.get("mode", "")).lower()
            api_msgs_ar = {
                "currency": "تم جلب أسعار العملات.",
                "weather_city": "تم جلب حالة الطقس للمدينة.",
                "translate_quick": "تم تنفيذ الترجمة السريعة.",
            }
            api_msgs_en = {
                "currency": "Fetched currency prices.",
                "weather_city": "Fetched city weather.",
                "translate_quick": "Executed quick translation.",
            }
            msg = api_msgs_ar.get(mode) if arabic else api_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "browser_deep_tools":
            mode = str(params.get("mode", "")).lower()
            deep_msgs_ar = {
                "multi_open": "تم فتح مجموعة الروابط.",
                "clear_chrome_cache": "تم مسح كاش Chrome.",
                "clear_edge_cache": "تم مسح كاش Edge.",
            }
            deep_msgs_en = {
                "multi_open": "Opened multiple URLs.",
                "clear_chrome_cache": "Cleared Chrome cache.",
                "clear_edge_cache": "Cleared Edge cache.",
            }
            msg = deep_msgs_ar.get(mode) if arabic else deep_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "office_tools":
            mode = str(params.get("mode", "")).lower()
            office_msgs_ar = {
                "open_word_new": "تم فتح ملف Word جديد.",
                "docx_to_pdf": "تم تنفيذ تحويل DOCX إلى PDF.",
                "silent_print": "تم إرسال الملف للطباعة.",
            }
            office_msgs_en = {
                "open_word_new": "Opened a new Word document.",
                "docx_to_pdf": "Executed DOCX to PDF conversion.",
                "silent_print": "Sent file to printer.",
            }
            msg = office_msgs_ar.get(mode) if arabic else office_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "driver_tools":
            mode = str(params.get("mode", "")).lower()
            driver_msgs_ar = {
                "drivers_list": "تم جلب قائمة التعريفات المثبتة.",
                "drivers_backup": "تم بدء أخذ نسخة احتياطية للتعريفات.",
                "updates_pending": "تم جلب التحديثات المعلقة.",
                "drivers_issues": "تم جلب التعريفات التي فيها مشاكل.",
            }
            driver_msgs_en = {
                "drivers_list": "Fetched installed drivers list.",
                "drivers_backup": "Started drivers backup.",
                "updates_pending": "Fetched pending updates.",
                "drivers_issues": "Fetched problematic drivers.",
            }
            msg = driver_msgs_ar.get(mode) if arabic else driver_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "info_tools":
            mode = str(params.get("mode", "")).lower()
            info_msgs_ar = {
                "windows_product_key": "تم جلب مفتاح تفعيل ويندوز.",
                "model_info": "تم جلب موديل الجهاز والشركة المصنعة.",
                "system_language": "تم جلب لغة النظام الحالية.",
                "timezone_get": "تم جلب المنطقة الزمنية الحالية.",
                "timezone_set": "تم تعديل المنطقة الزمنية.",
                "windows_install_date": "تم جلب تاريخ تثبيت ويندوز.",
                "refresh_rate": "تم جلب معدل تحديث الشاشة.",
            }
            info_msgs_en = {
                "windows_product_key": "Fetched Windows product key.",
                "model_info": "Fetched device model/manufacturer.",
                "system_language": "Fetched current system language.",
                "timezone_get": "Fetched current timezone.",
                "timezone_set": "Updated timezone.",
                "windows_install_date": "Fetched Windows install date.",
                "refresh_rate": "Fetched display refresh rate.",
            }
            msg = info_msgs_ar.get(mode) if arabic else info_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "window_control":
            mode = str(params.get("mode", "")).lower()
            window_msgs_ar = {
                "minimize": "تم تصغير النافذة.",
                "maximize": "تم تكبير النافذة.",
                "restore": "تمت استعادة حجم النافذة.",
                "close_current": "تم إغلاق النافذة الحالية.",
                "show_desktop": "تم تصغير كل النوافذ وإظهار سطح المكتب.",
                "show_desktop_verified": "تم عرض سطح المكتب.",
                "undo_show_desktop": "تمت إعادة إظهار النوافذ المصغرة.",
                "desktop_icons_show": "🖼️ تم إظهار أيقونات سطح المكتب.",
                "desktop_icons_hide": "🖼️ تم إخفاء أيقونات سطح المكتب.",
                "desktop_icons_toggle": "🖼️ تم تبديل حالة أيقونات سطح المكتب.",
                "split_left": "تم نقل النافذة لليسار.",
                "split_right": "تم نقل النافذة لليمين.",
                "task_view": "تم فتح عرض المهام.",
                "alt_tab": "تم تبديل النافذة.",
            }
            window_msgs_en = {
                "minimize": "Window minimized.",
                "maximize": "Window maximized.",
                "restore": "Window restored.",
                "close_current": "Current window closed.",
                "show_desktop": "Minimized all windows (Show Desktop).",
                "show_desktop_verified": "Show Desktop executed.",
                "undo_show_desktop": "Restored minimized windows.",
                "desktop_icons_show": "Desktop icons shown.",
                "desktop_icons_hide": "Desktop icons hidden.",
                "desktop_icons_toggle": "Desktop icons visibility toggled.",
                "split_left": "Moved window to the left side.",
                "split_right": "Moved window to the right side.",
                "task_view": "Opened Task View.",
                "alt_tab": "Switched window.",
            }
            msg = window_msgs_ar.get(mode) if arabic else window_msgs_en.get(mode)
            if msg:
                return True, msg

        if action == "app_tools":
            mode = str(params.get("mode", "")).lower()
            if mode == "close_all_apps" and isinstance(parsed, dict):
                count = int(parsed.get("count") or 0)
                dry_run = bool(parsed.get("dry_run"))
                max_kill = parsed.get("max_kill")
                if arabic:
                    return True, (
                        ("معاينة إغلاق كل البرامج: " if dry_run else "تم تنفيذ إغلاق البرامج: ")
                        + f"{'سيتم' if dry_run else 'تم'} التأثير على {count} تطبيق"
                        + (f" (حد أقصى {max_kill})" if max_kill is not None else "")
                        + "."
                    )
                return True, (
                    ("Preview close-all-apps: " if dry_run else "Executed close-all-apps: ")
                    + f"{'would affect' if dry_run else 'affected'} {count} app(s)"
                    + (f" (max_kill={max_kill})" if max_kill is not None else "")
                    + "."
                )
            app_msgs_ar = {
                "open_task_manager": "تم فتح مدير المهام.",
                "open_notepad": "تم فتح المفكرة.",
                "open_calc": "تم فتح الآلة الحاسبة.",
                "open_paint": "تم فتح الرسام.",
                "open_default_browser": "تم فتح المتصفح الافتراضي.",
                "open_chrome": "تم فتح Chrome.",
                "open_control_panel": "تم فتح لوحة التحكم.",
                "open_add_remove_programs": "تم فتح نافذة إضافة أو إزالة البرامج.",
                "open_store": "تم فتح متجر Microsoft.",
                "open_camera": "تم فتح الكاميرا.",
                "open_calendar": "تم فتح التقويم.",
                "open_mail": "تم فتح البريد.",
                "open_volume_mixer": "تم فتح خالط الصوت.",
                "open_mic_settings": "تم فتح إعدادات الميكروفون.",
                "open_sound_output": "تم فتح إعدادات مخرج الصوت.",
                "open_spatial_sound": "تم فتح إعدادات الصوت المكاني.",
                "open_sound_cpl": "تم فتح إعدادات الصوت الكلاسيكية.",
                "open_network_connections": "تم فتح اتصالات الشبكة.",
                "open_netconnections_cpl": "تم فتح لوحة اتصالات الشبكة.",
                "open_time_date": "تم فتح إعدادات الوقت والتاريخ.",
                "open_system_properties": "تم فتح خصائص النظام.",
                "open_power_options": "تم فتح خيارات الطاقة.",
                "open_firewall_cpl": "تم فتح جدار الحماية.",
                "open_internet_options_cpl": "تم فتح خصائص الإنترنت.",
                "open_display_cpl": "تم فتح إعدادات العرض الكلاسيكية.",
                "open_admin_tools_cpl": "تم فتح الأدوات الإدارية.",
                "open_schedtasks_cpl": "تم فتح لوحة المهام المجدولة.",
                "open_mouse_cpl": "تم فتح خصائص الفأرة.",
                "open_keyboard_cpl": "تم فتح إعدادات لوحة المفاتيح.",
                "open_fonts_cpl": "تم فتح لوحة الخطوط.",
                "open_region_cpl": "تم فتح إعدادات الإقليم.",
                "open_folder_options_cpl": "تم فتح خيارات المجلدات.",
                "open_color_cpl": "تم فتح إعدادات الألوان.",
                "open_desktop_cpl": "تم فتح إعدادات سطح المكتب.",
                "open_printers_cpl": "تم فتح لوحة الطابعات.",
                "open_user_accounts_cpl": "تم فتح إدارة حسابات المستخدمين.",
                "open_bluetooth_cpl": "تم فتح إعدادات البلوتوث الكلاسيكية.",
                "open_accessibility_cpl": "تم فتح خيارات سهولة الوصول.",
            }
            app_msgs_en = {
                "open_task_manager": "Opened Task Manager.",
                "open_notepad": "Opened Notepad.",
                "open_calc": "Opened Calculator.",
                "open_paint": "Opened Paint.",
                "open_default_browser": "Opened default browser.",
                "open_chrome": "Opened Chrome.",
                "open_control_panel": "Opened Control Panel.",
                "open_add_remove_programs": "Opened Programs and Features.",
                "open_store": "Opened Microsoft Store.",
                "open_camera": "Opened Camera.",
                "open_calendar": "Opened Calendar.",
                "open_mail": "Opened Mail.",
                "open_volume_mixer": "Opened Volume Mixer.",
                "open_mic_settings": "Opened Microphone settings.",
                "open_sound_output": "Opened audio output settings.",
                "open_spatial_sound": "Opened spatial sound settings.",
                "open_sound_cpl": "Opened classic Sound settings.",
                "open_network_connections": "Opened Network Connections.",
                "open_netconnections_cpl": "Opened Network Connections control panel.",
                "open_time_date": "Opened Date and Time settings.",
                "open_system_properties": "Opened System Properties.",
                "open_power_options": "Opened Power Options.",
                "open_firewall_cpl": "Opened Windows Firewall.",
                "open_internet_options_cpl": "Opened Internet Options.",
                "open_display_cpl": "Opened classic Display settings.",
                "open_admin_tools_cpl": "Opened Administrative Tools.",
                "open_schedtasks_cpl": "Opened Scheduled Tasks control panel.",
                "open_mouse_cpl": "Opened Mouse properties.",
                "open_keyboard_cpl": "Opened Keyboard settings.",
                "open_fonts_cpl": "Opened Fonts control panel.",
                "open_region_cpl": "Opened Region settings.",
                "open_folder_options_cpl": "Opened Folder Options.",
                "open_color_cpl": "Opened classic Color settings.",
                "open_desktop_cpl": "Opened classic Desktop settings.",
                "open_printers_cpl": "Opened Printers control panel.",
                "open_user_accounts_cpl": "Opened classic User Accounts settings.",
                "open_bluetooth_cpl": "Opened classic Bluetooth settings.",
                "open_accessibility_cpl": "Opened classic Ease of Access settings.",
            }
            if mode:
                msg = app_msgs_ar.get(mode) if arabic else app_msgs_en.get(mode)
                if msg:
                    return True, msg

        if action == "dev_tools":
            mode = str(params.get("mode", "")).lower()
            dev_msgs_ar = {
                "open_services": "تم فتح نافذة الخدمات.",
                "open_task_scheduler": "تم فتح جدول المهام.",
                "open_computer_management": "تم فتح إدارة الكمبيوتر.",
                "open_local_users_groups": "تم فتح المستخدمين والمجموعات المحلية.",
                "open_local_security_policy": "تم فتح سياسة الأمان المحلية.",
                "open_print_management": "تم فتح إدارة الطباعة.",
                "event_errors": "تم جلب أحدث أخطاء سجل الأحداث.",
                "analyze_bsod": "تم تحليل أحداث انهيار النظام (BSOD).",
            }
            dev_msgs_en = {
                "open_services": "Opened Services console.",
                "open_task_scheduler": "Opened Task Scheduler.",
                "open_computer_management": "Opened Computer Management.",
                "open_local_users_groups": "Opened Local Users and Groups.",
                "open_local_security_policy": "Opened Local Security Policy.",
                "open_print_management": "Opened Print Management.",
                "event_errors": "Fetched recent event log errors.",
                "analyze_bsod": "Analyzed recent BSOD/system crash events.",
            }
            if mode:
                msg = dev_msgs_ar.get(mode) if arabic else dev_msgs_en.get(mode)
                if msg:
                    return True, msg

        if action == "shell_tools":
            mode = str(params.get("mode", "")).lower()
            shell_msgs_ar = {
                "quick_settings": "تم فتح الإعدادات السريعة.",
                "notifications": "تم فتح مركز الإشعارات.",
                "search": "تم فتح بحث ويندوز.",
                "run": "تم فتح نافذة Run.",
                "file_explorer": "تم فتح مستكشف الملفات.",
                "quick_link_menu": "تم فتح قائمة الارتباط السريع (Win+X).",
                "task_view": "تم فتح عرض المهام.",
            }
            shell_msgs_en = {
                "quick_settings": "Opened Quick Settings.",
                "notifications": "Opened Notification Center.",
                "search": "Opened Windows Search.",
                "run": "Opened Run dialog.",
                "file_explorer": "Opened File Explorer.",
                "quick_link_menu": "Opened Quick Link menu (Win+X).",
                "task_view": "Opened Task View.",
            }
            if mode:
                msg = shell_msgs_ar.get(mode) if arabic else shell_msgs_en.get(mode)
                if msg:
                    return True, msg

        if isinstance(parsed, dict) and "ok" in parsed and "message" in parsed:
            return True, str(parsed.get("message") or "")

        ack_ar = (
            "✅ تم.",
            "✨ جاهز.",
            "👍 تم التنفيذ.",
            "🚀 خلصت.",
        )
        ack_en = (
            "✅ Done.",
            "✨ All set.",
            "👍 Executed.",
            "🚀 Finished.",
        )
        mode_for_ack = str(params.get("mode", "")).strip().lower()
        basis = f"{session_key}|{action}|{mode_for_ack}"
        idx = (sum(ord(ch) for ch in basis) % 4) if basis else 0
        return True, (ack_ar[idx] if arabic else ack_en[idx])

    def _timeout_message(self, *, backend: str, provider: str) -> str:
        backend_n = str(backend or "").strip().lower()
        provider_n = str(provider or "").strip().lower()
        hints: list[str] = []
        if backend_n == "claude_agent_sdk":
            hints.append("- Claude Code CLI may be missing (`npm install -g @anthropic-ai/claude-code`).")
        if provider_n == "gemini":
            hints.append("- Gemini quota/key may be exhausted or invalid in Settings -> API Keys.")
        if provider_n == "ollama":
            hints.append("- Ensure Ollama is running and the selected model is available locally.")
        if provider_n in {"openai", "openai_compatible", "anthropic"}:
            hints.append("- Check API key and model name in Settings -> API Keys.")
        hints.append("- You can switch backend in Settings -> General.")
        return "Request timed out — backend didn't respond.\n\nPossible causes:\n" + "\n".join(hints)

    async def start(self) -> None:
        """Start the agent loop."""
        self._running = True
        settings = Settings.load()
        logger.info(f"🤖 Agent Loop started (Backend: {settings.agent_backend})")
        await self._loop()

    async def stop(self) -> None:
        """Stop the agent loop."""
        self._running = False
        logger.info("🛑 Agent Loop stopped")

    async def _loop(self) -> None:
        """Main processing loop."""
        while self._running:
            # 1. Consume message from Bus
            message = await self.bus.consume_inbound(timeout=1.0)
            if not message:
                continue

            # 2. Process message in background task (to not block loop)
            task = asyncio.create_task(self._process_message(message))
            self._background_tasks.add(task)
            task.add_done_callback(self._background_tasks.discard)

    async def _process_message(self, message: InboundMessage) -> None:
        """Process a single message flow using AgentRouter."""
        session_key = message.session_key
        logger.info(f"⚡ Processing message from {session_key}")

        # Resolve alias so two chats aliased to the same session serialize correctly
        resolved_key = await self.memory.resolve_session_key(session_key)
        task = asyncio.current_task()
        if task is not None:
            self._session_tasks[resolved_key] = task

        try:
            # Global concurrency limit — blocks until a slot is available
            async with self._global_semaphore:
                # Per-session lock — serializes messages within the same session
                if resolved_key not in self._session_locks:
                    self._session_locks[resolved_key] = asyncio.Lock()
                lock = self._session_locks[resolved_key]
                async with lock:
                    await self._process_message_inner(message, resolved_key)

                # Clean up lock if no one else is waiting on it
                if not lock.locked():
                    self._session_locks.pop(resolved_key, None)
        except asyncio.CancelledError:
            logger.info("⏹️ Cancelled in-flight response for %s", resolved_key)
            raise
        finally:
            if task is not None and self._session_tasks.get(resolved_key) is task:
                self._session_tasks.pop(resolved_key, None)

    _WELCOME_EXCLUDED = frozenset({Channel.WEBSOCKET, Channel.CLI, Channel.SYSTEM, Channel.TELEGRAM})

    async def _process_message_inner(self, message: InboundMessage, session_key: str) -> None:
        """Inner message processing (called under concurrency guards)."""
        # Keep context_builder in sync if memory manager was hot-reloaded
        if self.context_builder.memory is not self.memory:
            self.context_builder.memory = self.memory

        # Command interception — handle /new, /sessions, /resume, /help
        # before any agent processing or memory storage
        cmd_handler = get_command_handler()
        if cmd_handler.is_command(message.content):
            response = await cmd_handler.handle(message)
            if response is not None:
                await self.bus.publish_outbound(response)
                await self.bus.publish_outbound(
                    OutboundMessage(
                        channel=message.channel,
                        chat_id=message.chat_id,
                        content="",
                        is_stream_end=True,
                    )
                )
                return

        # Welcome hint — one-time message on first interaction in a channel
        if self.settings.welcome_hint_enabled and message.channel not in self._WELCOME_EXCLUDED:
            existing = await self.memory.get_session_history(session_key, limit=1)
            if not existing:
                await self.bus.publish_outbound(
                    OutboundMessage(
                        channel=message.channel,
                        chat_id=message.chat_id,
                        content=(
                            f"مرحباً بك في {self.settings.assistant_display_name_ar}! اكتب /help لعرض الأوامر."
                        ),
                    )
                )

        try:
            # 0. Injection scan for non-owner sources
            content = message.content
            if self.settings.injection_scan_enabled:
                scanner = get_injection_scanner()
                source = message.metadata.get("source", message.channel.value)
                scan_result = scanner.scan(content, source=source)

                if scan_result.threat_level == ThreatLevel.HIGH:
                    if self.settings.injection_scan_llm:
                        scan_result = await scanner.deep_scan(content, source=source)

                    if scan_result.threat_level == ThreatLevel.HIGH:
                        logger.warning(
                            "Blocked HIGH threat injection from %s: %s",
                            source,
                            scan_result.matched_patterns,
                        )
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="error",
                                data={
                                    "message": "Message blocked by injection scanner",
                                    "patterns": scan_result.matched_patterns,
                                },
                            )
                        )
                        await self.bus.publish_outbound(
                            OutboundMessage(
                                channel=message.channel,
                                chat_id=message.chat_id,
                                content=(
                                    "Your message was flagged by the security scanner and blocked."
                                ),
                            )
                        )
                        return

                # Wrap suspicious (non-blocked) content with sanitization markers
                if scan_result.threat_level != ThreatLevel.NONE:
                    content = scan_result.sanitized_content

            # Deterministic Windows desktop fast-path (independent from backend/provider)
            handled_fastpath, fastpath_reply = await self._try_global_windows_fastpath(
                text=content, session_key=session_key
            )
            if handled_fastpath:
                await self.memory.add_to_session(
                    session_key=session_key,
                    role="user",
                    content=content,
                    metadata=message.metadata,
                )
                reply_text = str(fastpath_reply or "").strip() or "Done."
                await self.bus.publish_outbound(
                    OutboundMessage(
                        channel=message.channel,
                        chat_id=message.chat_id,
                        content=reply_text,
                        is_stream_chunk=True,
                    )
                )
                await self.bus.publish_outbound(
                    OutboundMessage(
                        channel=message.channel,
                        chat_id=message.chat_id,
                        content="",
                        is_stream_end=True,
                    )
                )
                await self.memory.add_to_session(
                    session_key=session_key, role="assistant", content=reply_text
                )
                return

            # 1. Store User Message
            await self.memory.add_to_session(
                session_key=session_key,
                role="user",
                content=content,
                metadata=message.metadata,
            )

            # 2. Build dynamic system prompt (identity + memory context + channel hint)
            sender_id = message.sender_id
            system_prompt = await self.context_builder.build_system_prompt(
                user_query=content,
                channel=message.channel,
                sender_id=sender_id,
                session_key=message.session_key,
            )

            # 2a. Retrieve session history with compaction
            history = await self.memory.get_compacted_history(
                session_key,
                recent_window=self.settings.compaction_recent_window,
                char_budget=self.settings.compaction_char_budget,
                summary_chars=self.settings.compaction_summary_chars,
                llm_summarize=self.settings.compaction_llm_summarize,
            )

            # 2b. Emit thinking event
            await self.bus.publish_system(
                SystemEvent(event_type="thinking", data={"session_key": session_key})
            )

            # 3. Run through AgentRouter (handles all backends)
            router = self._get_router()
            full_response = ""
            composition_events: list[dict] = []
            use_ai_composer = bool(
                getattr(self.settings, "ai_response_composer_enabled", True)
                and self.settings.agent_backend == "open_interpreter"
            )

            run_iter = router.run(content, system_prompt=system_prompt, history=history)
            try:
                async for chunk in _iter_with_timeout(
                    run_iter,
                    first_timeout=FIRST_RESPONSE_TIMEOUT_SECONDS,
                    timeout=STREAM_CHUNK_TIMEOUT_SECONDS,
                ):
                    chunk_type = chunk.get("type", "")
                    content = chunk.get("content", "")
                    metadata = chunk.get("metadata") or {}

                    if chunk_type == "message":
                        content = self._sanitize_stream_chunk(content)
                        if not content:
                            continue
                        composition_events.append({"type": "message", "content": content})
                        full_response += content
                        if not use_ai_composer:
                            # Stream text to user
                            await self.bus.publish_outbound(
                                OutboundMessage(
                                    channel=message.channel,
                                    chat_id=message.chat_id,
                                    content=content,
                                    is_stream_chunk=True,
                                )
                            )

                    elif chunk_type == "code":
                        # Code block from Open Interpreter - emit as tool_use
                        language = metadata.get("language", "code")
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="tool_start",
                                data={
                                    "name": f"run_{language}",
                                    "params": {"code": content[:100]},
                                },
                            )
                        )
                        # Also stream to user
                        code_block = f"\n```{language}\n{content}\n```\n"
                        composition_events.append({"type": "code", "language": language})
                        full_response += code_block
                        if not use_ai_composer:
                            await self.bus.publish_outbound(
                                OutboundMessage(
                                    channel=message.channel,
                                    chat_id=message.chat_id,
                                    content=code_block,
                                    is_stream_chunk=True,
                                )
                            )

                    elif chunk_type == "output":
                        # Output from code execution - emit as tool_result
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="tool_result",
                                data={
                                    "name": "code_execution",
                                    "result": content[:200],
                                    "status": "success",
                                },
                            )
                        )
                        # Also stream to user
                        output_block = f"\n```output\n{content}\n```\n"
                        composition_events.append({"type": "output", "content": content[:400]})
                        full_response += output_block
                        if not use_ai_composer:
                            await self.bus.publish_outbound(
                                OutboundMessage(
                                    channel=message.channel,
                                    chat_id=message.chat_id,
                                    content=output_block,
                                    is_stream_chunk=True,
                                )
                            )

                    elif chunk_type == "thinking":
                        # Thinking goes to Activity panel only
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="thinking",
                                data={"content": content, "session_key": session_key},
                            )
                        )

                    elif chunk_type == "thinking_done":
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="thinking_done",
                                data={"session_key": session_key},
                            )
                        )

                    elif chunk_type == "tool_use":
                        # Emit tool_start system event for Activity panel
                        tool_name = metadata.get("name") or metadata.get("tool", "unknown")
                        tool_input = metadata.get("input") or metadata
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="tool_start",
                                data={"name": tool_name, "params": tool_input},
                            )
                        )

                    elif chunk_type == "tool_result":
                        # Emit tool_result system event for Activity panel
                        tool_name = metadata.get("name") or metadata.get("tool", "unknown")
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="tool_result",
                                data={
                                    "name": tool_name,
                                    "result": content[:200],
                                    "status": "success",
                                },
                            )
                        )
                        composition_events.append(
                            {
                                "type": "tool_result",
                                "tool": metadata.get("name") or metadata.get("tool", "unknown"),
                                "content": str(content)[:400],
                            }
                        )

                    elif chunk_type == "result":
                        composition_events.append(
                            {
                                "type": "result",
                                "content": str(content or ""),
                                "metadata": metadata,
                            }
                        )
                        rendered = str(content or "")
                        full_response += rendered
                        if not use_ai_composer and rendered:
                            await self.bus.publish_outbound(
                                OutboundMessage(
                                    channel=message.channel,
                                    chat_id=message.chat_id,
                                    content=rendered,
                                    is_stream_chunk=True,
                                )
                            )

                    elif chunk_type == "error":
                        # Emit error and send to user
                        await self.bus.publish_system(
                            SystemEvent(
                                event_type="tool_result",
                                data={
                                    "name": "agent",
                                    "result": content,
                                    "status": "error",
                                },
                            )
                        )
                        composition_events.append({"type": "error", "content": content})
                        await self.bus.publish_outbound(
                            OutboundMessage(
                                channel=message.channel,
                                chat_id=message.chat_id,
                                content=content,
                                is_stream_chunk=True,
                            )
                        )

                    elif chunk_type == "done":
                        # Agent finished - will send stream_end below
                        pass

                    elif chunk_type == "status":
                        # Internal backend heartbeat/status chunk.
                        logger.debug(
                            "Internal status chunk ignored (session=%s content=%s)",
                            session_key,
                            content,
                        )
                        continue
            finally:
                # Always close the async generator to kill any subprocess
                await run_iter.aclose()

            # 4. Send stream end marker
            if use_ai_composer and full_response.strip():
                composed = await self._compose_response(
                    user_query=message.content,
                    events=composition_events,
                    fallback_text=full_response,
                )
                full_response = composed
                await self.bus.publish_outbound(
                    OutboundMessage(
                        channel=message.channel,
                        chat_id=message.chat_id,
                        content=composed,
                        is_stream_chunk=True,
                    )
                )

            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content="",
                    is_stream_end=True,
                )
            )

            # 5. Store assistant response in memory
            if full_response:
                await self.memory.add_to_session(
                    session_key=session_key, role="assistant", content=full_response
                )

                # 6. Auto-learn: extract facts from conversation (non-blocking)
                should_auto_learn = (
                    self.settings.memory_backend == "mem0" and self.settings.mem0_auto_learn
                ) or (self.settings.memory_backend == "file" and self.settings.file_auto_learn)
                if should_auto_learn:
                    task = asyncio.create_task(
                        self._auto_learn(
                            message.content,
                            full_response,
                            session_key,
                            sender_id=sender_id,
                        )
                    )
                    self._background_tasks.add(task)
                    task.add_done_callback(self._background_tasks.discard)

        except StreamTimeoutError as e:
            llm_provider = str(getattr(self.settings, "llm_provider", "auto") or "auto")
            llm_model = "unknown"
            try:
                from Mudabbir.llm.client import resolve_llm_client

                llm = resolve_llm_client(self.settings)
                llm_provider = llm.provider
                llm_model = llm.model
            except Exception:
                provider_model_map = {
                    "gemini": str(getattr(self.settings, "gemini_model", "") or "unknown"),
                    "openai": str(getattr(self.settings, "openai_model", "") or "unknown"),
                    "anthropic": str(getattr(self.settings, "anthropic_model", "") or "unknown"),
                    "ollama": str(getattr(self.settings, "ollama_model", "") or "unknown"),
                    "openai_compatible": str(
                        getattr(self.settings, "openai_compatible_model", "") or "unknown"
                    ),
                }
                llm_model = provider_model_map.get(llm_provider, "unknown")

            logger.error(
                "Agent backend timed out (session=%s backend=%s provider=%s model=%s phase=%s timeout_seconds=%s first_timeout=%ss chunk_timeout=%ss)",
                session_key,
                self.settings.agent_backend,
                llm_provider,
                llm_model,
                e.phase,
                e.timeout_seconds,
                FIRST_RESPONSE_TIMEOUT_SECONDS,
                STREAM_CHUNK_TIMEOUT_SECONDS,
            )
            # Kill the hung backend so it releases resources
            try:
                active_router = self._router
                if active_router is not None:
                    await active_router.stop()
            except Exception:
                pass
            # Force router re-init on next message
            self._router = None

            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content=self._timeout_message(
                        backend=self.settings.agent_backend,
                        provider=llm_provider,
                    ),
                    is_stream_chunk=True,
                )
            )
            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content="",
                    is_stream_end=True,
                )
            )
        except TimeoutError:
            llm_provider = str(getattr(self.settings, "llm_provider", "auto") or "auto")
            llm_model = "unknown"
            try:
                from Mudabbir.llm.client import resolve_llm_client

                llm = resolve_llm_client(self.settings)
                llm_provider = llm.provider
                llm_model = llm.model
            except Exception:
                provider_model_map = {
                    "gemini": str(getattr(self.settings, "gemini_model", "") or "unknown"),
                    "openai": str(getattr(self.settings, "openai_model", "") or "unknown"),
                    "anthropic": str(getattr(self.settings, "anthropic_model", "") or "unknown"),
                    "ollama": str(getattr(self.settings, "ollama_model", "") or "unknown"),
                    "openai_compatible": str(
                        getattr(self.settings, "openai_compatible_model", "") or "unknown"
                    ),
                }
                llm_model = provider_model_map.get(llm_provider, "unknown")

            logger.error(
                "Agent backend timed out (session=%s backend=%s provider=%s model=%s phase=unknown timeout_seconds=unknown first_timeout=%ss chunk_timeout=%ss)",
                session_key,
                self.settings.agent_backend,
                llm_provider,
                llm_model,
                FIRST_RESPONSE_TIMEOUT_SECONDS,
                STREAM_CHUNK_TIMEOUT_SECONDS,
            )
            try:
                active_router = self._router
                if active_router is not None:
                    await active_router.stop()
            except Exception:
                pass
            self._router = None
            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content=self._timeout_message(
                        backend=self.settings.agent_backend,
                        provider=llm_provider,
                    ),
                    is_stream_chunk=True,
                )
            )
            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content="",
                    is_stream_end=True,
                )
            )
        except asyncio.CancelledError:
            raise
        except Exception as e:
            logger.exception(f"❌ Error processing message: {e}")
            # Kill the backend on error
            try:
                active_router = self._router
                if active_router is not None:
                    await active_router.stop()
            except Exception:
                pass

            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content=f"An error occurred: {str(e)}",
                )
            )
            await self.bus.publish_outbound(
                OutboundMessage(
                    channel=message.channel,
                    chat_id=message.chat_id,
                    content="",
                    is_stream_end=True,
                )
            )

    async def _send_response(self, original: InboundMessage, content: str) -> None:
        """Helper to send a simple text response."""
        await self.bus.publish_outbound(
            OutboundMessage(channel=original.channel, chat_id=original.chat_id, content=content)
        )

    async def _auto_learn(
        self,
        user_msg: str,
        assistant_msg: str,
        session_key: str,
        sender_id: str | None = None,
    ) -> None:
        """Background task: feed conversation turn for fact extraction."""
        try:
            messages = [
                {"role": "user", "content": user_msg},
                {"role": "assistant", "content": assistant_msg},
            ]
            result = await self.memory.auto_learn(
                messages,
                file_auto_learn=self.settings.file_auto_learn,
                sender_id=sender_id,
            )
            extracted = len(result.get("results", []))
            if extracted:
                logger.debug("Auto-learned %d facts from %s", extracted, session_key)
        except Exception:
            logger.debug("Auto-learn background task failed", exc_info=True)

    async def cancel_session(self, session_key: str) -> bool:
        """Cancel the current in-flight message task for a session key."""
        resolved_key = await self.memory.resolve_session_key(session_key)
        task = self._session_tasks.get(resolved_key)
        if task is None or task.done():
            return False
        task.cancel()
        try:
            await asyncio.wait_for(asyncio.shield(task), timeout=2.0)
        except asyncio.TimeoutError:
            logger.warning("Timed out waiting for cancelled task to finish: %s", resolved_key)
        except asyncio.CancelledError:
            pass
        except Exception:
            logger.debug("Cancelled task finished with non-cancel exception", exc_info=True)
        return True

    def reset_router(self) -> None:
        """Reset the router to pick up new settings."""
        self._router = None

