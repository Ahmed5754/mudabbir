"""Open Interpreter agent wrapper.

Changes:
  2026-02-05 - Emit tool_use/tool_result events for Activity panel
  2026-02-04 - Filter out verbose console output, only show messages and final results
  2026-02-02 - Added executor layer logging for architecture visibility.
"""

import asyncio
import ast
import importlib
import json
import logging
import os
import re
import subprocess
import time
import warnings
from collections.abc import AsyncIterator
from pathlib import Path
from typing import Any

from Mudabbir.config import Settings
from Mudabbir.tools.capabilities.windows_intent_map import (
    RULES as WINDOWS_INTENT_RULES,
    is_confirmation_message,
    resolve_windows_intent,
)

logger = logging.getLogger(__name__)

DEFAULT_OI_MAX_CODE_BLOCKS = 3
DEFAULT_OI_REPEAT_COMMAND_LIMIT = 1
DEFAULT_OI_MAX_FIX_RETRIES = 1
OI_QUEUE_WAIT_CAP_SECONDS = 600
OI_QUEUE_ACQUIRE_POLL_SECONDS = 5
OI_QUEUE_HEARTBEAT_SECONDS = 20
DONE_PATTERNS = (
    "task is done",
    "done",
    "completed",
    "finished",
    "top 5 process",
    "top five process",
    "تم التنفيذ",
    "اكتمل التنفيذ",
    "انتهيت من التنفيذ",
)
QUOTA_OR_RATE_PATTERNS = (
    "you ran out of current quota",
    "insufficient_quota",
    "rate limit",
    "rate_limit",
    "quota exceeded",
    "quota_exceeded",
    "too many requests",
    "429",
    "resource_exhausted",
)
EXTRA_GUARDRAILS = """Execution guardrails:
- Target OS is Windows. Use PowerShell for system commands unless user explicitly requests Python.
- For desktop automation (mouse/keyboard/UI control), use Python with pyautogui.
- Use exactly one execution language per response and do not switch languages mid-task.
- Never output Markdown code fences/backticks in executable commands.
- Never emit incomplete code. Ensure brackets, quotes, and command syntax are complete before running.
- If execution fails, do one fix-and-retry only. If it still fails, stop and ask the user for output/screenshot.
- Avoid repetitive actions. Do not run the same command repeatedly unless the user explicitly asks.
- Do not repeatedly open GUI apps (for example Task Manager) in loops.
- If the user writes in Arabic, answer in clear Arabic.
- Never expose raw execute JSON payloads in user-facing responses.
- Never claim an action succeeded unless execution output/tool result confirms it.
- If uncertain or blocked, say that clearly instead of guessing."""

ERROR_PATTERNS = (
    "parsererror",
    "commandnotfoundexception",
    "the term",
    "is not recognized",
    "missing closing",
    "unexpected token",
    "write-error",
    "at line:",
)

TASK_MANAGER_KEYWORDS = ("task manager", "taskmgr", "مدير المهام")
TOP_PROCESS_KEYWORDS = (
    "top 5",
    "top five",
    "memory",
    "ram",
    "workingset",
    "cpu",
    "أكثر",
    "استهلاك",
    "الذاكرة",
)

GUI_AUTOMATION_GUARDRAILS = """
Task-specific execution mode:
- This request needs GUI interaction. Use Python with pyautogui only.
- Open GUI apps at most once.
- Do not emit Markdown fences/backticks in executable content.
- After one successful execution, stop and report the result clearly.
"""

TASK_MANAGER_PATTERN = re.compile(
    r"(task\s*manager|taskmgr|start-process\s+taskmgr|تاسك\s*مانجر|مدير\s*المهام)",
    re.IGNORECASE,
)
ARABIC_CHAR_PATTERN = re.compile(r"[\u0600-\u06ff]")
INT_PATTERN = re.compile(r"-?\d+")
ARABIC_DIACRITICS_PATTERN = re.compile(r"[\u0610-\u061a\u064b-\u065f\u0670\u06d6-\u06ed]")


def _normalize_command(command: str) -> str:
    """Normalize command text for repetition detection."""
    normalized = command.replace("`", "")
    normalized = normalized.replace("```", " ")
    normalized = re.sub(r"\s+", " ", normalized.strip()).lower()
    return normalized


def _contains_arabic(text: str) -> bool:
    """Return True when text contains Arabic characters."""
    if not text:
        return False
    return bool(ARABIC_CHAR_PATTERN.search(text))


def _normalize_text_for_match(text: str) -> str:
    """Normalize Arabic/English text for robust keyword matching."""
    if not text:
        return ""
    normalized = ARABIC_DIACRITICS_PATTERN.sub("", text)
    normalized = normalized.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    normalized = normalized.replace("ى", "ي").replace("ة", "ه")
    normalized = normalized.replace("ؤ", "و").replace("ئ", "ي")
    normalized = normalized.lower()
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return normalized


def _contains_done_signal(text: str) -> bool:
    """Return True if the assistant text indicates the task is complete."""
    lowered = text.lower()
    return any(pattern in lowered for pattern in DONE_PATTERNS)


def _is_google_quota_signal(text: str) -> bool:
    """Detect Google/Gemini quota or rate-limit signals."""
    lowered = str(text or "").lower()
    if not lowered.strip():
        return False
    google_markers = (
        "resource_exhausted",
        "generativelanguage.googleapis.com",
        "google.api_core",
        "googleapis.com",
        "gemini",
        "google",
    )
    return (
        any(marker in lowered for marker in google_markers)
        and any(marker in lowered for marker in QUOTA_OR_RATE_PATTERNS)
    ) or "resource_exhausted" in lowered


def _is_quota_or_rate_limit_message(text: str) -> bool:
    """Return True if text indicates API quota/rate-limit exhaustion."""
    lowered = str(text or "").lower()
    if not lowered.strip():
        return False
    if _is_google_quota_signal(lowered):
        return True
    return any(marker in lowered for marker in QUOTA_OR_RATE_PATTERNS)


def _should_retry_with_gemini_fallback(
    provider: str, fallback_attempted: bool, error_text: str
) -> bool:
    """Retry once for Gemini requests that failed due to quota/rate limits."""
    return (
        provider == "gemini"
        and not fallback_attempted
        and _is_quota_or_rate_limit_message(error_text)
    )


def _build_quota_error_message(
    provider: str,
    model: str,
    fallback_model: str,
    *,
    arabic: bool,
    fallback_attempted: bool,
) -> str:
    """Build a provider-aware quota/rate-limit error message."""
    if provider == "gemini":
        if arabic:
            if fallback_attempted:
                return (
                    "❌ تم استهلاك حصة Gemini أو تم تجاوز حد المعدل.\n\n"
                    f"حاولت إعادة الطلب تلقائياً باستخدام `{fallback_model}` ولم تنجح.\n"
                    "تحقق من حصة Google AI Studio/API Key ثم أعد المحاولة لاحقاً."
                )
            return (
                "❌ تم الوصول إلى حد Gemini (quota/rate limit).\n\n"
                f"النموذج الحالي: `{model}`\n"
                "تحقق من حصة Google AI Studio أو جرّب لاحقاً."
            )
        if fallback_attempted:
            return (
                "❌ Gemini quota/rate limit reached.\n\n"
                f"Automatic fallback to `{fallback_model}` was attempted and also failed.\n"
                "Check your Google AI Studio quota/API key and retry later."
            )
        return (
            "❌ Gemini quota/rate limit reached.\n\n"
            f"Current model: `{model}`\n"
            "Check your Google AI Studio quota/API key and retry later."
        )

    if arabic:
        return (
            "❌ تم الوصول إلى حد المزود (quota/rate limit).\n\n"
            "يرجى التحقق من حدود الحساب ثم إعادة المحاولة."
        )
    return (
        "❌ Provider quota/rate limit reached.\n\n"
        "Please check your account limits and retry."
    )


def _extract_command_fingerprints(command: str, *, allow_fallback: bool = True) -> set[str]:
    """Extract stable command fingerprints from a code block."""
    fingerprints: set[str] = set()
    command = command.replace("```powershell", "").replace("```", "")

    for raw_line in command.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        lowered = line.lower()

        # Skip wrappers/noise that should not count as meaningful commands.
        if lowered.startswith("#"):
            continue
        if lowered.startswith("write-output"):
            continue
        if lowered.startswith("write-error"):
            continue
        if lowered.startswith("$erroractionpreference"):
            continue
        if lowered.startswith("$"):
            continue
        if lowered in {"try", "try {", "catch", "catch {", "}", "{", "} catch {"}:
            continue

        raw_lowered = lowered.replace("`", "")
        lowered = raw_lowered.replace('"', "").replace("'", "")
        lowered = re.sub(r"\s+", " ", lowered).strip()
        if not lowered:
            continue

        looks_like_command = (
            "|" in lowered
            or lowered.startswith("start-process")
            or lowered.startswith("get-process")
            or lowered.startswith("tasklist")
            or lowered.startswith("explorer")
            or lowered.startswith("python ")
            or lowered.startswith("py ")
            or lowered.startswith("import ")
            or lowered.startswith("powershell")
            or lowered.startswith("pwsh")
            or lowered.startswith("cmd ")
            or lowered.startswith("try ")
            or lowered.startswith("select-object")
        )
        if not looks_like_command:
            continue

        start_process = re.search(
            r"start-process(?:\s+-filepath)?\s+(?:\"([^\"]+)\"|'([^']+)'|([^\s|;]+))",
            raw_lowered,
        )
        if start_process:
            target = (start_process.group(1) or start_process.group(2) or start_process.group(3) or "").strip()
            target = target.strip("\"'`")
            if target.startswith("task"):
                target = "taskmgr.exe"
            if "\\" in target or target.endswith(".exe"):
                target = target.lower()
            fingerprints.add(f"start-process:{target}")
            continue

        if lowered.startswith("get-process"):
            if "workingset" in lowered:
                fingerprints.add("get-process:workingset")
            else:
                fingerprints.add("get-process")
            continue

        if lowered.startswith("tasklist"):
            fingerprints.add("tasklist")
            continue

        if lowered.startswith("explorer"):
            fingerprints.add("explorer")
            continue

        tokens = lowered.split()
        if tokens:
            fingerprints.add(" ".join(tokens[:2]))

    if allow_fallback and not fingerprints:
        fingerprints.add(_normalize_command(command)[:200])

    return fingerprints


def _looks_like_error_output(output: str) -> bool:
    """Detect obvious execution errors in console output."""
    lowered = output.lower()
    return any(pattern in lowered for pattern in ERROR_PATTERNS)


def _looks_like_process_snapshot(output: str) -> bool:
    """Detect process table-like output."""
    lowered = output.lower()
    return "processname" in lowered and ("workingset" in lowered or "id" in lowered)


def _contains_any(text: str, keywords: tuple[str, ...]) -> bool:
    lowered = text.lower()
    return any(k in lowered for k in keywords)


def _build_windows_alias_fallback() -> tuple[str, ...]:
    """Build fallback phrases from intent-map aliases instead of rigid hand-picked keywords."""
    phrases: set[str] = set()
    for rule in WINDOWS_INTENT_RULES:
        for alias in tuple(getattr(rule, "aliases", ()) or ()):
            norm = _normalize_text_for_match(str(alias or ""))
            if not norm:
                continue
            # Skip tiny fragments to reduce false positives in normal chat.
            if len(norm) < 4:
                continue
            phrases.add(norm)
    return tuple(sorted(phrases, key=len, reverse=True))


WINDOWS_ALIAS_FALLBACK = _build_windows_alias_fallback()


def _looks_like_raw_command_leak(text: str) -> bool:
    """Detect leaked/fragmented command text that should never be user-facing."""
    if not text:
        return False
    compact = str(text).strip().strip("`")
    if not compact:
        return False
    lowered = compact.lower()

    command_markers = (
        "start-process",
        "get-process",
        "stop-process",
        "set-volume",
        "set-volumelevel",
        "set-culture",
        "set-win",
        "set-bluetoothstate",
        "write-output",
        "shell:appsfolder",
        "pyautogui",
        "import pyautogui",
        "-volume",
    )
    if any(m in lowered for m in command_markers) and (
        '{"name"' in lowered
        or '"arguments"' in lowered
        or '"language"' in lowered
        or '"code"' in lowered
        or lowered.endswith('"}}')
        or lowered.endswith("'}")
    ):
        return True

    # Truncated "powershell" chunks observed in streaming, e.g. "powers -Volume 100"}}.
    if re.search(r'^\s*(?:powers|powershell)\b', lowered) and any(
        t in lowered for t in ("-volume", "start-process", "get-process", "set-", "pyautogui")
    ):
        return True

    # Raw python snippets leaking as text instead of tool execution summaries.
    if (lowered.startswith("import pyautogui") or lowered.startswith("\\nimport pyautogui")) and (
        "\\n" in compact or "pyautogui." in lowered
    ):
        return True
    if "pyautogui.screenshot" in lowered and "screenshot.save" in lowered:
        return True
    if re.search(r"^\s*name\s+pyautogui", lowered) and "screenshot" in lowered:
        return True

    return False


def _is_execute_fragment(text: str) -> bool:
    """Return True for partial JSON snippets likely belonging to execute payload."""
    if not text:
        return False
    if _looks_like_raw_command_leak(text):
        return True
    lowered = text.lower()
    markers = (
        "execute",
        '"arguments"',
        '"language"',
        '"code"',
        "start-process",
        "get-process",
        "stop-process",
        "set-volume",
        "set-volumelevel",
        "set-culture",
        "set-win",
        "set-bluetoothstate",
        "shell:appsfolder",
        "write-output",
        "powershell",
        "powers ",
        "python",
        "pyautogui",
    )
    has_marker = any(m in lowered for m in markers)
    if has_marker and any(ch in text for ch in "{}:"):
        return True

    # Common fragmented chunks:
    #   ": "powershell",
    #   "code": "Start-Process ..."
    if re.search(r'^\s*"?\s*:\s*"(powershell|python|pwsh)"', lowered):
        return True
    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*:', lowered):
        return True
    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*,\s*"\s*:\s*', lowered):
        return True
    if '"language"' in lowered and '"code"' in lowered:
        return True
    return False


def _looks_like_execute_noise(text: str) -> bool:
    """Detect malformed JSON-like execute noise that should not be shown to users."""
    if not text:
        return False
    if _looks_like_raw_command_leak(text):
        return True
    lowered = text.lower().strip().strip("`")
    if not lowered:
        return False

    looks_jsonish = (
        lowered.startswith("{")
        or lowered.startswith("[")
        or lowered.startswith('{"')
        or '"arguments' in lowered
        or re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*:', lowered)
    )
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
        "shell:appsfolder",
        "write-output",
        "explorer.exe",
        "pyautogui",
        "powershell",
        "powers ",
        "python",
    )
    if looks_jsonish and any(m in lowered for m in markers):
        return True

    if re.search(r'^\s*"?\s*:\s*"(powershell|python|pwsh)"', lowered):
        return True
    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*,\s*"\s*:\s*', lowered):
        return True
    if '"language"' in lowered and '"code"' in lowered:
        return True
    return False


def _looks_like_execute_payload_fragment(text: str) -> bool:
    """Detect partial execute payload fragments (including broken chunks)."""
    if not text:
        return False
    if _looks_like_raw_command_leak(text):
        return True
    stripped = text.strip()
    if not stripped:
        return False
    compact = stripped.strip("`").strip()

    lowered = compact.lower()
    if _is_execute_fragment(compact):
        return True

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
        "shell:appsfolder",
        "write-output",
        "explorer.exe",
        "pyautogui",
        "powershell",
        "powers ",
        "python",
    )

    if compact in {"{", "}", "[", "]", "```", "`"}:
        return True

    if compact.startswith("{") and any(m in lowered for m in markers):
        return True

    # Broken/mangled payloads observed in streamed chunks.
    broken = ("namepowers", "argumentspowers", "namepython", "argumentspython")
    if any(b in lowered for b in broken) and any(m in lowered for m in markers):
        return True

    # Very short JSON-ish prefix chunks often precede malformed execute payloads.
    if compact.startswith("{") and len(compact) <= 24:
        return True

    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*:', lowered):
        return True
    if re.search(r'^\s*"?\s*:\s*"(powershell|python|pwsh)"', lowered):
        return True
    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*,\s*"\s*:\s*', lowered):
        return True
    if ('"language"' in lowered and '"code"' in lowered) or (
        "language" in lowered and "code" in lowered and any(m in lowered for m in markers)
    ):
        return True

    return False


def is_noisy_execution_text(text: str) -> bool:
    """Shared adapter-safe guard: true when text should be hidden from user output."""
    if not text:
        return False
    return (
        _looks_like_execute_noise(text)
        or _looks_like_execute_payload_fragment(text)
        or _looks_like_raw_command_leak(text)
    )


def _extract_first_error_line(output: str) -> str:
    """Return the first meaningful error line from command output."""
    for line in (output or "").splitlines():
        cleaned = line.strip()
        if not cleaned:
            continue
        if cleaned.lower().startswith("at line:"):
            continue
        return cleaned
    return "Unknown error."


def _extract_first_int(text: str) -> int | None:
    """Extract first integer value from text."""
    if not text:
        return None
    m = INT_PATTERN.search(text)
    if not m:
        return None
    try:
        return int(m.group(0))
    except Exception:
        return None


def _extract_python_code_fallback(text: str) -> str | None:
    """Best-effort extraction of desktop Python code from malformed execute payload text."""
    if not text:
        return None
    cleaned = text.replace("\r", "\n")
    lowered = cleaned.lower()
    if "pyautogui" not in lowered and "time.sleep" not in lowered:
        return None

    candidate = cleaned
    code_field = re.search(r'"code"\s*:\s*"([\s\S]+?)"\s*}\s*}?', cleaned, re.IGNORECASE)
    if code_field:
        candidate = code_field.group(1)
    candidate = candidate.replace('\\"', '"').replace("\\n", "\n")

    lines: list[str] = []
    for raw in candidate.splitlines():
        line = raw.strip().strip('"').strip("'").rstrip(",}").strip()
        if not line:
            continue
        lower_line = line.lower()
        if lower_line.startswith("import "):
            modules = [m.strip() for m in line[7:].split(",") if m.strip()]
            if modules and all(m.lower() in {"pyautogui", "time"} for m in modules):
                lines.append(line)
            continue
        if lower_line.startswith("pyautogui.") or lower_line.startswith("time.sleep("):
            lines.append(line)

    if not lines:
        return None

    has_py_cmd = any(ln.lower().startswith("pyautogui.") for ln in lines)
    has_py_import = any(ln.lower().startswith("import pyautogui") for ln in lines)
    if has_py_cmd and not has_py_import:
        lines.insert(0, "import pyautogui")
    return "\n".join(lines)


def _summarize_structured_result(code: str, output: str, *, arabic: bool) -> str:
    """Create a short, user-facing summary for structured command execution."""
    lowered = (code or "").lower()
    has_error = _looks_like_error_output(output)
    no_output = (output or "").strip() in {"", "(no output)"}

    if has_error:
        err_line = _extract_first_error_line(output)
        if arabic:
            return f"فشل تنفيذ الأمر.\nالسبب: {err_line}"
        return f"Command failed.\nReason: {err_line}"

    if "measure-object" in lowered and "mainwindowtitle" in lowered:
        count = _extract_first_int(output)
        if count is not None:
            if arabic:
                return f"عدد النوافذ المفتوحة حاليًا: {count}"
            return f"Open windows count: {count}"

    if "start-process" in lowered:
        target_match = re.search(
            r"start-process(?:\s+-filepath)?\s+(?:\"([^\"]+)\"|'([^']+)'|([^\s|;]+))",
            code,
            re.IGNORECASE,
        )
        target = (target_match.group(1) or target_match.group(2) or target_match.group(3) or "").strip() if target_match else ""
        target = target.strip("\"'`")
        app_name = Path(target).name if target else "application"
        if arabic:
            if target:
                return f"تم تشغيل التطبيق بنجاح: {app_name}\nالمسار: {target}"
            return "تم تنفيذ أمر تشغيل التطبيق بنجاح."
        if target:
            return f"Application launched: {app_name}\nPath: {target}"
        return "Launch command executed successfully."

    if "explorer.exe" in lowered or lowered.startswith("explorer "):
        if arabic:
            return "تم فتح مستكشف الملفات بنجاح."
        return "File Explorer opened successfully."

    if "imagegrab.grab" in lowered or (
        "screenshot" in lowered and ("pyautogui" in lowered or "imagegrab" in lowered)
    ):
        if arabic:
            return "تم تنفيذ طلب التقاط الشاشة بنجاح."
        return "Screen capture request executed successfully."

    if "pyautogui." in lowered:
        if arabic:
            return "تم تنفيذ أمر التحكم بسطح المكتب بنجاح."
        return "Desktop control command executed successfully."

    if no_output:
        if arabic:
            return "تم تنفيذ الأمر بنجاح."
        return "Command executed successfully."

    if arabic:
        return f"تم تنفيذ الأمر.\nالناتج:\n{output[:1200]}"
    return f"Command executed.\nOutput:\n{output[:1200]}"


def _looks_like_execute_continuation(text: str) -> bool:
    """Return True for chunks that likely continue a partial execute payload."""
    if not text:
        return False
    if _looks_like_raw_command_leak(text):
        return True
    lowered = text.lower()
    markers = (
        ":\\",
        ".exe",
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
        "https://",
        "http://",
        "powers ",
        "python",
        "pyautogui",
        "tasklist",
        "explorer.exe",
        '"language"',
        '"code"',
        '"arguments"',
        '"name"',
        "{",
        "}",
        "[",
        "]",
    )
    if any(m in lowered for m in markers):
        return True
    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*:', lowered):
        return True
    if re.search(r'^\s*"?\s*(name|arguments|language|code)\s*"?\s*,\s*"\s*:\s*', lowered):
        return True
    if re.search(r'^\s*"?\s*:\s*"(powershell|python|pwsh)"', lowered):
        return True
    return False


def _extract_execute_payload(text: str) -> tuple[str, str] | None:
    """Extract a JSON-ish execute payload from assistant text."""
    if not text:
        return None
    lowered = text.lower()
    if "code" not in lowered and not _looks_like_execute_payload_fragment(text):
        return None

    start = text.find("{")
    end = text.rfind("}")
    candidate = text[start : end + 1].strip() if (start >= 0 and end > start) else text.strip()

    def _extract_from_payload(payload: object) -> tuple[str, str] | None:
        if not isinstance(payload, dict):
            return None
        name = str(payload.get("name", "")).strip().lower()
        args = payload.get("arguments")
        if isinstance(args, dict):
            if name and name != "execute":
                return None
        else:
            args = payload
        if not isinstance(args, dict):
            return None
        language = str(args.get("language", "powershell")).strip().lower() or "powershell"
        code = args.get("code")
        if not isinstance(code, str) or not code.strip():
            return None
        return language, code.strip()

    payload: object | None = None
    for raw in (candidate, text.strip()):
        try:
            payload = json.loads(raw)
            parsed = _extract_from_payload(payload)
            if parsed is not None:
                return parsed
        except Exception:
            try:
                payload = ast.literal_eval(raw)
                parsed = _extract_from_payload(payload)
                if parsed is not None:
                    return parsed
            except Exception:
                continue

    fallback_code = _extract_powershell_code_fallback(candidate) or _extract_powershell_code_fallback(
        text
    )
    if fallback_code:
        language_match = re.search(
            r'"?language"?\s*:\s*"?([a-zA-Z0-9_\-]+)"?',
            f"{candidate}\n{text}",
            re.IGNORECASE,
        )
        language = (language_match.group(1).strip().lower() if language_match else "powershell") or "powershell"
        return language, fallback_code
    return None


def _run_powershell_once(code: str) -> str:
    """Run a PowerShell command once and return output."""
    code, repair_note = _repair_start_process_code(code)
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-Command", code],
        capture_output=True,
        text=True,
        timeout=25,
    )
    out = (proc.stdout or "").strip()
    err = (proc.stderr or "").strip()
    if out and err:
        combined = f"{out}\n{err}"
        return f"{repair_note}\n{combined}" if repair_note else combined
    if out:
        return f"{repair_note}\n{out}" if repair_note else out
    if err:
        return f"{repair_note}\n{err}" if repair_note else err
    if repair_note:
        return repair_note
    return "(no output)"


def _run_python_desktop_once(code: str) -> str:
    """Run a constrained Python desktop snippet (pyautogui/time only)."""
    if not code or len(code) > 5000:
        return "Error: Python snippet is empty or too large."

    lowered = code.lower()
    blocked_patterns = (
        "import os",
        "import subprocess",
        "import socket",
        "from os",
        "from subprocess",
        "from socket",
        "__import__",
        "eval(",
        "exec(",
        "open(",
        "pathlib",
        "shutil",
        "requests",
        "while true",
    )
    if any(p in lowered for p in blocked_patterns):
        return "Error: Unsupported Python snippet for desktop execution."

    try:
        import pyautogui
        import time

        pyautogui.FAILSAFE = False
    except Exception as e:
        return f"Error: Desktop Python runtime unavailable: {e}"

    safe_globals = {"pyautogui": pyautogui, "time": time, "__builtins__": {}}
    lines = [ln.strip() for ln in code.replace("\r", "\n").splitlines() if ln.strip()]
    if not lines:
        return "Error: Python snippet is empty."

    try:
        for line in lines:
            lowered_line = line.lower()
            if lowered_line.startswith("#"):
                continue
            if lowered_line.startswith("import "):
                modules = [m.strip().lower() for m in line[7:].split(",") if m.strip()]
                if not modules or any(m not in {"pyautogui", "time"} for m in modules):
                    return "Error: Only 'import pyautogui' and 'import time' are allowed."
                continue
            if lowered_line.startswith("from "):
                return "Error: 'from ... import ...' is not allowed in desktop snippets."
            if not (lowered_line.startswith("pyautogui.") or lowered_line.startswith("time.sleep(")):
                return "Error: Only pyautogui.* and time.sleep() commands are allowed."
            exec(line, safe_globals, {})
        return "(no output)"
    except Exception as e:
        return f"Error: Python desktop command failed: {e}"


def _repair_start_process_code(code: str) -> tuple[str, str | None]:
    """Fix common broken Start-Process paths (wrong username/malformed app path)."""
    match = re.search(
        r"start-process(?:\s+-filepath)?\s+(?:\"([^\"]+)\"|'([^']+)'|([^\s|;]+))",
        code,
        re.IGNORECASE,
    )
    if not match:
        return code, None

    raw_target = (match.group(1) or match.group(2) or match.group(3) or "").strip().strip("\"'`")
    if not raw_target:
        return code, None

    target_path = Path(raw_target)
    if target_path.exists():
        return code, None

    candidates: list[str] = []

    # If model hallucinated username in C:\Users\<name>\..., rewrite to current home.
    user_path_match = re.match(r"^[a-zA-Z]:\\Users\\[^\\]+\\(.+)$", raw_target, re.IGNORECASE)
    if user_path_match:
        suffix = user_path_match.group(1)
        home = Path.home()
        home_drive = home.drive or "C:"
        candidates.append(str(Path(f"{home_drive}\\Users\\{home.name}\\{suffix}")))

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    base_name = Path(raw_target).name
    if local_app_data and base_name:
        local = Path(local_app_data)
        candidates.append(str(local / base_name))
        stem = Path(base_name).stem
        if stem:
            candidates.append(str(local / stem / base_name))
        if base_name.lower() == "whatsapp.exe":
            candidates.append(str(local / "WhatsApp" / "WhatsApp.exe"))

    # De-duplicate while preserving order.
    deduped: list[str] = []
    for cand in candidates:
        if cand and cand not in deduped:
            deduped.append(cand)

    for cand in deduped:
        try:
            if Path(cand).exists():
                fixed_segment = f'Start-Process -FilePath "{cand}"'
                fixed_code = f"{code[:match.start()]}{fixed_segment}{code[match.end():]}"
                note = f"Adjusted launch path to existing file: {cand}"
                return fixed_code, note
        except Exception:
            continue

    return code, None


def _extract_powershell_code_fallback(text: str) -> str | None:
    """Best-effort extraction of safe PowerShell code from malformed payload text."""
    if not text:
        return None

    cleaned = text.replace("\r", "\n")
    lowered = cleaned.lower()

    # Common malformed case:
    # Start-Process
    # "C:\Path\To\App.exe""
    if "start-process" in lowered:
        quoted_target = re.search(
            r"start-process(?:\s+-filepath)?[ \t]+(?:\"([^\"]+)\"|'([^']+)')",
            cleaned,
            re.IGNORECASE,
        )
        if quoted_target:
            target = (quoted_target.group(1) or quoted_target.group(2) or "").strip().strip("\"'`")
            target = target.split("\\n", 1)[0].split("\\r", 1)[0].split("\n", 1)[0].split("\r", 1)[0]
            target = target.rstrip(",;")
            if target:
                return f'Start-Process -FilePath "{target}"'

        path_match = re.search(r"([a-zA-Z]:\\[^\n\r\"']+?\.exe)", cleaned)
        if path_match:
            path = path_match.group(1).strip()
            return f'Start-Process -FilePath "{path}"'

        # Fallback for token-like targets (e.g., taskmgr.exe)
        start_match = re.search(r"start-process(?:\s+-filepath)?[ \t]+([^\s{}\n\r]+)", cleaned, re.IGNORECASE)
        if start_match:
            target = start_match.group(1).strip().strip("\"'`")
            target = target.split("\\n", 1)[0].split("\\r", 1)[0].split("\n", 1)[0].split("\r", 1)[0]
            target = target.rstrip(",;")
            if target:
                is_uri_like = ":" in target and "\\" not in target
                if not is_uri_like and not target.lower().endswith(".exe") and "\\" not in target:
                    target += ".exe"
                return f'Start-Process -FilePath "{target}"'

    if "explorer.exe" in lowered or lowered.strip().startswith("explorer "):
        path_match = re.search(r"explorer(?:\.exe)?\s+\"?([a-zA-Z]:\\[^\n\r\"']+)\"?", cleaned, re.IGNORECASE)
        if path_match:
            path = path_match.group(1).strip().rstrip(",;")
            return f'explorer.exe "{path}"'
        return "explorer.exe"

    if "get-process" in lowered and "mainwindowtitle" in lowered and "measure-object" in lowered:
        return (
            "Get-Process | Where-Object {$_.MainWindowTitle -ne ''} | "
            "Measure-Object | Select-Object -ExpandProperty Count"
        )

    if "get-process" in lowered and "mainwindowtitle" in lowered:
        return (
            "Get-Process | Where-Object {$_.MainWindowTitle -ne ''} | "
            "Select-Object ProcessName, MainWindowTitle"
        )

    if "get-process" in lowered and "workingset" in lowered:
        return (
            "Get-Process | Sort-Object -Property WorkingSet -Descending | "
            "Select-Object -First 5 -Property ProcessName, WorkingSet, CPU, Id"
        )

    return None


class OpenInterpreterAgent:
    """Wraps Open Interpreter for autonomous task execution.

    In the Agent SDK architecture, this serves as the EXECUTOR layer:
    - Executes code and system commands
    - Handles file operations
    - Provides sandboxed execution environment
    """

    AI_DESKTOP_ACTIONS: set[str] = set()

    def __init__(self, settings: Settings):
        self.settings = settings
        self._interpreter = None
        self._stop_flag = False
        self._semaphore = asyncio.Semaphore(1)
        self._pending_dangerous_intent: dict[str, Any] | None = None
        self._desktop_context_cache: dict | None = None
        self._desktop_context_cache_at: float = 0.0
        if not self.AI_DESKTOP_ACTIONS:
            try:
                from Mudabbir.tools.capabilities.registry import DEFAULT_CAPABILITY_REGISTRY

                self.AI_DESKTOP_ACTIONS = DEFAULT_CAPABILITY_REGISTRY.allowed_actions_stage_a()
            except Exception:
                self.AI_DESKTOP_ACTIONS = {
                    "launch_start_app",
                    "open_settings_page",
                    "close_app",
                    "system_power",
                    "shutdown_schedule",
                    "system_info",
                    "network_tools",
                    "file_tools",
                    "window_control",
                    "process_tools",
                    "service_tools",
                    "background_tools",
                    "startup_tools",
                    "clipboard_tools",
                    "browser_control",
                    "user_tools",
                    "task_tools",
                    "registry_tools",
                    "disk_tools",
                    "security_tools",
                    "web_tools",
                    "hardware_tools",
                    "update_tools",
                    "ui_tools",
                    "automation_tools",
                    "app_tools",
                    "info_tools",
                    "dev_tools",
                    "shell_tools",
                    "office_tools",
                    "remote_tools",
                    "search_tools",
                    "performance_tools",
                    "media_tools",
                    "browser_deep_tools",
                    "maintenance_tools",
                    "driver_tools",
                    "power_user_tools",
                    "screenshot_tools",
                    "text_tools",
                    "api_tools",
                    "vision_tools",
                    "threat_tools",
                    "content_tools",
                    "list_processes",
                    "battery_status",
                    "volume",
                    "brightness",
                    "mouse_move",
                    "click",
                    "press_key",
                    "type_text",
                    "hotkey",
                    "focus_window",
                    "search_files",
                    "list_windows",
                    "desktop_overview",
                    "ui_target",
                    "move_mouse_to_desktop_file",
                }
        self._initialize()

    def _initialize(self) -> None:
        """Initialize the Open Interpreter instance."""
        try:
            warnings.filterwarnings(
                "ignore",
                message="pkg_resources is deprecated as an API.*",
                category=UserWarning,
            )
            warnings.filterwarnings(
                "ignore",
                message=r".*asyncio\.iscoroutinefunction.*deprecated.*",
                category=DeprecationWarning,
            )
            warnings.filterwarnings(
                "ignore",
                message=r"remove second argument of ws_handler",
                category=DeprecationWarning,
            )
            from interpreter import interpreter

            from Mudabbir.llm.client import resolve_llm_client

            # Configure interpreter
            interpreter.auto_run = bool(getattr(self.settings, "oi_auto_run", True))
            interpreter.loop = bool(getattr(self.settings, "oi_loop", False))
            interpreter.verbose = False
            interpreter.plain_text_display = False
            # Silence Open Interpreter terminal rendering; Mudabbir streams via MessageBus.
            try:
                interpreter.display_message = lambda *args, **kwargs: None
            except Exception:
                pass
            # Compatibility shim for Open Interpreter builds where
            # interpreter.core.respond references display_markdown_message
            # without importing it, causing NameError at runtime.
            try:
                respond_mod = importlib.import_module("interpreter.core.respond")
                if not hasattr(respond_mod, "display_markdown_message"):
                    try:
                        from interpreter.terminal_interface.utils.display_markdown_message import (
                            display_markdown_message as _display_markdown_message,
                        )
                    except Exception:

                        def _display_markdown_message(msg):
                            interpreter.display_message(str(msg))

                    respond_mod.display_markdown_message = _display_markdown_message
            except Exception as patch_err:
                logger.debug(
                    "Open Interpreter compatibility shim skipped: %s",
                    patch_err,
                )

            # Set LLM based on resolved provider
            llm = resolve_llm_client(self.settings)

            if llm.is_ollama:
                interpreter.llm.model = f"ollama/{llm.model}"
                interpreter.llm.api_base = llm.ollama_host
                logger.info(f"🤖 Using Ollama: {llm.model}")
            elif llm.is_gemini and llm.api_key:
                # Force Gemini through AI Studio's OpenAI-compatible endpoint.
                # This avoids LiteLLM auto-routing Gemini models to Vertex AI
                # (which requires Google ADC and causes runtime failures).
                model_name = llm.model if llm.model.startswith("openai/") else f"openai/{llm.model}"
                interpreter.llm.model = model_name
                interpreter.llm.api_key = llm.api_key
                interpreter.llm.api_base = llm.openai_compatible_base_url
                if str(llm.model).lower().startswith("gemini-3"):
                    # Gemini 3 models are more stable at temperature=1.0.
                    try:
                        interpreter.llm.temperature = 1.0
                    except Exception:
                        pass
                # Some dependency stacks read these keys directly.
                os.environ.setdefault("GOOGLE_API_KEY", llm.api_key)
                os.environ.setdefault("GEMINI_API_KEY", llm.api_key)
                logger.info(f"🤖 Using Gemini (AI Studio): {llm.model}")
                logger.info(
                    "Gemini is running via OpenAI-compatible AI Studio endpoint"
                )
            elif llm.api_key:
                interpreter.llm.model = llm.model
                interpreter.llm.api_key = llm.api_key
                logger.info(f"🤖 Using {llm.provider.title()}: {llm.model}")

            # Safety settings
            safe_mode = str(getattr(self.settings, "oi_safe_mode", "ask")).strip().lower()
            interpreter.safe_mode = safe_mode or "ask"

            self._interpreter = interpreter
            logger.info("=" * 50)
            logger.info("🔧 EXECUTOR: Open Interpreter initialized")
            logger.info("   └─ Role: Code execution, file ops, system commands")
            logger.info("=" * 50)

        except ImportError as e:
            missing = getattr(e, "name", None)
            if missing == "pkg_resources":
                logger.error(
                    "❌ Open Interpreter import failed: missing pkg_resources. "
                    "Use: pip install \"setuptools<81\". Root error: %s",
                    e,
                )
            elif missing:
                logger.error(
                    "❌ Open Interpreter import failed (missing module: %s). "
                    "Install/fix dependency and retry. Root error: %s",
                    missing,
                    e,
                )
            else:
                logger.error(
                    "❌ Open Interpreter import failed. Install/fix dependency and retry. Root error: %s",
                    e,
                )
            self._interpreter = None
        except Exception as e:
            logger.error(f"❌ Failed to initialize Open Interpreter: {e}")
            self._interpreter = None

    def _is_gui_request(self, message: str) -> bool:
        """Detect GUI/control intents with intent-map first, then alias-derived fallback."""
        try:
            resolved = resolve_windows_intent(message or "")
            if resolved.matched:
                return True
        except Exception:
            pass
        normalized = _normalize_text_for_match(str(message or ""))
        if not normalized:
            return False
        return any(alias in normalized for alias in WINDOWS_ALIAS_FALLBACK)

    def _wants_pointer_control(self, message: str) -> bool:
        """Pointer-intent check using intent-map/context, not rigid keyword lists."""
        text = str(message or "")
        try:
            resolved = resolve_windows_intent(text)
            cap = str(getattr(resolved, "capability_id", "") or "")
            action = str(getattr(resolved, "action", "") or "")
            if resolved.matched and (
                cap.startswith("mouse.")
                    or action in {"mouse_move", "click"}
            ):
                return True
        except Exception:
            pass
        normalized = _normalize_text_for_match(text)
        try:
            pointer_aliases: set[str] = set()
            for rule in WINDOWS_INTENT_RULES:
                cap_id = str(getattr(rule, "capability_id", "") or "")
                action_id = str(getattr(rule, "action", "") or "")
                if not (cap_id.startswith("mouse.") or action_id in {"mouse_move", "click"}):
                    continue
                for alias in tuple(getattr(rule, "aliases", ()) or ()):
                    alias_norm = _normalize_text_for_match(str(alias or ""))
                    if alias_norm:
                        pointer_aliases.add(alias_norm)
            if any(alias in normalized for alias in pointer_aliases):
                return True
        except Exception:
            pass
        # Context fallback: movement/click verbs + coordinates usually imply pointer control.
        has_coords = bool(re.search(r"-?\d+\s*[,،]\s*-?\d+", normalized))
        has_pointer_verb = bool(
            re.search(r"\b(move|click|hover|drag|drop|scroll|حرك|حرّك|انقر|اضغط|اسحب|مرر)\b", normalized, re.IGNORECASE)
        )
        return has_coords and has_pointer_verb

    def _is_task_manager_request(self, message: str) -> bool:
        """Detect requests that mention Task Manager."""
        if _contains_any(message, TASK_MANAGER_KEYWORDS):
            return True
        return bool(TASK_MANAGER_PATTERN.search(message or ""))

    def _wants_top_process_list(self, message: str) -> bool:
        """Detect if the user also asked for a top process snapshot."""
        return _contains_any(message, TOP_PROCESS_KEYWORDS)

    def _get_top_processes_by_memory(self) -> list[str]:
        """Get top 5 processes by memory usage."""
        # Preferred path: psutil (faster and structured)
        try:
            import psutil

            rows: list[tuple[str, int, float]] = []
            for proc in psutil.process_iter(["name", "pid", "memory_info"]):
                try:
                    info = proc.info
                    mem = info.get("memory_info")
                    rss = float(mem.rss) if mem else 0.0
                    rows.append((info.get("name") or "unknown", int(info.get("pid") or 0), rss))
                except Exception:
                    continue

            rows.sort(key=lambda r: r[2], reverse=True)
            lines = []
            for idx, (name, pid, rss) in enumerate(rows[:5], start=1):
                lines.append(f"{idx}. {name} (PID {pid}) - {rss / (1024 * 1024):.1f} MB")
            if lines:
                return lines
        except Exception:
            pass

        # Fallback path: PowerShell JSON output
        ps_cmd = (
            "Get-Process | Sort-Object -Property WorkingSet -Descending | "
            "Select-Object -First 5 -Property ProcessName,Id,"
            "@{Name='MemoryMB';Expression={[math]::Round($_.WorkingSet64/1MB,1)}} | "
            "ConvertTo-Json -Depth 3 -Compress"
        )
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_cmd],
                capture_output=True,
                text=True,
                timeout=20,
            )
            raw = (proc.stdout or "").strip()
            if not raw:
                return []

            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                parsed = [parsed]

            lines = []
            for idx, item in enumerate(parsed[:5], start=1):
                name = item.get("ProcessName", "unknown")
                pid = item.get("Id", 0)
                mem_mb = item.get("MemoryMB", 0)
                lines.append(f"{idx}. {name} (PID {pid}) - {mem_mb} MB")
            return lines
        except Exception:
            return []

    async def _try_direct_task_manager_response(self, message: str) -> dict | None:
        """Run a deterministic local flow for Task Manager requests."""
        if not self._is_task_manager_request(message):
            return None

        arabic = _contains_arabic(message)
        notes: list[str] = []

        # Open Task Manager once.
        try:
            subprocess.Popen(["taskmgr.exe"])
            notes.append("تم فتح مدير المهام مرة واحدة." if arabic else "Opened Task Manager once.")
        except Exception as e:
            if arabic:
                notes.append(f"تعذر فتح مدير المهام: {e}")
            else:
                notes.append(f"Could not open Task Manager: {e}")

        # Move mouse to confirm GUI control works.
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            size = pyautogui.size()
            x = int(size.width * 0.55)
            y = int(size.height * 0.35)
            pyautogui.moveTo(x, y, duration=0.2)
            if arabic:
                notes.append("تم تحريك الماوس إلى الموضع المطلوب.")
            else:
                notes.append("Moved mouse to the requested position.")
        except Exception as e:
            if arabic:
                notes.append(f"التحكم بالماوس غير متاح: {e}")
            else:
                notes.append(f"Mouse control unavailable: {e}")

        if self._wants_top_process_list(message):
            top5 = self._get_top_processes_by_memory()
            if top5:
                notes.append("أعلى 5 عمليات استهلاكًا للذاكرة:" if arabic else "Top 5 processes by memory:")
                notes.extend(top5)
            else:
                if arabic:
                    notes.append("تعذر قراءة أعلى العمليات من الجهاز المحلي.")
                else:
                    notes.append("Could not read top processes from the local machine.")

        return {"type": "message", "content": "\n".join(notes)}

    def _open_whatsapp_app(self) -> tuple[bool, str]:
        """Open WhatsApp desktop app if installed."""
        def _is_whatsapp_running() -> bool:
            try:
                import psutil

                for proc in psutil.process_iter(["name"]):
                    name = (proc.info.get("name") or "").lower()
                    if "whatsapp" in name:
                        return True
            except Exception:
                pass
            return False

        def _focus_whatsapp_window() -> bool:
            try:
                import pygetwindow as gw

                wins = [
                    w
                    for w in gw.getAllWindows()
                    if "whatsapp" in str(getattr(w, "title", "") or "").lower()
                ]
                if not wins:
                    return False
                win = wins[0]
                try:
                    if bool(getattr(win, "isMinimized", False)):
                        win.restore()
                except Exception:
                    pass
                win.activate()
                return True
            except Exception:
                return False

        def _wait_for_whatsapp(timeout_sec: float = 8.0) -> bool:
            deadline = time.time() + max(0.5, timeout_sec)
            while time.time() < deadline:
                if _is_whatsapp_running():
                    _focus_whatsapp_window()
                    return True
                time.sleep(0.35)
            return False

        candidates: list[Path] = []
        local = os.environ.get("LOCALAPPDATA", "")
        if local:
            local_path = Path(local)
            candidates.extend(
                [
                    local_path / "WhatsApp" / "WhatsApp.exe",
                    local_path / "Programs" / "WhatsApp" / "WhatsApp.exe",
                ]
            )
        candidates.append(Path.home() / "AppData" / "Local" / "WhatsApp" / "WhatsApp.exe")

        for candidate in candidates:
            try:
                if candidate.exists():
                    subprocess.Popen([str(candidate)])
                    if _wait_for_whatsapp():
                        return True, str(candidate)
                    return False, f"Tried launch but process did not appear: {candidate}"
            except Exception:
                continue

        # Fallback protocol/appx launch paths.
        fallbacks = [["cmd", "/c", "start", "", "whatsapp:"]]

        # Dynamic AppUserModelId detection for Store-installed WhatsApp variants.
        try:
            probe = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    (
                        "Get-StartApps | Where-Object {$_.Name -match 'WhatsApp'} | "
                        "Select-Object -First 1 -ExpandProperty AppID"
                    ),
                ],
                capture_output=True,
                text=True,
                timeout=8,
            )
            app_id = (probe.stdout or "").strip().splitlines()
            if app_id:
                fallbacks.append(["cmd", "/c", "start", "", f"shell:AppsFolder\\{app_id[0].strip()}"])
        except Exception:
            pass

        # Legacy known AppID fallback.
        fallbacks.append(
            ["cmd", "/c", "start", "", "shell:AppsFolder\\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"]
        )

        for cmd in fallbacks:
            try:
                subprocess.Popen(cmd, shell=False)
                if _wait_for_whatsapp():
                    return True, "fallback protocol launch"
            except Exception:
                continue
        return False, "WhatsApp.exe not found in common install paths."

    def _count_open_windows(self) -> int | None:
        """Count visible windows with non-empty titles."""
        try:
            import pygetwindow as gw

            windows = [w for w in gw.getAllWindows() if str(getattr(w, "title", "") or "").strip()]
            return len(windows)
        except Exception:
            return None

    async def _get_live_desktop_context(self) -> dict | None:
        """Return a lightweight live desktop snapshot with short-term caching."""
        now = time.time()
        if (
            isinstance(self._desktop_context_cache, dict)
            and (now - float(self._desktop_context_cache_at)) < 2.5
        ):
            return self._desktop_context_cache

        try:
            from Mudabbir.tools.builtin.desktop import DesktopTool

            raw = await DesktopTool().execute(action="desktop_overview")
            if not isinstance(raw, str) or raw.lower().startswith("error:"):
                return self._desktop_context_cache
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                self._desktop_context_cache = parsed
                self._desktop_context_cache_at = now
                return parsed
        except Exception:
            return self._desktop_context_cache
        return self._desktop_context_cache

    @staticmethod
    def _extract_first_json_object(text: str) -> dict | None:
        """Extract the first JSON object from a noisy LLM response."""
        if not text:
            return None
        cleaned = text.strip()
        if cleaned.startswith("```"):
            cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
            cleaned = re.sub(r"\s*```$", "", cleaned)

        start = cleaned.find("{")
        if start < 0:
            return None

        depth = 0
        end = -1
        in_string = False
        escape = False
        for idx, ch in enumerate(cleaned[start:], start=start):
            if in_string:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == '"':
                    in_string = False
                continue
            if ch == '"':
                in_string = True
                continue
            if ch == "{":
                depth += 1
                continue
            if ch == "}":
                depth -= 1
                if depth == 0:
                    end = idx
                    break

        if end < 0:
            return None
        candidate = cleaned[start : end + 1]
        try:
            parsed = json.loads(candidate)
            return parsed if isinstance(parsed, dict) else None
        except Exception:
            return None

    async def _llm_one_shot_text(
        self,
        *,
        system_prompt: str,
        user_prompt: str,
        max_tokens: int = 500,
        temperature: float = 0.1,
    ) -> str | None:
        """Run a single non-streaming LLM completion using the active provider."""
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

            # Anthropic
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
        except Exception as e:
            logger.debug("AI one-shot completion failed: %s", e)
            return None

    async def _llm_plan_desktop_actions(
        self,
        *,
        message: str,
        history: list[dict] | None,
    ) -> dict | None:
        """Use AI to parse desktop intent into structured desktop tool actions."""
        history_lines: list[str] = []
        for item in (history or [])[-4:]:
            if not isinstance(item, dict):
                continue
            role = str(item.get("role", "user"))
            content = str(item.get("content", "") or "").strip()
            if not content:
                continue
            if (
                _looks_like_execute_noise(content)
                or _looks_like_execute_payload_fragment(content)
                or _looks_like_raw_command_leak(content)
            ):
                continue
            history_lines.append(f"{role}: {content[:220]}")
        history_blob = "\n".join(history_lines) if history_lines else "(none)"
        live_desktop = await self._get_live_desktop_context()
        if isinstance(live_desktop, dict):
            screen = live_desktop.get("screen") if isinstance(live_desktop.get("screen"), dict) else {}
            windows = (
                live_desktop.get("windows") if isinstance(live_desktop.get("windows"), dict) else {}
            )
            window_items = windows.get("items") if isinstance(windows.get("items"), list) else []
            title_parts: list[str] = []
            for item in window_items[:6]:
                if not isinstance(item, dict):
                    continue
                title = str(item.get("title", "") or "").strip()
                if title:
                    title_parts.append(title[:60])
            live_blob = (
                f"screen={screen.get('width', '?')}x{screen.get('height', '?')}; "
                f"open_windows={windows.get('count', '?')}; "
                f"top_titles={title_parts or ['(none)']}"
            )
        else:
            live_blob = "(unavailable)"

        system_prompt = (
            "You are a Windows desktop action planner. "
            "Convert user intent (any language) into JSON actions for desktop automation.\n"
            "Return ONLY valid JSON, no markdown."
        )
        user_prompt = f"""
Conversation context:
{history_blob}

Live desktop context:
{live_blob}

User request:
{message}

Available actions and parameters:
- launch_start_app: query
- open_settings_page: page
- close_app: process_name, force?
- list_processes: only_windowed?, max_results?
- system_power: mode=lock|sleep|hibernate|logoff|shutdown|restart|bios|rename_pc|screen_off|power_plan_balanced|power_plan_saver|power_plan_high, name?
- shutdown_schedule: mode=set|cancel, minutes?
- system_info: mode=uptime|windows_version|about|battery
- battery_status: no params
- volume: mode=get|set|up|down|max|min|mute|unmute, level?, delta?
- brightness: mode=get|set|up|down|max|min, level?, delta?
- network_tools: mode=ip_internal|ip_external|flush_dns|renew_ip|ping|wifi_on|wifi_off|wifi_passwords|disconnect_wifi|disconnect_current_network|connect_wifi|route_table|net_scan|open_settings|open_ports|port_owner|hotspot_on|hotspot_off|file_sharing_on|file_sharing_off|shared_folders|server_online|last_login_events, host?, port?
- file_tools: mode=open_documents|open_downloads|open_pictures|open_videos|organize_desktop|organize_desktop_semantic|smart_rename|smart_rename_content|create_folder|delete|rename|copy|move|zip|unzip|search_ext|folder_size|open_cmd_here|open_powershell_here|empty_recycle_bin, path?, target?, name?, ext?, permanent?
- window_control: mode=minimize|maximize|restore|close_current|show_desktop|undo_show_desktop|split_left|split_right|alt_tab|task_view|bring_to_front|set_focus|hide|show|minimize_to_tray|restore_from_tray|coords|move_resize|transparency|borderless_on|borderless_off|disable_close_on|disable_close_off|span_all_screens|move_next_monitor_right|move_next_monitor_left|always_on_top_on|always_on_top_off, app?, x?, y?, width?, height?, opacity?
- process_tools: mode=list|top_cpu|top_ram|kill_pid|kill_name|close_browsers|close_office|kill_high_cpu|restart_explorer|kill_unresponsive|path_by_pid|cpu_by_pid|ram_by_pid|threads_by_pid|start_time_by_pid|app_uptime|set_priority|suspend_pid|resume_pid|unresponsive, pid?, name?, priority?, threshold?
- service_tools: mode=list|start|stop|restart|startup|user_services, name?, startup?
- background_tools: mode=count_background|list_minimized_windows|list_visible_windows|ghost_apps|activity_time|network_usage_per_app|camera_usage|mic_usage|camera_usage_now|mic_usage_now|wake_lock_apps|process_paths, max_results?
- startup_tools: mode=list|impact_time|impact_breakdown|registry_startups|folder_startups|disable|enable|detect_new|watch_new_live|signature_check|full_audit, name?, seconds?, monitor_seconds?, notify?
- clipboard_tools: mode=clear|history
- browser_control: mode=new_tab|close_tab|reopen_tab|next_tab|prev_tab|reload|incognito|history|downloads|find|zoom_in|zoom_out|zoom_reset|save_pdf|home
- user_tools: mode=list|create|delete|set_password|set_type, username?, password?, group?
- task_tools: mode=list|create|run|delete, name?, command?, trigger?
- registry_tools: mode=query|add_key|delete_key|set_value|backup|restore, key, value_name?, value_data?, value_type?
- disk_tools: mode=smart_status|temp_files_clean|disk_usage|chkdsk_scan|prefetch_clean|logs_clean|defrag, drive?
- security_tools: mode=firewall_status|firewall_enable|firewall_disable|block_port|unblock_rule|disable_usb|enable_usb|disable_camera|enable_camera|logged_in_users|remote_sessions_list|recent_files|recent_files_clear|current_connections_ips|admin_processes|failed_audit_logins|close_remote_sessions|intrusion_summary, port?, rule_name?
- web_tools: mode=open_url|download_file|weather, url?, city?
- hardware_tools: mode=cpu_info|cores_info|gpu_info|gpu_temp|mobo_serial|mobo_model|ram_info|ram_speed_type|battery_report|battery_minutes|battery_cycle_count|smart_status, drive?
- update_tools: mode=list_updates|last_update_time|check_updates|install_kb|winsxs_cleanup|stop_background_updates, target?
- ui_tools: mode=dark_mode|light_mode|transparency_on|transparency_off|taskbar_autohide_on|taskbar_autohide_off|night_light_on|night_light_off|desktop_icons_show|desktop_icons_hide|game_mode_on|game_mode_off
- automation_tools: mode=delay|popup|tts|repeat_key|mouse_keys_toggle|timer|mouse_lock_on|mouse_lock_off|mouse_lock_region|anti_idle_f5|battery_guard|screen_off_mouse_lock, seconds?, monitor_seconds?, text?, key?, repeat_count?, x?, y?, width?, height?
- app_tools: mode=open_default_browser|open_notepad|open_calc|open_paint|open_task_manager|open_control_panel|open_store|open_registry|open_camera|open_calendar|open_mail|open_chrome|open_edge|open_add_remove_programs|open_volume_mixer|open_mic_settings|close_all_apps|open_app, app?
- info_tools: mode=timezone_get|timezone_set|system_language|windows_product_key|model_info|windows_install_date|refresh_rate, timezone?
- dev_tools: mode=open_cmd_admin|open_powershell_admin|open_disk_management|open_device_manager|open_perfmon|open_event_viewer|open_services|open_registry|sfc_scan|chkdsk|env_vars|runtime_versions|git_last_log|open_editor|event_errors|bsod_summary|analyze_bsod|interpret_powershell, drive?, path?, editor?, max_results?, target?, text?, force?
- shell_tools: mode=quick_settings|notifications|search|run|file_explorer|quick_link_menu|task_view|new_virtual_desktop|next_virtual_desktop|prev_virtual_desktop|close_virtual_desktop|emoji_panel|start_menu|refresh|magnifier_open|magnifier_close|narrator_toggle
- office_tools: mode=open_word_new|silent_print|docx_to_pdf, path?, target?
- remote_tools: mode=rdp_open|vpn_connect|vpn_disconnect, host?
- search_tools: mode=search_text|files_larger_than|modified_today|find_images|find_videos|count_files|search_open_windows|search_open_windows_content, folder?, pattern?, size_mb?
- performance_tools: mode=top_cpu|top_ram|top_disk|total_ram_percent|total_cpu_percent|cpu_clock|available_ram|pagefile_used|disk_io_rate|gpu_util|top_gpu_processes|gpu_temp|empty_ram|cpu_popup|kill_high_cpu, threshold?
- media_tools: mode=stop_all_media|youtube_open|media_next|media_prev|play_pause, url?
- browser_deep_tools: mode=multi_open|clear_chrome_cache|clear_edge_cache, urls?
- maintenance_tools: mode=empty_ram|winsxs_cleanup|temp_clean
- driver_tools: mode=drivers_list|drivers_backup|drivers_issues|updates_pending
- power_user_tools: mode=airplane_on|airplane_off|god_mode|invert_colors
- screenshot_tools: mode=full|snipping_tool|window_active|region, x?, y?, width?, height?, path?
- text_tools: mode=text_to_file|clipboard_to_file|word_count|search_replace_files|create_batch, path?, content?, folder?, pattern?, replace_with?
- api_tools: mode=currency|weather_city|translate_quick, target?, city?, text?
- vision_tools: mode=describe_screen|ocr_screen|ocr_image|ocr_active_window|ocr_region|copy_ocr_to_clipboard, path?, x?, y?, width?, height?
- threat_tools: mode=file_hash|sha256|vt_lookup|external_ips|suspicious_connections|suspicious_apps|behavior_scan, path?, target?, max_results?
- content_tools: mode=draft_reply|email_draft|email_auto_reply_docx|auto_reply_word|draft_to_word|save_word_draft|text_numbers_to_excel|text_to_excel, content?, path?, target?
- mouse_move: x, y, duration?
- click: x?, y?, button?, clicks?
- press_key: key
- type_text: text, press_enter?
- hotkey: keys[]
- focus_window: window_title?, process_name?, timeout_sec?
- search_files: query, max_results?
- list_windows: include_untitled?, limit?
- desktop_overview: no params
- ui_target: control_name, window_title?, process_name?, control_type?, interaction=move|click|double_click|right_click, timeout_sec?
- move_mouse_to_desktop_file: query, timeout_sec?

Rules:
1) If request is NOT desktop/system control, return should_execute=false.
2) If request is desktop control, return should_execute=true and actions list.
3) Support multilingual understanding (Arabic, Russian, English, etc.).
4) For combined requests (e.g. volume + brightness), output multiple actions.
5) Use launch_start_app for opening apps, and focus_window when intent says switch/focus.
6) Keep max 6 actions.

Examples:
- "كم نسبة الصوت والاضاءة الان" -> actions: [volume get, brightness get]
- "اعمل الصوت 33 والاضاءة كمان" -> actions: [volume set 33, brightness set 33]
- "увеличь громкость до 40 и яркость тоже" -> actions: [volume set 40, brightness set 40]

Required JSON schema:
{{
  "should_execute": true|false,
  "reason": "short reason",
  "actions": [{{"action": "name", "params": {{...}}}}],
  "response_language": "ar|en|ru|auto"
}}
"""
        raw = await self._llm_one_shot_text(
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            max_tokens=700,
            temperature=0.0,
        )
        if not raw:
            return None
        plan = self._extract_first_json_object(raw)
        if plan is None:
            repair_prompt = (
                "Convert the following model output into STRICT JSON that matches the requested schema. "
                "Output JSON only.\n\n"
                f"Output to repair:\n{raw}"
            )
            repaired = await self._llm_one_shot_text(
                system_prompt=system_prompt,
                user_prompt=repair_prompt,
                max_tokens=500,
                temperature=0.0,
            )
            if repaired:
                plan = self._extract_first_json_object(repaired)
        if not isinstance(plan, dict):
            return None
        return plan

    async def _execute_ai_desktop_actions(self, actions: list[dict]) -> list[dict]:
        """Execute AI-planned desktop actions via DesktopTool with allowlist validation."""
        from Mudabbir.tools.builtin.desktop import DesktopTool

        tool = DesktopTool()
        outputs: list[dict] = []
        for step in actions[:6]:
            if not isinstance(step, dict):
                continue
            action = str(step.get("action", "") or "").strip().lower()
            params_raw = step.get("params", {})
            params = params_raw if isinstance(params_raw, dict) else {}
            if action not in self.AI_DESKTOP_ACTIONS:
                outputs.append(
                    {"action": action or "unknown", "ok": False, "error": "action not allowed"}
                )
                continue
            try:
                raw = await tool.execute(action=action, **params)
                if isinstance(raw, str) and raw.lower().startswith("error:"):
                    outputs.append(
                        {
                            "action": action,
                            "ok": False,
                            "error": raw.replace("Error: ", "", 1),
                            "params": params,
                        }
                    )
                    continue
                parsed: dict | list | str = raw
                if isinstance(raw, str):
                    try:
                        parsed = json.loads(raw)
                    except Exception:
                        parsed = raw
                outputs.append(
                    {
                        "action": action,
                        "ok": True,
                        "params": params,
                        "result": parsed,
                    }
                )
            except Exception as e:
                outputs.append(
                    {
                        "action": action,
                        "ok": False,
                        "error": str(e),
                        "params": params,
                    }
                )
        return outputs

    async def _try_intent_map_desktop_response(self, message: str) -> dict | None:
        """Deterministic intent-map route for Windows desktop/system commands."""
        if not bool(getattr(self.settings, "oi_deterministic_desktop_first", True)):
            return None

        text = message or ""
        arabic = _contains_arabic(text)
        normalized = _normalize_text_for_match(text)
        cancel_tokens = ("cancel", "stop", "no", "لا", "الغاء", "إلغاء", "وقف")

        if self._pending_dangerous_intent:
            pending = self._pending_dangerous_intent
            if is_confirmation_message(text):
                self._pending_dangerous_intent = None
                resolution = pending
            elif any(tok in normalized for tok in cancel_tokens):
                self._pending_dangerous_intent = None
                return {
                    "type": "result",
                    "content": (
                        "تم إلغاء العملية الخطرة."
                        if arabic
                        else "Canceled the pending dangerous operation."
                    ),
                    "metadata": {
                        "capability_id": pending.get("capability_id", "unknown"),
                        "risk_level": pending.get("risk_level", "destructive"),
                        "facts": {"status": "canceled"},
                    },
                }
            else:
                return {
                    "type": "result",
                    "content": (
                        "لدي عملية خطرة بانتظار التأكيد. اكتب 'نعم' للتنفيذ أو 'إلغاء' للإلغاء."
                        if arabic
                        else "A dangerous operation is pending. Reply 'yes' to execute or 'cancel' to abort."
                    ),
                    "metadata": {
                        "capability_id": pending.get("capability_id", "unknown"),
                        "risk_level": pending.get("risk_level", "destructive"),
                        "facts": {"status": "awaiting_confirmation"},
                    },
                }
        else:
            resolved = resolve_windows_intent(text)
            if not resolved.matched:
                return None
            resolution = {
                "capability_id": resolved.capability_id,
                "action": resolved.action,
                "params": dict(resolved.params or {}),
                "risk_level": resolved.risk_level,
                "unsupported": resolved.unsupported,
                "unsupported_reason": resolved.unsupported_reason,
            }

        if bool(resolution.get("unsupported")):
            return {
                "type": "result",
                "content": str(
                    resolution.get("unsupported_reason")
                    or "Requested capability is not supported yet."
                ),
                "metadata": {
                    "capability_id": resolution.get("capability_id", "unknown"),
                    "risk_level": resolution.get("risk_level", "safe"),
                    "facts": {"status": "unsupported"},
                },
            }

        risk_level = str(resolution.get("risk_level", "safe"))
        if risk_level == "destructive" and not is_confirmation_message(text):
            self._pending_dangerous_intent = resolution
            return {
                "type": "result",
                "content": (
                    "هذا أمر خطِر. للتأكيد اكتب: نعم. للإلغاء اكتب: إلغاء."
                    if arabic
                    else "This is a destructive command. Reply 'yes' to confirm or 'cancel' to abort."
                ),
                "metadata": {
                    "capability_id": resolution.get("capability_id", "unknown"),
                    "risk_level": "destructive",
                    "facts": {"status": "awaiting_confirmation"},
                    "resolved_action": resolution.get("action"),
                    "resolved_params": resolution.get("params", {}),
                },
            }

        action = str(resolution.get("action", "")).strip()
        params = resolution.get("params", {})
        if not action:
            return None

        from Mudabbir.tools.builtin.desktop import DesktopTool

        raw = await DesktopTool().execute(action=action, **(params if isinstance(params, dict) else {}))
        if isinstance(raw, str) and raw.lower().startswith("error:"):
            return {
                "type": "result",
                "content": raw,
                "metadata": {
                    "capability_id": resolution.get("capability_id", "unknown"),
                    "risk_level": risk_level,
                    "facts": {"ok": False},
                    "resolved_action": action,
                    "resolved_params": params,
                },
            }
        parsed: Any = raw
        if isinstance(raw, str):
            try:
                parsed = json.loads(raw)
            except Exception:
                parsed = raw
        return {
            "type": "result",
            "content": f"Executed: {action}",
            "metadata": {
                "capability_id": resolution.get("capability_id", "unknown"),
                "risk_level": risk_level,
                "facts": {"ok": True},
                "resolved_action": action,
                "resolved_params": params,
                "result": parsed,
            },
        }

    async def _try_ai_desktop_response(
        self, message: str, history: list[dict] | None = None
    ) -> dict | None:
        """AI-first desktop control path: parse intent -> execute desktop actions -> reply."""
        if not bool(getattr(self.settings, "oi_ai_desktop_planner", True)):
            return None

        plan = await self._llm_plan_desktop_actions(message=message, history=history)
        if not isinstance(plan, dict):
            return None

        should_execute = bool(plan.get("should_execute", False))
        if not should_execute:
            return {"type": "ai_noop", "content": "", "metadata": {"plan": plan}}

        actions = plan.get("actions", [])
        if not isinstance(actions, list) or not actions:
            reason = str(plan.get("reason", "") or "").strip()
            if not reason:
                reason = "Could not determine executable desktop actions."
            return {
                "type": "result",
                "content": reason,
                "metadata": {
                    "capability_id": "desktop.plan",
                    "risk_level": "safe",
                    "facts": {"reason": reason, "should_execute": False},
                    "plan": plan,
                },
            }

        results = await self._execute_ai_desktop_actions(actions)
        if not results:
            return None
        ok_steps = [r for r in results if isinstance(r, dict) and r.get("ok")]
        fail_steps = [r for r in results if isinstance(r, dict) and not r.get("ok")]
        fallback_lines: list[str] = []
        if ok_steps:
            fallback_lines.append(
                "Executed actions: " + ", ".join(str(step.get("action", "")) for step in ok_steps)
            )
        if fail_steps:
            fallback_lines.append(
                "Failed actions: "
                + ", ".join(
                    f"{step.get('action', 'unknown')} ({step.get('error', 'error')})"
                    for step in fail_steps
                )
            )
        fallback_text = "\n".join(fallback_lines).strip() or "Desktop actions processed."
        return {
            "type": "result",
            "content": fallback_text,
            "metadata": {
                "capability_id": "desktop.batch",
                "risk_level": "safe",
                "facts": {
                    "success_count": len(ok_steps),
                    "failure_count": len(fail_steps),
                },
                "plan": plan,
                "results": results,
            },
        }

    async def _try_direct_desktop_response(
        self, message: str, history: list[dict] | None = None
    ) -> dict | None:
        """Deterministic handling for frequent desktop commands in Arabic/English."""
        text = message or ""
        lowered = text.lower()
        normalized = _normalize_text_for_match(text)
        arabic = _contains_arabic(text)
        history_items = history or []
        history_text = " ".join(
            str(item.get("content", ""))
            for item in history_items
            if isinstance(item, dict) and item.get("content")
        )
        history_normalized = _normalize_text_for_match(history_text)

        def has_any(haystack: str, needles: tuple[str, ...]) -> bool:
            return any(n in haystack for n in needles)

        def extract_first_number() -> int | None:
            m = re.search(r"-?\d+", text)
            if not m:
                return None
            try:
                return int(m.group(0))
            except Exception:
                return None

        def extract_first_two_numbers() -> tuple[int, int] | None:
            nums = re.findall(r"-?\d+", text)
            if len(nums) < 2:
                return None
            try:
                return int(nums[0]), int(nums[1])
            except Exception:
                return None

        def extract_all_numbers() -> list[int]:
            values: list[int] = []
            for token in re.findall(r"-?\d+", text):
                try:
                    values.append(int(token))
                except Exception:
                    continue
            return values

        def parse_compound_keys(raw: str) -> list[str]:
            if not raw:
                return []
            alias_map = {
                "control": "ctrl",
                "ctl": "ctrl",
                "كنترول": "ctrl",
                "كترل": "ctrl",
                "ctrl": "ctrl",
                "shift": "shift",
                "شفت": "shift",
                "alt": "alt",
                "الت": "alt",
                "win": "win",
                "windows": "win",
                "window": "win",
                "زر الويندوز": "win",
                "ويندوز": "win",
                "cmd": "win",
                "command": "win",
                "esc": "esc",
                "escape": "esc",
                "زر الهروب": "esc",
                "tab": "tab",
                "enter": "enter",
                "انتر": "enter",
                "return": "enter",
                "space": "space",
                "مسافة": "space",
                "delete": "delete",
                "del": "delete",
                "backspace": "backspace",
                "prtsc": "printscreen",
                "printscreen": "printscreen",
                "print": "printscreen",
            }
            raw_norm = _normalize_text_for_match(raw)
            parts = re.findall(
                r"(?:ctrl|control|ctl|كنترول|كترل|shift|شفت|alt|الت|win|windows|window|cmd|command|"
                r"esc|escape|tab|enter|انتر|return|space|مسافة|delete|del|backspace|prtsc|printscreen|"
                r"f(?:[1-9]|1[0-2]))",
                raw_norm,
            )
            keys: list[str] = []
            for part in parts:
                if part in alias_map:
                    keys.append(alias_map[part])
                elif re.fullmatch(r"f([1-9]|1[0-2])", part):
                    keys.append(part)
            deduped: list[str] = []
            for key in keys:
                if key not in deduped:
                    deduped.append(key)
            return deduped

        def clean_query(candidate: str) -> str:
            raw = _normalize_text_for_match(candidate or "")
            raw = re.sub(r"[\"'`]", " ", raw)
            raw = re.sub(r"[،,:;!?()\[\]{}]", " ", raw)
            words = [w for w in re.split(r"\s+", raw) if w]
            stop_words = {
                "افتح",
                "افتحي",
                "فتح",
                "شغل",
                "شغللي",
                "شغللي",
                "شغل",
                "launch",
                "open",
                "search",
                "show",
                "list",
                "ابحث",
                "دور",
                "اعرض",
                "عرض",
                "اغلق",
                "سكر",
                "close",
                "app",
                "apps",
                "application",
                "applications",
                "program",
                "programs",
                "تطبيق",
                "التطبيق",
                "تطبيقات",
                "التطبيقات",
                "برنامج",
                "البرامج",
                "المثبته",
                "المثبتة",
                "المفتوحه",
                "المفتوحة",
                "القائمه",
                "قائمه",
                "قائمة",
                "القايمه",
                "قايمه",
                "ابدا",
                "ابدأ",
                "menu",
                "start",
                "file",
                "files",
                "ملف",
                "ملفات",
                "عن",
                "for",
                "in",
                "في",
                "الى",
                "الي",
                "على",
                "تبع",
                "تبعها",
                "تبعه",
                "please",
                "plz",
                "and",
                "طيب",
                "لو",
                "و",
                "ثم",
                "ابدا",
                "ابدأ",
                "قاعده",
                "قائمه",
                "بحث",
            }
            filtered: list[str] = []
            for w in words:
                if w in stop_words:
                    continue
                # Handle Arabic conjunction prefix (e.g. "وابحث", "وافتح").
                if w.startswith("و") and len(w) > 1 and w[1:] in stop_words:
                    continue
                filtered.append(w)
            return " ".join(filtered).strip()

        def extract_trailing_query() -> str:
            return clean_query(text)

        def extract_search_query() -> str:
            m = re.search(r"(?i)(?:عن|for)\s+(.+)$", text)
            if m and m.group(1).strip():
                q = clean_query(m.group(1))
            else:
                q = extract_trailing_query()
            q = re.sub(r"(?i)^(?:file|files|ملف|ملفات)\s+", "", q).strip()
            return q

        def extract_desktop_file_query(raw_text: str) -> str:
            candidate = ""
            m = re.search(
                r"(?:file|files|ملف|ملفات)\s+(.+?)(?:\s+(?:على|on)\s+(?:سطح المكتب|desktop).*)?$",
                raw_text,
                re.IGNORECASE,
            )
            if m:
                candidate = m.group(1) or ""
            if not candidate:
                m2 = re.search(r"(?:على|on)\s+(.+?)(?:\s+(?:على|on)\s+(?:سطح المكتب|desktop).*)?$", raw_text, re.IGNORECASE)
                if m2:
                    candidate = m2.group(1) or ""
            if not candidate:
                candidate = raw_text
            cleaned = clean_query(candidate)
            cleaned = re.sub(
                r"(?i)\b(?:file|files|ملف|ملفات|desktop|سطح|المكتب|mouse|cursor|ماوس|المؤشر|move|حرك|حرّك)\b",
                " ",
                cleaned,
            )
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            return cleaned

        def extract_topic_level(topic_tokens: tuple[str, ...]) -> int | None:
            for token in topic_tokens:
                token_re = re.escape(token)
                m = re.search(rf"{token_re}[^\d-]{{0,20}}(-?\d{{1,3}})", normalized, re.IGNORECASE)
                if m:
                    try:
                        return int(m.group(1))
                    except Exception:
                        pass
                m = re.search(rf"(-?\d{{1,3}})[^\d-]{{0,20}}{token_re}", normalized, re.IGNORECASE)
                if m:
                    try:
                        return int(m.group(1))
                    except Exception:
                        pass
            return None

        def extract_ui_target(raw_text: str) -> tuple[str, str]:
            """Extract (control_name, window_name) from natural desktop phrases."""
            window_name = ""
            left_part = raw_text or ""
            window_match = re.search(r"(?:\s+(?:في|داخل|in)\s+)(.+)$", raw_text, re.IGNORECASE)
            if window_match:
                window_name = clean_query(window_match.group(1) or "")
                left_part = raw_text[: window_match.start()]

            control_name = ""
            control_match = re.search(
                r"(?:زر|button|control|element|label|icon)\s+(.+)$",
                left_part,
                re.IGNORECASE,
            )
            if control_match:
                control_name = clean_query(control_match.group(1) or "")

            if not control_name:
                target_match = re.search(r"(?:على|لعند|to|towards)\s+(.+)$", left_part, re.IGNORECASE)
                if target_match:
                    control_name = clean_query(target_match.group(1) or "")

            if not control_name:
                control_name = clean_query(left_part)

            control_name = re.sub(
                r"(?i)^(?:الزر|زر|button|control|element|label|icon)\s+",
                "",
                control_name,
            ).strip()
            control_name = re.sub(
                r"(?i)\b(?:mouse|cursor|ماوس|الماوس|المؤشر|move|حرك|حرّك)\b",
                " ",
                control_name,
            )
            control_name = re.sub(r"\s+", " ", control_name).strip()
            return control_name, window_name

        def mentions_ui_target_context(raw_text: str) -> bool:
            """Detect targetable UI element mention without rigid keyword tuples."""
            candidate = _normalize_text_for_match(str(raw_text or ""))
            if not candidate:
                return False
            if re.search(
                r"(?:زر|button|control|element|label|icon|ايقون|حقل|input|textbox|field)\b",
                candidate,
                re.IGNORECASE,
            ):
                return True
            if re.search(r"(?:على|لعند|to|towards|on)\s+\S+", candidate, re.IGNORECASE):
                return True
            return False

        def infer_recent_ui_target() -> tuple[str, str]:
            """Try to recover last UI target/window from recent conversation context."""
            for item in reversed(history_items):
                if not isinstance(item, dict):
                    continue
                content = str(item.get("content", "") or "").strip()
                if not content:
                    continue
                # assistant messages that include resolved targets
                m_btn = re.search(r"(?i)(?:clicked button|moved mouse to ui element)\s*:\s*(.+)$", content)
                if m_btn:
                    return clean_query(m_btn.group(1) or ""), ""
                m_ar = re.search(r"(?:تم الضغط على الزر|تم تحريك الماوس إلى العنصر)\s*:\s*(.+)$", content)
                if m_ar:
                    return clean_query(m_ar.group(1) or ""), ""
                m_win = re.search(r"(?i)window\s*:\s*(.+)$", content)
                if m_win:
                    return "", clean_query(m_win.group(1) or "")
                m_ar_win = re.search(r"(?:ضمن النافذة)\s*:\s*(.+)$", content)
                if m_ar_win:
                    return "", clean_query(m_ar_win.group(1) or "")
            return "", ""

        def is_pronoun_target(value: str) -> bool:
            v = _normalize_text_for_match(str(value or "")).strip()
            if not v:
                return True
            pronouns = {
                "عليه",
                "عليها",
                "عليهم",
                "عليه؟",
                "عليها؟",
                "it",
                "this",
                "that",
                "there",
                "here",
                "same",
                "نفسه",
                "نفسها",
                "نفس",
            }
            return v in pronouns

        def parse_tool_json(raw: str) -> tuple[dict | list | None, str | None]:
            if not isinstance(raw, str):
                return None, "Unexpected tool response type."
            if raw.lower().startswith("error:"):
                return None, raw.replace("Error: ", "", 1)
            try:
                return json.loads(raw), None
            except Exception as e:
                return None, f"Invalid tool response: {e}"

        def normalize_brightness_error(err: str) -> str:
            msg = str(err or "").strip()
            lowered_err = msg.lower()
            if "no displays detected" in lowered_err:
                return (
                    "لا يمكن التحكم بسطوع الشاشة برمجيًا على هذا الجهاز/الجلسة الحالية."
                    if arabic
                    else "Programmatic brightness control is unavailable on this device/session."
                )
            if "wmi" in lowered_err and "failed" in lowered_err:
                return (
                    "فشل التحكم بالسطوع عبر WMI. قد تحتاج تشغيل التطبيق بصلاحيات أعلى."
                    if arabic
                    else "Brightness control via WMI failed. Try running with elevated privileges."
                )
            return msg

        def extract_typed_text(raw_text: str) -> str:
            quoted = re.search(r"[\"“](.+?)[\"”]|'(.+?)'", raw_text or "")
            if quoted:
                return (quoted.group(1) or quoted.group(2) or "").strip()

            m = re.search(r"(?:اكتب|ادخل|type|write)\s+(.+)$", raw_text or "", re.IGNORECASE)
            if not m:
                return ""
            candidate = (m.group(1) or "").strip()
            if not candidate:
                return ""

            candidate = re.sub(
                r"(?:\s+(?:then|and then|ثم|وبعدين)\s*(?:press|اضغط|اكبس)\s*(?:enter|انتر).*)$",
                "",
                candidate,
                flags=re.IGNORECASE,
            )
            candidate = re.sub(
                r"(?:\s+(?:press|اضغط|اكبس)\s*(?:enter|انتر).*)$",
                "",
                candidate,
                flags=re.IGNORECASE,
            )
            candidate = re.sub(
                r"(?:\s+(?:في|داخل|in)\s+(?:حقل|مربع|textbox|field|input|نافذه|نافذة|window|app).*)$",
                "",
                candidate,
                flags=re.IGNORECASE,
            )
            candidate = re.sub(
                r"(?:\s+(?:على|بال|ب|with)\s+(?:الكيبورد|كيبورد|لوحة المفاتيح|keyboard).*)$",
                "",
                candidate,
                flags=re.IGNORECASE,
            )
            return candidate.strip().strip(".")

        def format_seconds_human(seconds_value: Any) -> str:
            try:
                total = int(seconds_value)
            except Exception:
                return ""
            if total <= 0:
                return ""
            hours = total // 3600
            minutes = (total % 3600) // 60
            if arabic:
                if hours and minutes:
                    return f"{hours} ساعة و{minutes} دقيقة"
                if hours:
                    return f"{hours} ساعة"
                if minutes:
                    return f"{minutes} دقيقة"
                return "أقل من دقيقة"
            if hours and minutes:
                return f"{hours}h {minutes}m"
            if hours:
                return f"{hours}h"
            if minutes:
                return f"{minutes}m"
            return "<1m"

        volume_topic_tokens = ("الصوت", "صوت", "volume", "vol", "الميديا", "ميديا")
        brightness_topic_tokens = ("سطوع", "brightness", "الاضاءه", "الاضاءة", "اضاءه", "اضاءة")
        battery_topic_tokens = (
            "battery",
            "بطاري",
            "بطارية",
            "شحن",
            "الشحن",
            "charging",
            "charge",
            "power level",
            "battery level",
        )
        volume_intent_tokens = (
            "اكتم",
            "كتم",
            "اسكت",
            "سكت",
            "unmute",
            "mute",
            "الغ الكتم",
            "الغي الكتم",
            "فك كتم",
            "فك الكتم",
            "رجع الصوت",
            "شغل الصوت",
            "اعمله ماكس",
            "ماكس",
            "اقصى",
            "أقصى",
            "على الاخر",
            "على الآخر",
            "للاخر",
            "للآخر",
            "صفر",
            "للصفر",
        )
        window_topic_tokens = ("نافذه", "نافذة", "windows", "window", "شاشه", "شاشة")

        def infer_recent_topic() -> str | None:
            for item in reversed(history_items):
                if not isinstance(item, dict):
                    continue
                content = _normalize_text_for_match(str(item.get("content", "") or ""))
                if not content:
                    continue
                if has_any(content, volume_topic_tokens):
                    return "volume"
                if has_any(content, brightness_topic_tokens):
                    return "brightness"
                if has_any(content, window_topic_tokens):
                    return "windows"
                if has_any(content, ("whatsapp", "واتساب", "وتساب", "واتس")):
                    return "whatsapp"
            return None

        recent_topic = infer_recent_topic()

        volume_intent_without_topic = has_any(normalized, volume_intent_tokens) and not has_any(
            normalized, brightness_topic_tokens
        )
        has_volume_topic = has_any(normalized, volume_topic_tokens)
        has_brightness_topic = has_any(normalized, brightness_topic_tokens)
        is_volume_request = has_volume_topic or (
            recent_topic == "volume"
            and (extract_first_number() is not None or volume_intent_without_topic)
        ) or volume_intent_without_topic

        # Combined audio+brightness handling (status/set in one request).
        if has_volume_topic and has_brightness_topic:
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                query_words = (
                    "كم",
                    "قديش",
                    "الحالي",
                    "current",
                    "status",
                    "ما هو",
                    "شو",
                    "الان",
                    "now",
                    "هلا",
                    "هلق",
                )
                set_words = (
                    "اجعل",
                    "خلي",
                    "خليه",
                    "خليها",
                    "اعمل",
                    "اعمله",
                    "اضبط",
                    "set",
                    "الى",
                    "to",
                    "%",
                    "بالميه",
                    "بالمئة",
                    "كمان",
                    "also",
                )
                up_words = ("ارفع", "زيد", "علي", "increase", "up")
                down_words = ("نزل", "اخفض", "خفض", "وطي", "decrease", "down")
                max_words = ("ماكس", "اقصى", "أقصى", "maximum", "max", "على الاخر", "على الآخر")
                min_words = ("صفر", "اقل", "minimum", "min", "للصفر", "على الصفر")

                # Status query for both in one response.
                if has_any(normalized, query_words) and not has_any(
                    normalized, set_words + up_words + down_words + max_words + min_words
                ):
                    vol_raw = await DesktopTool().execute(action="volume", mode="get")
                    bri_raw = await DesktopTool().execute(action="brightness", mode="get")
                    vol_data, vol_err = parse_tool_json(vol_raw)
                    bri_data, bri_err = parse_tool_json(bri_raw)

                    if vol_err and bri_err:
                        merged = f"volume={vol_err}; brightness={normalize_brightness_error(bri_err)}"
                        return {"type": "message", "content": merged}

                    vol_pct = (vol_data or {}).get("level_percent") if isinstance(vol_data, dict) else None
                    vol_muted = (vol_data or {}).get("muted") if isinstance(vol_data, dict) else None
                    bri_pct = (
                        (bri_data or {}).get("brightness_percent") if isinstance(bri_data, dict) else None
                    )

                    if arabic:
                        lines = []
                        if vol_err:
                            lines.append(f"الصوت: فشل القراءة ({vol_err})")
                        else:
                            lines.append(f"الصوت: {vol_pct}% (كتم: {'نعم' if vol_muted else 'لا'})")
                        if bri_err:
                            lines.append(f"السطوع: فشل القراءة ({normalize_brightness_error(bri_err)})")
                        else:
                            lines.append(f"السطوع: {bri_pct}%")
                        return {"type": "message", "content": "\n".join(lines)}

                    lines = []
                    if vol_err:
                        lines.append(f"Volume: read failed ({vol_err})")
                    else:
                        lines.append(f"Volume: {vol_pct}% (muted={vol_muted})")
                    if bri_err:
                        lines.append(
                            f"Brightness: read failed ({normalize_brightness_error(bri_err)})"
                        )
                    else:
                        lines.append(f"Brightness: {bri_pct}%")
                    return {"type": "message", "content": "\n".join(lines)}

                numbers = extract_all_numbers()
                vol_level = extract_topic_level(volume_topic_tokens)
                bri_level = extract_topic_level(brightness_topic_tokens)
                if vol_level is None and numbers:
                    vol_level = numbers[0]
                if bri_level is None:
                    if len(numbers) >= 2:
                        bri_level = numbers[1]
                    elif vol_level is not None:
                        bri_level = vol_level

                vol_raw = bri_raw = ""
                vol_err = bri_err = None
                vol_data = bri_data = None

                if has_any(normalized, max_words):
                    vol_raw = await DesktopTool().execute(action="volume", mode="max")
                    bri_raw = await DesktopTool().execute(action="brightness", mode="max")
                elif has_any(normalized, min_words):
                    vol_raw = await DesktopTool().execute(action="volume", mode="set", level=0)
                    bri_raw = await DesktopTool().execute(action="brightness", mode="min")
                elif has_any(normalized, up_words):
                    delta = max(1, min(100, abs(numbers[0]))) if numbers else 10
                    vol_raw = await DesktopTool().execute(action="volume", mode="up", delta=delta)
                    bri_raw = await DesktopTool().execute(action="brightness", mode="up", delta=delta)
                elif has_any(normalized, down_words):
                    delta = max(1, min(100, abs(numbers[0]))) if numbers else 10
                    vol_raw = await DesktopTool().execute(action="volume", mode="down", delta=delta)
                    bri_raw = await DesktopTool().execute(action="brightness", mode="down", delta=delta)
                elif vol_level is not None or bri_level is not None:
                    if vol_level is None:
                        vol_level = bri_level
                    if bri_level is None:
                        bri_level = vol_level
                    vol_level = max(0, min(100, int(vol_level or 0)))
                    bri_level = max(0, min(100, int(bri_level or 0)))
                    vol_raw = await DesktopTool().execute(action="volume", mode="set", level=vol_level)
                    bri_raw = await DesktopTool().execute(action="brightness", mode="set", level=bri_level)

                if vol_raw or bri_raw:
                    vol_data, vol_err = parse_tool_json(vol_raw) if vol_raw else (None, "not updated")
                    bri_data, bri_err = parse_tool_json(bri_raw) if bri_raw else (None, "not updated")

                    vol_pct = (vol_data or {}).get("level_percent") if isinstance(vol_data, dict) else None
                    bri_pct = (
                        (bri_data or {}).get("brightness_percent") if isinstance(bri_data, dict) else None
                    )

                    if arabic:
                        lines = []
                        if vol_err:
                            lines.append(f"الصوت: فشل التعديل ({vol_err})")
                        else:
                            lines.append(f"الصوت: تم ضبطه على {vol_pct}%")
                        if bri_err:
                            lines.append(
                                f"السطوع: فشل التعديل ({normalize_brightness_error(bri_err)})"
                            )
                        else:
                            lines.append(f"السطوع: تم ضبطه على {bri_pct}%")
                        return {"type": "message", "content": "\n".join(lines)}

                    lines = []
                    if vol_err:
                        lines.append(f"Volume: update failed ({vol_err})")
                    else:
                        lines.append(f"Volume set to {vol_pct}%")
                    if bri_err:
                        lines.append(
                            f"Brightness: update failed ({normalize_brightness_error(bri_err)})"
                        )
                    else:
                        lines.append(f"Brightness set to {bri_pct}%")
                    return {"type": "message", "content": "\n".join(lines)}
            except Exception:
                pass

        # Battery status queries.
        battery_query_tokens = (
            "كم",
            "قديش",
            "نسبه",
            "نسبة",
            "status",
            "state",
            "level",
            "percent",
            "percentage",
            "remaining",
            "متبقي",
            "المتبقي",
            "حاله",
            "حالة",
            "يشحن",
            "موصول",
            "plugged",
            "charging",
        )
        asks_battery_status = has_any(normalized, battery_topic_tokens) and (
            has_any(normalized, battery_query_tokens)
            or has_any(
                normalized,
                (
                    "battery?",
                    "battery ?",
                    "كم البطاريه",
                    "كم البطارية",
                    "شو البطاريه",
                    "شو البطارية",
                ),
            )
        )
        if asks_battery_status:
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="battery_status")
                data, err = parse_tool_json(raw)
                if err:
                    return {
                        "type": "message",
                        "content": f"فشل قراءة حالة البطارية: {err}"
                        if arabic
                        else f"Failed to read battery status: {err}",
                    }

                payload = data if isinstance(data, dict) else {}
                available = bool(payload.get("available", False))
                percent_raw = payload.get("percent")
                percent_val: float | None = None
                try:
                    if percent_raw is not None:
                        percent_val = float(percent_raw)
                except Exception:
                    percent_val = None

                if not available or percent_val is None:
                    return {
                        "type": "message",
                        "content": "ما قدرت أوصل لحالة البطارية على هذا الجهاز حالياً."
                        if arabic
                        else "Battery status is not available on this machine/session right now.",
                    }

                percent_val = max(0.0, min(100.0, percent_val))
                rounded = round(percent_val)
                percent_text = f"{int(rounded)}%" if abs(percent_val - rounded) < 0.05 else f"{percent_val:.1f}%"
                plugged_raw = payload.get("plugged")
                plugged: bool | None
                if isinstance(plugged_raw, bool):
                    plugged = plugged_raw
                else:
                    plugged = None
                eta_text = format_seconds_human(payload.get("secs_left"))

                if arabic:
                    parts = [f"نسبة البطارية حالياً: {percent_text}."]
                    if plugged is True:
                        parts.append("الجهاز موصول بالشاحن.")
                    elif plugged is False:
                        parts.append("الجهاز يعمل على البطارية.")
                        if eta_text:
                            parts.append(f"الوقت المتبقي تقريباً: {eta_text}.")
                    return {"type": "message", "content": " ".join(parts)}

                parts = [f"Battery level: {percent_text}."]
                if plugged is True:
                    parts.append("The device is plugged in.")
                elif plugged is False:
                    parts.append("The device is running on battery.")
                    if eta_text:
                        parts.append(f"Estimated remaining time: {eta_text}.")
                return {"type": "message", "content": " ".join(parts)}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل قراءة البطارية: {e}" if arabic else f"Battery status failed: {e}",
                }

        async def take_snapshot_flow(open_after: bool) -> dict:
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="screen_snapshot")
                if isinstance(raw, str) and raw.lower().startswith("error:"):
                    msg = raw.replace("Error: ", "", 1)
                    return {
                        "type": "message",
                        "content": f"فشل أخذ لقطة الشاشة: {msg}" if arabic else f"Failed to capture screenshot: {msg}",
                    }
                path_match = re.search(r"saved to\s+(.+)$", str(raw), re.IGNORECASE)
                snap_path = path_match.group(1).strip() if path_match else ""
                opened = False
                if open_after and snap_path:
                    try:
                        os.startfile(snap_path)  # type: ignore[attr-defined]
                        opened = True
                    except Exception:
                        opened = False
                if arabic:
                    return {
                        "type": "message",
                        "content": f"تم حفظ لقطة الشاشة{' وفتحها' if opened else ''}."
                        + (f"\nالمسار: {snap_path}" if snap_path else ""),
                    }
                return {
                    "type": "message",
                    "content": f"Screenshot saved{' and opened' if opened else ''}."
                    + (f"\nPath: {snap_path}" if snap_path else ""),
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تنفيذ لقطة الشاشة: {e}" if arabic else f"Screenshot flow failed: {e}",
                }

        # Open WhatsApp app / web.
        whatsapp_tokens = ("whatsapp", "واتساب", "وتساب", "الوتساب", "واتس", "الوتس", "وتس")
        open_tokens = ("open", "launch", "افتح", "شغل", "فتح")
        if has_any(normalized, whatsapp_tokens) and has_any(normalized, open_tokens):
            if "web" in normalized or "واتساب ويب" in normalized:
                try:
                    subprocess.Popen(["cmd", "/c", "start", "", "https://web.whatsapp.com"])
                    return {
                        "type": "message",
                        "content": "تم فتح WhatsApp Web في المتصفح." if arabic else "Opened WhatsApp Web in browser.",
                    }
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل فتح واتساب ويب: {e}" if arabic else f"Failed to open WhatsApp Web: {e}",
                    }

            ok, info = self._open_whatsapp_app()
            if ok:
                if arabic:
                    return {"type": "message", "content": f"تم فتح واتساب بنجاح.\nالمسار: {info}"}
                return {"type": "message", "content": f"WhatsApp opened successfully.\nPath: {info}"}
            if arabic:
                return {"type": "message", "content": f"تعذر فتح واتساب.\nالسبب: {info}"}
            return {"type": "message", "content": f"Failed to open WhatsApp.\nReason: {info}"}

        # Close WhatsApp quickly without requiring "app/program" words.
        if has_any(normalized, ("اغلق", "سكر", "اقفل", "close", "kill")) and has_any(
            normalized, whatsapp_tokens
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="close_app", process_name="whatsapp", force=True)
                if isinstance(raw, str) and raw.lower().startswith("error:"):
                    reason = raw.replace("Error: ", "", 1)
                    return {
                        "type": "message",
                        "content": f"فشل إغلاق الواتساب: {reason}" if arabic else f"Failed to close WhatsApp: {reason}",
                    }
                closed = 0
                try:
                    payload = json.loads(raw)
                    closed = int(payload.get("closed", 0))
                except Exception:
                    pass
                return {
                    "type": "message",
                    "content": f"تم إغلاق الواتساب (عمليات مغلقة: {closed})."
                    if arabic
                    else f"WhatsApp closed (closed processes: {closed}).",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل إغلاق الواتساب: {e}" if arabic else f"Failed to close WhatsApp: {e}",
                }

        # Enter top pinned chat in WhatsApp (best-effort click inside WhatsApp sidebar).
        if has_any(
            normalized,
            (
                "المحادثه المثبته",
                "المحادثة المثبتة",
                "المثبتة فوق",
                "pinned chat",
                "top pinned",
            ),
        ):
            try:
                import pyautogui
                import pygetwindow as gw

                pyautogui.FAILSAFE = False
                wins = [
                    w
                    for w in gw.getAllWindows()
                    if "whatsapp" in str(getattr(w, "title", "") or "").lower()
                ]
                if not wins:
                    return {
                        "type": "message",
                        "content": "لم أجد نافذة واتساب مفتوحة. افتح واتساب أولًا."
                        if arabic
                        else "No WhatsApp window found. Open WhatsApp first.",
                    }
                win = wins[0]
                try:
                    if bool(getattr(win, "isMinimized", False)):
                        win.restore()
                except Exception:
                    pass
                try:
                    win.activate()
                except Exception:
                    pass
                left = int(getattr(win, "left", 0))
                top = int(getattr(win, "top", 0))
                width = int(getattr(win, "width", 1000))
                height = int(getattr(win, "height", 700))
                x = left + int(max(140, min(260, width * 0.18)))
                y = top + int(max(130, min(220, height * 0.22)))
                pyautogui.click(x=x, y=y)
                return {
                    "type": "message",
                    "content": "تم الدخول للمحادثة المثبتة (تقريبياً)."
                    if arabic
                    else "Clicked the likely top pinned chat.",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل الدخول للمحادثة المثبتة: {e}" if arabic else f"Failed to open pinned chat: {e}",
                }

        # Open Windows settings and specific sections (default apps, bluetooth, etc.).
        if (
            has_any(normalized, open_tokens + ("روح", "go", "to"))
            and has_any(normalized, ("الاعدادات", "الاعداد", "اعدادات", "settings", "ms-settings"))
        ) or has_any(
            normalized,
            (
                "default apps",
                "default app",
                "التطبيقات الافتراضيه",
                "التطبيقات الافتراضية",
                "البرامج الافتراضيه",
                "البرامج الافتراضية",
            ),
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                page = "settings"
                if has_any(normalized, ("default apps", "default app", "التطبيقات الافتراضيه", "التطبيقات الافتراضية")):
                    page = "default apps"
                elif has_any(normalized, ("bluetooth", "بلوتوث")):
                    page = "bluetooth"
                elif has_any(normalized, ("apps", "applications", "app", "تطبيقات", "برامج")):
                    page = "apps"
                elif has_any(normalized, ("display", "شاشه", "شاشة")):
                    page = "display"
                elif has_any(normalized, ("sound", "volume", "صوت")):
                    page = "sound"
                elif has_any(normalized, ("network", "wifi", "واي فاي", "شبكه", "شبكة")):
                    page = "network"
                elif has_any(normalized, ("language", "keyboard", "لغه", "لغة", "كيبورد")):
                    page = "language"
                elif has_any(normalized, ("privacy", "خصوصيه", "خصوصية")):
                    page = "privacy"
                elif has_any(normalized, ("update", "windows update", "تحديث")):
                    page = "update"

                raw = await DesktopTool().execute(action="open_settings_page", page=page)
                if str(raw).lower().startswith("error:"):
                    reason = str(raw).replace("Error: ", "", 1)
                    return {
                        "type": "message",
                        "content": f"فشل فتح الإعدادات: {reason}" if arabic else f"Failed to open Settings: {reason}",
                    }

                if arabic:
                    if page == "default apps":
                        return {"type": "message", "content": "تم فتح الإعدادات على قسم التطبيقات الافتراضية."}
                    if page == "bluetooth":
                        return {"type": "message", "content": "تم فتح إعدادات البلوتوث."}
                    return {"type": "message", "content": "تم فتح إعدادات ويندوز."}
                if page == "default apps":
                    return {"type": "message", "content": "Opened Settings at Default apps."}
                if page == "bluetooth":
                    return {"type": "message", "content": "Opened Bluetooth settings."}
                return {"type": "message", "content": "Opened Windows Settings."}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل فتح الإعدادات: {e}" if arabic else f"Failed to open Settings: {e}",
                }

        # Open browser / YouTube shortcuts.
        if has_any(normalized, open_tokens) and has_any(normalized, ("يوتيوب", "youtube")):
            try:
                subprocess.Popen(["cmd", "/c", "start", "", "https://www.youtube.com"], shell=False)
                return {
                    "type": "message",
                    "content": "تم فتح يوتيوب في المتصفح." if arabic else "Opened YouTube in browser.",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل فتح يوتيوب: {e}" if arabic else f"Failed to open YouTube: {e}",
                }

        if has_any(normalized, open_tokens) and has_any(
            normalized, ("المتصفح", "browser", "edge", "chrome", "firefox")
        ):
            try:
                subprocess.Popen(["cmd", "/c", "start", "", "https://www.google.com"], shell=False)
                return {
                    "type": "message",
                    "content": "تم فتح المتصفح." if arabic else "Opened browser.",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل فتح المتصفح: {e}" if arabic else f"Failed to open browser: {e}",
                }

        # Open Telegram (App or protocol).
        if has_any(normalized, open_tokens) and has_any(normalized, ("telegram", "تلجرام", "تيليجرام", "تليجرام")):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                result = await DesktopTool().execute(action="launch_start_app", query="Telegram")
                if not str(result).lower().startswith("error:"):
                    return {
                        "type": "message",
                        "content": "تم فتح تلجرام." if arabic else "Opened Telegram.",
                    }
            except Exception:
                pass
            try:
                subprocess.Popen(["cmd", "/c", "start", "", "telegram:"], shell=False)
                return {
                    "type": "message",
                    "content": "تم فتح تلجرام." if arabic else "Opened Telegram.",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل فتح تلجرام: {e}" if arabic else f"Failed to open Telegram: {e}",
                }

        # Generic open intent without requiring explicit "app/program" keyword.
        if has_any(normalized, open_tokens):
            avoid_tokens = (
                "youtube",
                "يوتيوب",
                "browser",
                "المتصفح",
                "file",
                "files",
                "ملف",
                "ملفات",
                "settings",
                "الاعدادات",
                "الإعدادات",
                "bluetooth",
                "بلوتوث",
                "whatsapp",
                "واتساب",
                "telegram",
                "تلجرام",
            )
            app_query = extract_trailing_query()
            if app_query and not has_any(normalized, avoid_tokens):
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    raw = await DesktopTool().execute(action="launch_start_app", query=app_query)
                    if not str(raw).lower().startswith("error:"):
                        return {
                            "type": "message",
                            "content": f"تم فتح {app_query}." if arabic else f"Opened {app_query}.",
                        }
                except Exception:
                    pass

        # Bluetooth control: on/off/toggle with fallback to settings.
        if has_any(normalized, ("bluetooth", "بلوتوث")) and has_any(
            normalized,
            ("open", "on", "enable", "افتح", "شغل", "فعّل", "فعل", "off", "disable", "اطفي", "طفي", "سكّر", "سكر"),
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                mode = "open_settings"
                if has_any(normalized, ("off", "disable", "اطفي", "طفي", "سكّر", "سكر")):
                    mode = "off"
                elif has_any(normalized, ("on", "enable", "شغل", "فعّل", "فعل")):
                    mode = "on"

                raw = await DesktopTool().execute(action="bluetooth_control", mode=mode)
                if str(raw).lower().startswith("error:"):
                    reason = str(raw).replace("Error: ", "", 1)
                    return {
                        "type": "message",
                        "content": f"فشل التحكم بالبلوتوث: {reason}" if arabic else f"Bluetooth control failed: {reason}",
                    }

                try:
                    payload = json.loads(raw)
                except Exception:
                    payload = {}
                if isinstance(payload, dict) and payload.get("ok") is False and payload.get("opened_settings"):
                    msg = "تعذر تبديل البلوتوث تلقائيًا، فتم فتح إعدادات البلوتوث." if arabic else "Could not toggle Bluetooth automatically, so Bluetooth settings were opened."
                    return {"type": "message", "content": msg}

                if mode == "off":
                    return {"type": "message", "content": "تم إيقاف البلوتوث." if arabic else "Bluetooth turned off."}
                if mode == "on":
                    return {"type": "message", "content": "تم تشغيل البلوتوث." if arabic else "Bluetooth turned on."}
                return {"type": "message", "content": "تم فتح إعدادات البلوتوث." if arabic else "Opened Bluetooth settings."}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل التحكم بالبلوتوث: {e}" if arabic else f"Bluetooth control failed: {e}",
                }

        # Screenshot save/open direct flow.
        if has_any(normalized, ("لقطه", "لقطة", "سكرين", "شاشه", "شاشة", "screenshot", "screen shot")) and has_any(
            normalized, ("احفظ", "خزن", "save", "افتح", "عرض", "open")
        ):
            return await take_snapshot_flow(
                open_after=has_any(normalized, ("افتح", "open", "عرض"))
            )

        # Contextual screenshot follow-up (e.g. "خليه يحفظها ويفتحها").
        if recent_topic == "windows" and has_any(normalized, ("يحفظها", "احفظها", "save", "خزن")) and has_any(
            normalized, ("يفتحها", "افتحها", "open", "عرض")
        ):
            return await take_snapshot_flow(open_after=True)

        # Generic app launch from Start menu search.
        if has_any(normalized, open_tokens) and has_any(
            normalized, ("تطبيق", "برنامج", "app", "application", "program")
        ):
            app_query = extract_trailing_query()
            if app_query and not has_any(app_query.lower(), whatsapp_tokens):
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    result = await DesktopTool().execute(
                        action="launch_start_app",
                        query=app_query,
                    )
                    if result.lower().startswith("error:"):
                        msg = result.replace("Error: ", "", 1)
                        return {
                            "type": "message",
                            "content": f"تعذر فتح التطبيق المطلوب: {msg}" if arabic else f"Failed to launch app: {msg}",
                        }
                    return {
                        "type": "message",
                        "content": f"تم فتح التطبيق: {app_query}" if arabic else f"Opened app: {app_query}",
                    }
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل فتح التطبيق: {e}" if arabic else f"Failed to launch app: {e}",
                    }

        # Start menu app search.
        if has_any(normalized, ("ابحث", "search", "دور", "find")) and has_any(
            normalized, ("قائمه ابدا", "قايمه ابدا", "قائمة ابدأ", "start menu", "start")
        ):
            query = extract_search_query()
            if query in {"قائمه", "قايمه", "ابدا", "ابدأ", "قائمه ابدا", "قايمه ابدا"}:
                query = ""
            if not query:
                try:
                    import pyautogui

                    pyautogui.FAILSAFE = False
                    pyautogui.press("win")
                except Exception:
                    pass
                return {
                    "type": "message",
                    "content": "فتحت قائمة ابدأ. اكتب اسم التطبيق الذي تريد البحث عنه."
                    if arabic
                    else "Opened Start menu. Tell me the app name to search for.",
                }
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="search_start_apps", query=query, max_results=10)
                data, err = parse_tool_json(raw)
                if err:
                    return {"type": "message", "content": err}
                data = data if isinstance(data, dict) else {}
                apps = data.get("apps", []) if isinstance(data, dict) else []
                if not apps:
                    return {
                        "type": "message",
                        "content": f"لم أجد تطبيقات مطابقة لـ: {query}" if arabic else f"No Start menu apps matched: {query}",
                    }
                lines = [f"{idx + 1}. {a.get('Name', '')}" for idx, a in enumerate(apps[:10])]
                header = f"نتائج البحث في قائمة ابدأ لـ '{query}':" if arabic else f"Start menu search results for '{query}':"
                return {"type": "message", "content": header + "\n" + "\n".join(lines)}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل البحث في قائمة ابدأ: {e}" if arabic else f"Start menu search failed: {e}",
                }

        # File search.
        if has_any(normalized, ("ابحث", "search")) and has_any(normalized, ("ملف", "ملفات", "file", "files")):
            query = extract_search_query()
            if query:
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    raw = await DesktopTool().execute(action="search_files", query=query, max_results=20)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": err}
                    data = data if isinstance(data, dict) else {}
                    items = data.get("results", []) if isinstance(data, dict) else []
                    if not items:
                        return {
                            "type": "message",
                            "content": f"لم أجد ملفات مطابقة لـ: {query}" if arabic else f"No files matched: {query}",
                        }
                    lines = [f"{idx + 1}. {p}" for idx, p in enumerate(items[:10])]
                    header = f"نتائج البحث عن الملفات '{query}':" if arabic else f"File search results for '{query}':"
                    return {"type": "message", "content": header + "\n" + "\n".join(lines)}
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل البحث عن الملفات: {e}" if arabic else f"File search failed: {e}",
                    }

        # Open first matching file from query.
        if has_any(normalized, open_tokens) and has_any(normalized, ("ملف", "file")):
            query = extract_search_query()
            if query:
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    raw = await DesktopTool().execute(action="search_files", query=query, max_results=1)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": err}
                    data = data if isinstance(data, dict) else {}
                    items = data.get("results", []) if isinstance(data, dict) else []
                    if not items:
                        return {
                            "type": "message",
                            "content": f"لم أجد ملفًا مطابقًا لـ: {query}" if arabic else f"No file matched: {query}",
                        }
                    target = str(items[0])
                    suffix = Path(target).suffix.lower()
                    if suffix in {".bat", ".cmd", ".ps1", ".vbs", ".js"}:
                        if arabic:
                            return {
                                "type": "message",
                                "content": f"تم العثور على الملف: {target}\nتشغيل ملفات السكربت معطّل من هذا المسار المباشر للسلامة.",
                            }
                        return {
                            "type": "message",
                            "content": f"Found file: {target}\nOpening script files is blocked in direct mode for safety.",
                        }
                    subprocess.Popen(["cmd", "/c", "start", "", target], shell=False)
                    return {
                        "type": "message",
                        "content": f"تم فتح الملف: {target}" if arabic else f"Opened file: {target}",
                    }
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل فتح الملف: {e}" if arabic else f"Failed to open file: {e}",
                    }

        # Installed apps listing/search.
        if has_any(normalized, ("التطبيقات المثبته", "التطبيقات المثبتة", "installed apps", "installed programs")):
            query = extract_trailing_query()
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="list_installed_apps", query=query, max_results=25)
                data, err = parse_tool_json(raw)
                if err:
                    return {"type": "message", "content": err}
                data = data if isinstance(data, dict) else {}
                apps = data.get("apps", []) if isinstance(data, dict) else []
                if not apps:
                    return {
                        "type": "message",
                        "content": "لم أجد تطبيقات مثبتة مطابقة." if arabic else "No matching installed apps found.",
                    }
                lines = [f"{idx + 1}. {a.get('DisplayName', '')}" for idx, a in enumerate(apps[:20])]
                title = "أبرز التطبيقات المثبتة:" if arabic else "Installed apps:"
                return {"type": "message", "content": title + "\n" + "\n".join(lines)}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل قراءة التطبيقات المثبتة: {e}" if arabic else f"Failed to list installed apps: {e}",
                }

        # WhatsApp "latest message/sender" requests must use image-based vision analysis,
        # not generated pyautogui scripts from the model.
        if has_any(normalized, ("واتساب", "whatsapp")) and has_any(
            normalized,
            ("اخر", "آخر", "latest", "last", "recent"),
        ) and has_any(
            normalized,
            ("رساله", "رسالة", "message", "messaged", "بعت", "بعث", "ارسل", "أرسل"),
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="vision_tools", mode="describe_screen")
                data, err = parse_tool_json(raw)
                if err:
                    return {
                        "type": "message",
                        "content": "ما قدرت أحلل الشاشة حالياً." if arabic else "I couldn't analyze the screen right now.",
                    }
                data = data if isinstance(data, dict) else {}
                source = str(data.get("source", "") or "")
                if source == "vision":
                    top_app = str(data.get("top_app", "") or "").strip()
                    summary = str(data.get("ui_summary", "") or "").strip()
                    if arabic:
                        msg = "حلّلت الشاشة بالصورة."
                        if top_app:
                            msg += f"\nالتطبيق الأوضح: {top_app}"
                        if summary:
                            msg += f"\n{summary}"
                        return {"type": "message", "content": msg}
                    msg = "I analyzed the screen from the image."
                    if top_app:
                        msg += f"\nTop app: {top_app}"
                    if summary:
                        msg += f"\n{summary}"
                    return {"type": "message", "content": msg}

                if arabic:
                    return {
                        "type": "message",
                        "content": (
                            "تحليل الصورة غير مفعّل حالياً، لذلك ما بقدر أحدد آخر رسالة واتساب بدقة.\n"
                            "فعّل مزوّد Vision (OpenAI/Gemini) من الإعدادات."
                        ),
                    }
                return {
                    "type": "message",
                    "content": (
                        "Image vision is not configured, so I can't reliably identify the latest WhatsApp message.\n"
                        "Enable an OpenAI/Gemini vision provider in settings."
                    ),
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تحليل الشاشة: {e}" if arabic else f"Screen analysis failed: {e}",
                }

        # Running apps/processes overview.
        if has_any(
            normalized,
            (
                "التطبيقات المفتوحه",
                "التطبيقات المفتوحة",
                "البرامج المفتوحه",
                "running apps",
                "running programs",
                "شو شغال",
                "شو مفتوح",
                "شو على الشاشه",
                "شو على الشاشة",
                "شو في عل شاشه",
                "شو في عل شاشة",
                "شو في على الشاشه",
                "شو في على الشاشة",
                "شو في عالشاشه",
                "شو في عالشاشة",
                "شو على الشاشة هلا",
                "شو في الشاشة حاليا",
                "شو في على الشاشة حاليا",
                "what is on screen",
                "whats on screen",
                "what is open",
                "whats open",
            ),
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="list_processes", only_windowed=True, max_results=15)
                data, err = parse_tool_json(raw)
                if err:
                    return {"type": "message", "content": err}
                data = data if isinstance(data, dict) else {}
                rows = data.get("processes", []) if isinstance(data, dict) else []
                if not rows:
                    return {
                        "type": "message",
                        "content": "لا توجد تطبيقات مرئية مفتوحة حالياً." if arabic else "No visible running apps found.",
                    }
                lines = []
                for idx, row in enumerate(rows[:10], start=1):
                    name = row.get("Name") or row.get("name") or ""
                    title = row.get("MainWindowTitle") or ""
                    pid = row.get("Id") or row.get("pid") or ""
                    lines.append(f"{idx}. {name} (PID {pid}) - {title}")
                title = "التطبيقات المفتوحة حالياً:" if arabic else "Currently open apps:"
                return {"type": "message", "content": title + "\n" + "\n".join(lines)}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل قراءة التطبيقات المفتوحة: {e}" if arabic else f"Failed to list running apps: {e}",
                }

        if has_any(normalized, ("شو هنن", "شو هن", "شو هم", "show them", "list them")) and (
            infer_recent_topic() == "windows" or has_any(history_normalized, window_topic_tokens)
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                raw = await DesktopTool().execute(action="list_processes", only_windowed=True, max_results=15)
                data, err = parse_tool_json(raw)
                if err:
                    return {"type": "message", "content": err}
                data = data if isinstance(data, dict) else {}
                rows = data.get("processes", []) if isinstance(data, dict) else []
                if not rows:
                    return {
                        "type": "message",
                        "content": "لا توجد تطبيقات مرئية مفتوحة حالياً." if arabic else "No visible running apps found.",
                    }
                lines = []
                for idx, row in enumerate(rows[:10], start=1):
                    name = row.get("Name") or row.get("name") or ""
                    title = row.get("MainWindowTitle") or ""
                    pid = row.get("Id") or row.get("pid") or ""
                    lines.append(f"{idx}. {name} (PID {pid}) - {title}")
                return {
                    "type": "message",
                    "content": ("هاي النوافذ المفتوحة حالياً:\n" if arabic else "Currently open windows:\n")
                    + "\n".join(lines),
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل عرض النوافذ المفتوحة: {e}" if arabic else f"Failed to list open windows: {e}",
                }

        # Close app/process.
        if has_any(normalized, ("اغلق", "سكر", "اقفل", "close", "kill")) and has_any(
            normalized, ("تطبيق", "برنامج", "process", "app")
        ):
            target = extract_trailing_query()
            if target:
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    raw = await DesktopTool().execute(action="close_app", process_name=target, force=True)
                    if raw.lower().startswith("error:"):
                        msg = raw.replace("Error: ", "", 1)
                        return {
                            "type": "message",
                            "content": f"فشل إغلاق التطبيق: {msg}" if arabic else f"Failed to close app: {msg}",
                        }
                    closed = 0
                    try:
                        data = json.loads(raw)
                        closed = int(data.get("closed", 0))
                    except Exception:
                        pass
                    return {
                        "type": "message",
                        "content": f"تم إغلاق {closed} عملية مطابقة لـ: {target}" if arabic else f"Closed {closed} process(es) matching: {target}",
                    }
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل إغلاق التطبيق: {e}" if arabic else f"Failed to close app: {e}",
                    }

        # Volume controls (get/set/max/min/mute/up/down) with context-aware phrasing.
        if is_volume_request:
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                if has_any(normalized, ("كم", "الحالي", "current", "status", "قديش", "ما هو", "صار", "هلا", "هلق")) and (
                    has_any(normalized, volume_topic_tokens) or recent_topic == "volume"
                ):
                    raw = await DesktopTool().execute(action="volume", mode="get")
                    data, err = parse_tool_json(raw)
                    if err:
                        return {
                            "type": "message",
                            "content": f"فشل قراءة مستوى الصوت: {err}" if arabic else err,
                        }
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    muted = data.get("muted")
                    if arabic:
                        return {"type": "message", "content": f"مستوى الصوت الحالي: {lvl}%\nكتم الصوت: {'نعم' if muted else 'لا'}"}
                    return {"type": "message", "content": f"Current volume: {lvl}% (muted={muted})"}

                if has_any(normalized, ("صفر", "0", "للصفر", "على الصفر")) and has_any(
                    normalized, ("اجعل", "خلي", "set", "اضبط", "نزل", "خفض", "اخفض", "وطي")
                ):
                    raw = await DesktopTool().execute(action="volume", mode="set", level=0)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل ضبط الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم ضبط الصوت على {lvl}%." if arabic else f"Volume set to {lvl}%."}

                if has_any(
                    normalized,
                    ("ماكس", "اقصى", "أقصى", "maximum", "max", "على الاخر", "على الآخر", "للاخر", "للآخر", "full"),
                ):
                    raw = await DesktopTool().execute(action="volume", mode="max")
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل ضبط الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم ضبط الصوت على الحد الأقصى ({lvl}%)." if arabic else f"Volume set to maximum ({lvl}%)."}

                if has_any(normalized, ("اكتم", "كتم", "اسكت", "سكت", "mute")):
                    raw = await DesktopTool().execute(action="volume", mode="mute")
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل كتم الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    return {"type": "message", "content": "تم كتم الصوت." if arabic else f"Muted audio (state={data})."}

                if has_any(normalized, ("الغ كتم", "الغي كتم", "الغ الكتم", "الغي الكتم", "فك كتم", "فك الكتم", "unmute")):
                    raw = await DesktopTool().execute(action="volume", mode="unmute")
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل إلغاء الكتم: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم إلغاء كتم الصوت ({lvl}%)." if arabic else f"Audio unmuted ({lvl}%)."}

                number = extract_first_number()
                if (
                    number is not None
                    and has_any(
                        normalized,
                        (
                            "اجعل",
                            "خلي",
                            "خليه",
                            "خليها",
                            "اضبط",
                            "اعمله",
                            "اعمله",
                            "حطه",
                            "set",
                            "%",
                            "بالميه",
                            "بالمئة",
                            "الى",
                            "ل",
                        ),
                    )
                ) or (
                    number is not None
                    and has_any(normalized, ("ارفع", "زيد", "علي", "نزل", "اخفض", "خفض", "وطي"))
                    and re.search(r"(?:\bالى\b|\bl\b|\bto\b)", normalized)
                ):
                    number = max(0, min(100, number))
                    raw = await DesktopTool().execute(action="volume", mode="set", level=number)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل ضبط الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم ضبط الصوت على {lvl}%." if arabic else f"Volume set to {lvl}%."}

                if has_any(normalized, ("ارفع", "زيد", "علي", "عليها", "عليه", "increase", "up")):
                    raw = await DesktopTool().execute(action="volume", mode="up", delta=8)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل رفع الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم رفع الصوت إلى {lvl}%." if arabic else f"Volume increased to {lvl}%."}

                if has_any(normalized, ("نزل", "اخفض", "خفض", "وطي", "قلل", "decrease", "down")):
                    raw = await DesktopTool().execute(action="volume", mode="down", delta=8)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل خفض الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم خفض الصوت إلى {lvl}%." if arabic else f"Volume decreased to {lvl}%."}
            except Exception:
                pass

        # Contextual semantic setter (e.g. "اعمله ماكس") using recent conversation topic.
        if recent_topic == "volume" and has_any(
            normalized,
            (
                "ماكس",
                "اقصى",
                "أقصى",
                "maximum",
                "max",
                "على الاخر",
                "على الآخر",
                "للآخر",
                "للصفر",
                "صفر",
                "اكتم",
                "كتم",
                "فك كتم",
                "فك الكتم",
                "الغ الكتم",
                "الغي الكتم",
            ),
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                if has_any(normalized, ("اكتم", "كتم")):
                    raw = await DesktopTool().execute(action="volume", mode="mute")
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل كتم الصوت: {err}" if arabic else err}
                    return {"type": "message", "content": "تم كتم الصوت." if arabic else "Muted audio."}
                if has_any(normalized, ("فك كتم", "فك الكتم", "الغ الكتم", "الغي الكتم")):
                    raw = await DesktopTool().execute(action="volume", mode="unmute")
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل إلغاء الكتم: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {"type": "message", "content": f"تم إلغاء كتم الصوت ({lvl}%)." if arabic else f"Audio unmuted ({lvl}%)."}
                if has_any(normalized, ("للصفر", "صفر")):
                    raw = await DesktopTool().execute(action="volume", mode="set", level=0)
                else:
                    raw = await DesktopTool().execute(action="volume", mode="max")
                data, err = parse_tool_json(raw)
                if err:
                    return {"type": "message", "content": f"فشل ضبط الصوت: {err}" if arabic else err}
                data = data if isinstance(data, dict) else {}
                lvl = data.get("level_percent")
                return {
                    "type": "message",
                    "content": f"تم ضبط الصوت على {lvl}%." if arabic else f"Volume set to {lvl}%.",
                }
            except Exception:
                pass

        # Contextual numeric setter (e.g. "اعمله 37") using recent conversation topic.
        contextual_number = extract_first_number()
        if contextual_number is not None and has_any(
            normalized,
            (
                "اعمله",
                "خليه",
                "خليها",
                "اجعله",
                "اجعلها",
                "اضبطه",
                "اضبطها",
                "حطه",
                "set",
                "ارفعه",
                "ارفعها",
                "ارفع",
                "نزله",
                "نزلها",
                "نزل",
                "خفضه",
                "خفضها",
                "خفض",
                "اخفض",
            ),
        ):
            topic = recent_topic
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                if topic == "volume":
                    level = max(0, min(100, contextual_number))
                    raw = await DesktopTool().execute(action="volume", mode="set", level=level)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل ضبط الصوت: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("level_percent")
                    return {
                        "type": "message",
                        "content": f"تم ضبط الصوت على {lvl}%." if arabic else f"Volume set to {lvl}%.",
                    }
                if topic == "brightness":
                    level = max(0, min(100, contextual_number))
                    raw = await DesktopTool().execute(action="brightness", mode="set", level=level)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل ضبط السطوع: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {
                        "type": "message",
                        "content": f"تم ضبط السطوع على {lvl}%." if arabic else f"Brightness set to {lvl}%.",
                    }
            except Exception:
                pass

        # Focus/activate a specific application window using UI Automation.
        if has_any(normalized, ("focus", "activate", "switch to", "ركز", "نشط", "فعّل النافذه", "فعل النافذة")) and has_any(
            normalized, ("window", "app", "application", "نافذه", "نافذة", "تطبيق", "برنامج")
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                window_query = ""
                m = re.search(
                    r"(?:نافذه|نافذة|window|app|application|تطبيق|برنامج)\s+(.+)$",
                    text,
                    re.IGNORECASE,
                )
                if m:
                    window_query = clean_query(m.group(1) or "")
                if not window_query:
                    window_query = extract_trailing_query()
                if window_query:
                    raw = await DesktopTool().execute(action="focus_window", window_title=window_query, timeout_sec=5.0)
                    data, err = parse_tool_json(raw)
                    if err:
                        return {"type": "message", "content": f"فشل تفعيل النافذة: {err}" if arabic else err}
                    data = data if isinstance(data, dict) else {}
                    title = str(data.get("title", "") or window_query)
                    return {
                        "type": "message",
                        "content": f"تم تفعيل نافذة: {title}" if arabic else f"Focused window: {title}",
                    }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تفعيل النافذة: {e}" if arabic else f"Failed to focus window: {e}",
                }

        # Move mouse to a UI element by label/name (e.g. "move mouse to Save button").
        if has_any(normalized, ("حرك", "move", "hover", "وجّه", "وجه")) and self._wants_pointer_control(text) and mentions_ui_target_context(
            text
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                control_name, window_name = extract_ui_target(text)
                if is_pronoun_target(control_name):
                    prev_target, prev_window = infer_recent_ui_target()
                    control_name = prev_target or control_name
                    window_name = window_name or prev_window
                if not control_name:
                    return {
                        "type": "message",
                        "content": "حدد العنصر المطلوب، أو قل: حرك عليه."
                        if arabic
                        else "Please specify the button/element name to move to.",
                    }

                raw = await DesktopTool().execute(
                    action="ui_target",
                    control_name=control_name,
                    window_title=window_name,
                    interaction="move",
                    duration=0.25,
                    timeout_sec=6.0,
                )
                data, err = parse_tool_json(raw)
                if err:
                    vision_raw = await DesktopTool().execute(
                        action="vision_tools",
                        mode="locate_element",
                        target=control_name,
                        window_hint=window_name,
                        interaction="move",
                    )
                    vision_data, vision_err = parse_tool_json(vision_raw)
                    if not vision_err and isinstance(vision_data, dict) and bool(vision_data.get("ok")):
                        picked = str(vision_data.get("matched_label") or control_name)
                        return {
                            "type": "message",
                            "content": (f"🎯 تم توجيه الماوس إلى: {picked}" if arabic else f"🎯 Mouse moved to: {picked}"),
                        }
                    return {
                        "type": "message",
                        "content": f"ما قدرت أحدد العنصر المطلوب." if arabic else "Could not locate that UI element.",
                    }

                data = data if isinstance(data, dict) else {}
                x = int(data.get("x", 0) or 0)
                y = int(data.get("y", 0) or 0)
                picked = str(data.get("control_name", "") or control_name)
                if arabic:
                    msg = f"تم تحريك الماوس إلى العنصر: {picked}"
                    if x > 0 or y > 0:
                        msg += "\nتم توجيه المؤشر إلى العنصر."
                    return {"type": "message", "content": msg}
                msg = f"Moved mouse to UI element: {picked}"
                if x > 0 or y > 0:
                    msg += "\nCursor moved to the target element."
                return {"type": "message", "content": msg}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تحريك الماوس للعنصر المطلوب: {e}"
                    if arabic
                    else f"Failed moving mouse to requested UI element: {e}",
                }

        # Click UI controls by label/text (e.g. "click the Save button in Settings").
        keyboard_key_tokens = (
            "ctrl",
            "control",
            "كنترول",
            "كترل",
            "alt",
            "shift",
            "win",
            "windows",
            "esc",
            "escape",
            "tab",
            "enter",
            "prtsc",
            "printscreen",
            "f1",
            "f2",
            "f3",
            "f4",
            "f5",
            "f6",
            "f7",
            "f8",
            "f9",
            "f10",
            "f11",
            "f12",
        )
        if has_any(normalized, ("اضغط", "اكبس", "انقر", "click", "press")) and (
            mentions_ui_target_context(text)
            or (
                self._wants_pointer_control(text)
                and extract_first_two_numbers() is None
                and has_any(normalized, ("على", "on", "to", "towards", "لعند"))
            )
        ) and not has_any(normalized, keyboard_key_tokens + ("keyboard", "كيبورد", "shortcut", "اختصار", "key")):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                control_name, window_name = extract_ui_target(text)
                if is_pronoun_target(control_name):
                    prev_target, prev_window = infer_recent_ui_target()
                    control_name = prev_target or control_name
                    window_name = window_name or prev_window
                if not control_name:
                    return {
                        "type": "message",
                        "content": "حدد العنصر المطلوب، أو قل: اضغط عليه."
                        if arabic
                        else "Please specify which button/element to click.",
                    }
                raw = await DesktopTool().execute(
                    action="ui_target",
                    control_name=control_name,
                    window_title=window_name,
                    interaction="click",
                    timeout_sec=5.0,
                )
                data, err = parse_tool_json(raw)
                if err:
                    vision_raw = await DesktopTool().execute(
                        action="vision_tools",
                        mode="locate_element",
                        target=control_name,
                        window_hint=window_name,
                        interaction="click",
                    )
                    vision_data, vision_err = parse_tool_json(vision_raw)
                    if not vision_err and isinstance(vision_data, dict) and bool(vision_data.get("ok")):
                        clicked_name = str(vision_data.get("matched_label") or control_name)
                        return {
                            "type": "message",
                            "content": (f"✅ تم الضغط على: {clicked_name}" if arabic else f"✅ Clicked: {clicked_name}"),
                        }
                    return {
                        "type": "message",
                        "content": "ما قدرت أضغط العنصر المطلوب." if arabic else "Could not click that UI element.",
                    }
                data = data if isinstance(data, dict) else {}
                clicked_name = data.get("control_name") or control_name
                win_title = data.get("window_title") or window_name
                if arabic:
                    msg = f"تم الضغط على الزر: {clicked_name}" if clicked_name else "تم الضغط على العنصر المطلوب."
                    if win_title:
                        msg += f"\nضمن النافذة: {win_title}"
                    return {"type": "message", "content": msg}
                msg = f"Clicked button: {clicked_name}" if clicked_name else "Clicked requested UI element."
                if win_title:
                    msg += f"\nWindow: {win_title}"
                return {"type": "message", "content": msg}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تنفيذ النقر على الزر: {e}" if arabic else f"UI button click failed: {e}",
                }

        # Type text into UI controls by label/name.
        if has_any(normalized, ("اكتب", "type", "write", "ادخل", "enter text")) and (
            has_any(normalized, ("حقل", "مربع", "textbox", "field", "input", "نافذه", "نافذة", "window", "app"))
            or recent_topic == "windows"
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                typed_text = ""
                q = re.search(r"[\"“](.+?)[\"”]|'(.+?)'", text)
                if q:
                    typed_text = (q.group(1) or q.group(2) or "").strip()
                if not typed_text:
                    typed_text = extract_typed_text(text)

                control_name = ""
                window_name = ""
                window_match = re.search(r"(?:\s+(?:في نافذه|في نافذة|داخل نافذه|داخل نافذة|in window)\s+)(.+)$", text, re.IGNORECASE)
                if window_match:
                    window_name = clean_query(window_match.group(1) or "")
                target_match = re.search(r"(?:حقل|مربع|textbox|field|input)\s+(.+?)(?:\s+(?:في|داخل|in)\s+.+)?$", text, re.IGNORECASE)
                if target_match:
                    control_name = clean_query(target_match.group(1) or "")

                if not typed_text:
                    return {
                        "type": "message",
                        "content": "اذكر النص الذي تريد كتابته داخل علامات اقتباس."
                        if arabic
                        else "Please provide the text to type (preferably in quotes).",
                    }

                raw = await DesktopTool().execute(
                    action="ui_set_text",
                    text=typed_text,
                    control_name=control_name,
                    window_title=window_name,
                    press_enter=has_any(
                        normalized,
                        ("press enter", "اضغط انتر", "اكبس انتر", "ثم انتر", "enter key", "زر enter", "send", "ارسل"),
                    ),
                    timeout_sec=5.0,
                )
                data, err = parse_tool_json(raw)
                if err:
                    # Fallback for active-window typing when direct UI targeting is unavailable.
                    if not control_name:
                        fallback_raw = await DesktopTool().execute(
                            action="type_text",
                            text=typed_text,
                            press_enter=has_any(
                                normalized,
                                (
                                    "press enter",
                                    "اضغط انتر",
                                    "اكبس انتر",
                                    "ثم انتر",
                                    "enter key",
                                    "زر enter",
                                    "send",
                                    "ارسل",
                                ),
                            ),
                            interval=0.01,
                        )
                        fallback_data, fallback_err = parse_tool_json(fallback_raw)
                        if not fallback_err and isinstance(fallback_data, dict) and fallback_data.get("ok"):
                            return {
                                "type": "message",
                                "content": "تمت الكتابة في النافذة النشطة."
                                if arabic
                                else "Typed text in the active window.",
                            }
                    return {
                        "type": "message",
                        "content": f"فشل إدخال النص: {err}" if arabic else f"Failed to type text: {err}",
                    }
                data = data if isinstance(data, dict) else {}
                ctl = data.get("control_name") or control_name
                if arabic:
                    return {
                        "type": "message",
                        "content": f"تم إدخال النص بنجاح{' في ' + ctl if ctl else ''}.",
                    }
                return {"type": "message", "content": f"Text entered successfully{' in ' + ctl if ctl else ''}."}
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل إدخال النص: {e}" if arabic else f"Failed to type text: {e}",
                }

        # Generic keyboard typing into active window (no explicit UI field target).
        if has_any(normalized, ("اكتب", "type", "write", "ادخل")) and not has_any(
            normalized, ("حقل", "مربع", "textbox", "field", "input")
        ):
            quoted_text = bool(re.search(r"[\"“](.+?)[\"”]|'(.+?)'", text))
            keyboard_context = has_any(
                normalized,
                (
                    "keyboard",
                    "كيبورد",
                    "بالكيبورد",
                    "لوحة المفاتيح",
                    "على الكيبورد",
                    "active window",
                    "window",
                    "app",
                    "تطبيق",
                    "برنامج",
                    "نافذه",
                    "نافذة",
                ),
            )
            if quoted_text or keyboard_context or recent_topic == "windows":
                typed_text = extract_typed_text(text)
                if typed_text:
                    try:
                        from Mudabbir.tools.builtin.desktop import DesktopTool

                        raw = await DesktopTool().execute(
                            action="type_text",
                            text=typed_text,
                            press_enter=has_any(
                                normalized,
                                (
                                    "press enter",
                                    "اضغط انتر",
                                    "اكبس انتر",
                                    "ثم انتر",
                                    "enter key",
                                    "زر enter",
                                    "send",
                                    "ارسل",
                                ),
                            ),
                            interval=0.01,
                        )
                        data, err = parse_tool_json(raw)
                        if err:
                            return {
                                "type": "message",
                                "content": f"فشل تنفيذ الكتابة على الكيبورد: {err}"
                                if arabic
                                else f"Keyboard typing failed: {err}",
                            }
                        data = data if isinstance(data, dict) else {}
                        chars = int(data.get("chars", len(typed_text)) or len(typed_text))
                        if arabic:
                            return {
                                "type": "message",
                                "content": f"تمت الكتابة على الكيبورد بنجاح ({chars} حرف).",
                            }
                        return {
                            "type": "message",
                            "content": f"Typed successfully via keyboard ({chars} chars).",
                        }
                    except Exception as e:
                        return {
                            "type": "message",
                            "content": f"فشل تنفيذ الكتابة على الكيبورد: {e}"
                            if arabic
                            else f"Keyboard typing failed: {e}",
                        }

        # Keyboard key press / hotkeys.
        key_hint_pattern = re.compile(
            r"\b(ctrl|control|كنترول|كترل|shift|shft|alt|win|windows|esc|escape|tab|enter|انتر|"
            r"space|delete|del|backspace|prtsc|printscreen|print\s+screen|f(?:[1-9]|1[0-2]))\b"
        )
        key_hint_present = bool(key_hint_pattern.search(normalized))
        if has_any(normalized, ("اضغط", "اكبس", "press", "انقر")) and (
            has_any(normalized, ("زر", "key", "كيبورد", "keyboard", "بالكيبورد", "shortcut", "اختصار"))
            or "+" in normalized
            or key_hint_present
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                if "+" in normalized:
                    keys = parse_compound_keys(normalized)
                    if len(keys) >= 2:
                        raw = await DesktopTool().execute(action="hotkey", keys=keys)
                        if raw.lower().startswith("error:"):
                            return {"type": "message", "content": raw.replace("Error: ", "", 1)}
                        return {
                            "type": "message",
                            "content": f"تم تنفيذ الاختصار: {' + '.join(keys)}"
                            if arabic
                            else f"Hotkey executed: {' + '.join(keys)}",
                        }

                key_aliases = {
                    "win": ("win", "windows", "window", "زر الويندوز", "ويندوز"),
                    "ctrl": ("ctrl", "control", "كنترول", "كترل"),
                    "shift": ("shift", "شفت"),
                    "alt": ("alt", "الت"),
                    "esc": ("esc", "escape", "زر الهروب"),
                    "enter": ("enter", "انتر", "return"),
                    "tab": ("tab",),
                    "space": ("space", "مسافة"),
                    "delete": ("delete", "del"),
                    "backspace": ("backspace",),
                    "printscreen": ("prtsc", "printscreen", "print screen"),
                }
                selected_key = ""
                for key_name, aliases in key_aliases.items():
                    if has_any(normalized, aliases):
                        selected_key = key_name
                        break
                if not selected_key and re.search(r"\bf([1-9]|1[0-2])\b", normalized):
                    selected_key = re.search(r"\bf([1-9]|1[0-2])\b", normalized).group(0)

                if selected_key:
                    raw = await DesktopTool().execute(action="press_key", key=selected_key)
                    if raw.lower().startswith("error:"):
                        return {"type": "message", "content": raw.replace("Error: ", "", 1)}
                    return {
                        "type": "message",
                        "content": f"تم ضغط الزر: {selected_key}" if arabic else f"Pressed key: {selected_key}",
                    }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تنفيذ أمر الكيبورد: {e}" if arabic else f"Keyboard command failed: {e}",
                }

        # Mouse click by coordinates.
        if has_any(normalized, ("اكبس", "اضغط", "انقر", "click")):
            coords = extract_first_two_numbers()
            if coords is not None and (
                self._wants_pointer_control(text) or has_any(normalized, ("هنا", "here"))
            ):
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    x, y = coords
                    raw = await DesktopTool().execute(action="click", x=int(x), y=int(y), button="left", clicks=1)
                    if raw.lower().startswith("error:"):
                        return {"type": "message", "content": raw.replace("Error: ", "", 1)}
                    return {
                        "type": "message",
                        "content": f"تم النقر عند ({int(x)}, {int(y)})." if arabic else f"Clicked at ({int(x)}, {int(y)}).",
                    }
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل تنفيذ النقر: {e}" if arabic else f"Click failed: {e}",
                    }

        # Start menu / Windows key.
        wants_start_menu = (
            has_any(
                normalized,
                (
                    "start menu",
                    "open start",
                    "windows key",
                    "win key",
                    "windows",
                    "win",
                    "قائمه ابدا",
                    "قايمه ابدا",
                    "زر الويندوز",
                ),
            )
            and has_any(normalized, ("open", "press", "افتح", "اضغط", "اكبس", "انقر"))
        ) or normalized in {"win", "windows"}
        if wants_start_menu:
            try:
                import pyautogui

                pyautogui.FAILSAFE = False
                pyautogui.press("win")
                return {"type": "message", "content": "تم فتح قائمة ابدأ." if arabic else "Start menu opened."}
            except Exception as e:
                return {"type": "message", "content": f"فشل فتح قائمة ابدأ: {e}" if arabic else f"Failed to open Start menu: {e}"}

        # Brightness controls.
        if has_any(normalized, ("سطوع", "brightness", "الاضاءه", "الاضاءة", "اضاءه", "اضاءة", "اضويه", "اضوي")):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                if has_any(normalized, ("كم", "الحالي", "current", "status")):
                    raw = await DesktopTool().execute(action="brightness", mode="get")
                    data, err = parse_tool_json(raw)
                    if err:
                        friendly = normalize_brightness_error(err)
                        return {
                            "type": "message",
                            "content": f"فشل قراءة السطوع: {friendly}" if arabic else friendly,
                        }
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {"type": "message", "content": f"السطوع الحالي: {lvl}%." if arabic else f"Current brightness: {lvl}%."}

                if has_any(normalized, ("ماكس", "اقصى", "maximum", "max")):
                    raw = await DesktopTool().execute(action="brightness", mode="max")
                    data, err = parse_tool_json(raw)
                    if err:
                        friendly = normalize_brightness_error(err)
                        return {"type": "message", "content": f"فشل ضبط السطوع: {friendly}" if arabic else friendly}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {"type": "message", "content": f"تم رفع السطوع إلى {lvl}%." if arabic else f"Brightness set to {lvl}%."}

                if has_any(normalized, ("صفر", "اقل", "minimum", "min")):
                    raw = await DesktopTool().execute(action="brightness", mode="min")
                    data, err = parse_tool_json(raw)
                    if err:
                        friendly = normalize_brightness_error(err)
                        return {"type": "message", "content": f"فشل ضبط السطوع: {friendly}" if arabic else friendly}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {"type": "message", "content": f"تم خفض السطوع إلى {lvl}%." if arabic else f"Brightness set to {lvl}%."}

                number = extract_first_number()
                if number is not None and has_any(
                    normalized,
                    (
                        "اجعل",
                        "خلي",
                        "خليه",
                        "خليها",
                        "اضبط",
                        "اعمله",
                        "حطه",
                        "set",
                        "%",
                        "بالميه",
                        "بالمئة",
                    ),
                ):
                    number = max(0, min(100, number))
                    raw = await DesktopTool().execute(action="brightness", mode="set", level=number)
                    data, err = parse_tool_json(raw)
                    if err:
                        friendly = normalize_brightness_error(err)
                        return {"type": "message", "content": f"فشل ضبط السطوع: {friendly}" if arabic else friendly}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {"type": "message", "content": f"تم ضبط السطوع على {lvl}%." if arabic else f"Brightness set to {lvl}%."}

                if has_any(normalized, ("ارفع", "زيد", "علي", "عليها", "عليه", "increase", "up")):
                    raw = await DesktopTool().execute(action="brightness", mode="up", delta=10)
                    data, err = parse_tool_json(raw)
                    if err:
                        friendly = normalize_brightness_error(err)
                        return {"type": "message", "content": f"فشل ضبط السطوع: {friendly}" if arabic else friendly}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {"type": "message", "content": f"تم رفع السطوع إلى {lvl}%." if arabic else f"Brightness increased to {lvl}%."}

                if has_any(normalized, ("خفض", "نزل", "وطي", "قلل", "decrease", "down")):
                    raw = await DesktopTool().execute(action="brightness", mode="down", delta=10)
                    data, err = parse_tool_json(raw)
                    if err:
                        friendly = normalize_brightness_error(err)
                        return {"type": "message", "content": f"فشل ضبط السطوع: {friendly}" if arabic else friendly}
                    data = data if isinstance(data, dict) else {}
                    lvl = data.get("brightness_percent")
                    return {"type": "message", "content": f"تم خفض السطوع إلى {lvl}%." if arabic else f"Brightness decreased to {lvl}%."}
            except Exception as e:
                return {"type": "message", "content": f"فشل التحكم بالسطوع: {e}" if arabic else f"Brightness control failed: {e}"}

        # Media controls.
        if has_any(normalized, ("ميديا", "media", "music", "موسيقى", "اغنيه", "اغنية")):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                if has_any(normalized, ("التالي", "next")):
                    raw = await DesktopTool().execute(action="media_control", mode="next")
                elif has_any(normalized, ("السابق", "previous", "prev")):
                    raw = await DesktopTool().execute(action="media_control", mode="previous")
                elif has_any(normalized, ("ايقاف", "وقف", "stop")):
                    raw = await DesktopTool().execute(action="media_control", mode="stop")
                elif has_any(normalized, ("تشغيل", "pause", "play", "استكمال")):
                    raw = await DesktopTool().execute(action="media_control", mode="play_pause")
                else:
                    raw = await DesktopTool().execute(action="media_control", mode="play_pause")
                return {"type": "message", "content": raw}
            except Exception as e:
                return {"type": "message", "content": f"فشل التحكم بالميديا: {e}" if arabic else f"Media control failed: {e}"}

        # Move mouse to a file icon on Desktop (e.g. "move mouse to game file on desktop").
        if (
            has_any(normalized, ("حرك", "move", "hover", "وجّه", "وجه"))
            and self._wants_pointer_control(text)
            and has_any(normalized, ("file", "files", "ملف", "ملفات"))
            and has_any(normalized, ("desktop", "سطح المكتب", "سطح", "المكتب"))
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                desktop_dir = Path.home() / "Desktop"
                if not desktop_dir.is_dir():
                    return {
                        "type": "message",
                        "content": "تعذر الوصول إلى مجلد سطح المكتب." if arabic else "Desktop folder is not accessible.",
                    }

                query = extract_desktop_file_query(text)
                if not query:
                    return {
                        "type": "message",
                        "content": "اذكر اسم الملف على سطح المكتب." if arabic else "Please specify the desktop file name.",
                    }

                query_cf = query.casefold()
                items: list[Path] = []
                for item in desktop_dir.iterdir():
                    name_cf = item.name.casefold()
                    stem_cf = item.stem.casefold()
                    if query_cf == name_cf or query_cf == stem_cf:
                        items.append(item)
                        continue
                    if query_cf in name_cf or query_cf in stem_cf:
                        items.append(item)
                if not items:
                    return {
                        "type": "message",
                        "content": f"ما لقيت ملف '{query}' على سطح المكتب."
                        if arabic
                        else f"Could not find '{query}' on Desktop.",
                    }
                items.sort(key=lambda p: (len(p.name), p.name.casefold()))
                target_item = items[0]
                target_label = target_item.stem or target_item.name

                attempts = [
                    {"window_title": "Desktop", "control_type": "ListItem"},
                    {"window_title": "Desktop", "control_type": ""},
                    {"window_title": "", "control_type": "ListItem"},
                ]
                last_err = ""
                for attempt in attempts:
                    raw = await DesktopTool().execute(
                        action="ui_target",
                        window_title=attempt["window_title"],
                        control_name=target_label,
                        control_type=attempt["control_type"],
                        interaction="move",
                        duration=0.2,
                        timeout_sec=2.5,
                    )
                    data, err = parse_tool_json(raw)
                    if not err and isinstance(data, dict) and data.get("ok"):
                        x = data.get("x")
                        y = data.get("y")
                        if arabic:
                            return {
                                "type": "message",
                                "content": f"تم تحريك الماوس إلى الملف '{target_item.name}' على سطح المكتب.",
                            }
                        return {
                            "type": "message",
                            "content": f"Moved mouse to desktop file '{target_item.name}'.",
                        }
                    last_err = err or str(raw)

                # Fallback: open Desktop folder and retry once in Explorer window.
                try:
                    subprocess.Popen(["explorer.exe", str(desktop_dir)], shell=False)
                    await asyncio.sleep(0.8)
                    raw = await DesktopTool().execute(
                        action="ui_target",
                        window_title="Desktop",
                        control_name=target_label,
                        control_type="ListItem",
                        interaction="move",
                        duration=0.2,
                        timeout_sec=4.0,
                    )
                    data, err = parse_tool_json(raw)
                    if not err and isinstance(data, dict) and data.get("ok"):
                        x = data.get("x")
                        y = data.get("y")
                        if arabic:
                            return {
                                "type": "message",
                                "content": f"تم تحريك الماوس إلى '{target_item.name}' بعد فتح سطح المكتب.",
                            }
                        return {
                            "type": "message",
                            "content": f"Moved mouse to '{target_item.name}' after opening Desktop.",
                        }
                    last_err = err or str(raw)
                except Exception:
                    pass

                if arabic:
                    return {
                        "type": "message",
                        "content": f"لقيت الملف '{target_item.name}' لكن ما قدرت أحدد مكان الأيقونة على الشاشة.\nالمسار: {target_item}\nالسبب: {last_err}",
                    }
                return {
                    "type": "message",
                    "content": f"Found '{target_item.name}' but could not locate its icon on screen.\nPath: {target_item}\nReason: {last_err}",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تحريك الماوس لملف سطح المكتب: {e}"
                    if arabic
                    else f"Failed moving mouse to desktop file: {e}",
                }

        # Absolute mouse move by coordinates.
        if has_any(normalized, ("حرك", "move")) and self._wants_pointer_control(text):
            coords = extract_first_two_numbers()
            if coords is not None:
                try:
                    import pyautogui

                    x, y = coords
                    pyautogui.FAILSAFE = False
                    if "%" in text:
                        size = pyautogui.size()
                        x = int(max(0.0, min(100.0, float(x))) / 100.0 * size.width)
                        y = int(max(0.0, min(100.0, float(y))) / 100.0 * size.height)
                    pyautogui.moveTo(int(x), int(y), duration=0.2)
                    return {
                        "type": "message",
                        "content": f"تم تحريك الماوس إلى ({int(x)}, {int(y)})." if arabic else f"Moved mouse to ({int(x)}, {int(y)}).",
                    }
                except Exception as e:
                    return {
                        "type": "message",
                        "content": f"فشل تحريك الماوس للإحداثيات المطلوبة: {e}" if arabic else f"Failed moving mouse to coordinates: {e}",
                    }
            else:
                # Best-effort: move to center of a target app/window mentioned in text.
                try:
                    from Mudabbir.tools.builtin.desktop import DesktopTool

                    target_match = re.search(r"(?:لعند|على|to|towards)\s+(.+)$", text, re.IGNORECASE)
                    target_query = clean_query(target_match.group(1) if target_match else text)
                    if target_query:
                        raw = await DesktopTool().execute(action="list_windows", include_untitled=False, limit=120)
                        data, err = parse_tool_json(raw)
                        if not err and isinstance(data, dict):
                            windows = data.get("windows", []) if isinstance(data, dict) else []
                            nq = _normalize_text_for_match(target_query)
                            for w in windows:
                                title = str((w or {}).get("title", "") or "")
                                ntitle = _normalize_text_for_match(title)
                                if not title or not nq or nq not in ntitle:
                                    continue
                                left = int((w or {}).get("left", 0) or 0)
                                top = int((w or {}).get("top", 0) or 0)
                                width = int((w or {}).get("width", 0) or 0)
                                height = int((w or {}).get("height", 0) or 0)
                                x = left + max(0, width // 2)
                                y = top + max(0, height // 2)
                                move_raw = await DesktopTool().execute(action="mouse_move", x=x, y=y, duration=0.2)
                                if str(move_raw).lower().startswith("error:"):
                                    return {
                                        "type": "message",
                                        "content": move_raw.replace("Error: ", "", 1),
                                    }
                                return {
                                    "type": "message",
                                    "content": f"تم تحريك الماوس إلى نافذة '{title}'."
                                    if arabic
                                    else f"Moved mouse to '{title}'.",
                                }
                except Exception:
                    pass

        # Keyboard language switch.
        if has_any(
            normalized,
            (
                "غير لغه الكيبورد",
                "تغيير لغه الكيبورد",
                "keyboard language",
                "switch keyboard",
                "change keyboard language",
                "غير اللغه",
                "غير اللغة",
                "بدل اللغه",
                "بدل اللغة",
                "switch language",
                "change language",
            ),
        ) or (
            has_any(normalized, ("عربي", "انجليزي", "english", "arabic"))
            and has_any(normalized, ("غير", "بدل", "switch", "change"))
        ):
            try:
                import pyautogui

                pyautogui.FAILSAFE = False
                pyautogui.hotkey("alt", "shift")
                return {
                    "type": "message",
                    "content": "تم تبديل لغة لوحة المفاتيح (Alt+Shift)." if arabic else "Keyboard language switched (Alt+Shift).",
                }
            except Exception as e:
                return {
                    "type": "message",
                    "content": f"فشل تبديل لغة الكيبورد: {e}" if arabic else f"Failed to switch keyboard language: {e}",
                }

        # Count open windows.
        wants_window_count = any(
            key in lowered
            for key in (
                "open windows",
                "window count",
                "mainwindowtitle",
                "count windows",
            )
        ) or has_any(normalized, ("كم نافذه", "عدد النوافذ", "النوافذ المفتوحه", "كم نافذة", "النوافذ المفتوحة"))
        if wants_window_count:
            count = self._count_open_windows()
            if count is None:
                if arabic:
                    return {"type": "message", "content": "تعذر حساب عدد النوافذ المفتوحة."}
                return {"type": "message", "content": "Could not count open windows."}
            if arabic:
                return {"type": "message", "content": f"عدد النوافذ المفتوحة حاليًا: {count}"}
            return {"type": "message", "content": f"Open windows count: {count}"}

        # Volume shortcuts.
        if has_any(normalized, ("ارفع الصوت", "علي الصوت", "volume up")):
            try:
                import pyautogui

                pyautogui.FAILSAFE = False
                for _ in range(4):
                    pyautogui.press("volumeup")
                return {"type": "message", "content": "تم رفع الصوت." if arabic else "Volume increased."}
            except Exception as e:
                return {"type": "message", "content": f"فشل رفع الصوت: {e}" if arabic else f"Failed to increase volume: {e}"}

        if has_any(normalized, ("نزل الصوت", "اخفض الصوت", "وطى الصوت", "volume down")):
            try:
                import pyautogui

                pyautogui.FAILSAFE = False
                for _ in range(4):
                    pyautogui.press("volumedown")
                return {"type": "message", "content": "تم خفض الصوت." if arabic else "Volume decreased."}
            except Exception as e:
                return {"type": "message", "content": f"فشل خفض الصوت: {e}" if arabic else f"Failed to decrease volume: {e}"}

        # Common keyboard shortcuts.
        if has_any(normalized, ("esc", "escape", "زر الهروب")):
            try:
                import pyautogui

                pyautogui.FAILSAFE = False
                pyautogui.press("esc")
                return {"type": "message", "content": "تم ضغط زر ESC." if arabic else "Pressed ESC."}
            except Exception as e:
                return {"type": "message", "content": f"فشل ضغط ESC: {e}" if arabic else f"Failed to press ESC: {e}"}

        if has_any(normalized, ("prtsc", "print screen", "سكرين شوت", "printscreen")):
            try:
                import pyautogui

                pyautogui.FAILSAFE = False
                pyautogui.press("printscreen")
                return {"type": "message", "content": "تم ضغط زر Print Screen." if arabic else "Pressed Print Screen."}
            except Exception as e:
                return {"type": "message", "content": f"فشل ضغط Print Screen: {e}" if arabic else f"Failed to press Print Screen: {e}"}

        # Camera snapshot (direct tool call).
        if ("camera" in lowered or "كاميرا" in text) and (
            "snapshot" in lowered
            or "صورة" in text
            or "لقطة" in text
            or "افتح" in text
            or "شغل" in text
        ):
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                result = await DesktopTool().execute(action="camera_snapshot", camera_index=0)
                if isinstance(result, str) and result.lower().startswith("error:"):
                    reason = result.replace("Error: ", "", 1)
                    return {
                        "type": "message",
                        "content": f"فشل تنفيذ أمر الكاميرا: {reason}" if arabic else f"Camera action failed: {reason}",
                    }
                details = str(result).strip()
                return {
                    "type": "message",
                    "content": f"تم التقاط صورة من الكاميرا بنجاح.\n{details}"
                    if arabic
                    else f"Camera snapshot captured successfully.\n{details}",
                }
            except Exception as e:
                return {"type": "message", "content": f"فشل تشغيل الكاميرا: {e}" if arabic else f"Camera action failed: {e}"}

        # Microphone recording (direct tool call).
        if "microphone" in lowered or "mic" in lowered or "مايك" in text or "ميكروفون" in text:
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                result = await DesktopTool().execute(action="microphone_record", seconds=3.0)
                if isinstance(result, str) and result.lower().startswith("error:"):
                    reason = result.replace("Error: ", "", 1)
                    return {
                        "type": "message",
                        "content": f"فشل تنفيذ أمر المايك: {reason}" if arabic else f"Microphone action failed: {reason}",
                    }
                details = str(result).strip()
                return {
                    "type": "message",
                    "content": f"تم تسجيل مقطع صوتي بنجاح.\n{details}"
                    if arabic
                    else f"Microphone recording completed.\n{details}",
                }
            except Exception as e:
                return {"type": "message", "content": f"فشل تشغيل المايك: {e}" if arabic else f"Microphone action failed: {e}"}

        # Lightweight real-time screen watch (captures a short sequence).
        if "تحليل الشاشة" in text or "real-time screen" in lowered or "screen analysis" in lowered:
            try:
                from Mudabbir.tools.builtin.desktop import DesktopTool

                watch = await DesktopTool().execute(
                    action="screen_watch",
                    frames=3,
                    interval_sec=0.7,
                )
                overview = await DesktopTool().execute(action="desktop_overview")
                summary_title = "تحليل شاشة سريع" if arabic else "Quick screen analysis"
                try:
                    overview_data = json.loads(overview)
                    watch_data = json.loads(watch)
                    window_count = (
                        overview_data.get("windows", {}).get("count")
                        if isinstance(overview_data, dict)
                        else None
                    )
                    screen = overview_data.get("screen", {}) if isinstance(overview_data, dict) else {}
                    captured = watch_data.get("captured_frames") if isinstance(watch_data, dict) else None
                    paths = watch_data.get("paths", []) if isinstance(watch_data, dict) else []
                    first_path = paths[0] if paths else ""
                    if arabic:
                        content = (
                            f"{summary_title}:\n"
                            f"- أبعاد الشاشة: {screen.get('width', '?')}x{screen.get('height', '?')}\n"
                            f"- عدد النوافذ المفتوحة: {window_count if window_count is not None else '?'}\n"
                            f"- اللقطات الملتقطة: {captured if captured is not None else '?'}\n"
                            f"- أول لقطة: {first_path or 'غير متاحة'}"
                        )
                    else:
                        content = (
                            f"{summary_title}:\n"
                            f"- Screen: {screen.get('width', '?')}x{screen.get('height', '?')}\n"
                            f"- Open windows: {window_count if window_count is not None else '?'}\n"
                            f"- Captured frames: {captured if captured is not None else '?'}\n"
                            f"- First snapshot: {first_path or 'N/A'}"
                        )
                    return {"type": "message", "content": content}
                except Exception:
                    return {"type": "message", "content": f"{summary_title}:\n{overview}\n\n{watch}"}
            except Exception as e:
                return {"type": "message", "content": f"فشل تحليل الشاشة: {e}" if arabic else f"Screen analysis failed: {e}"}

        return None

    async def run(
        self,
        message: str,
        *,
        system_prompt: str | None = None,
        history: list[dict] | None = None,
        system_message: str | None = None,
    ) -> AsyncIterator[dict]:
        """Run a message through Open Interpreter with real-time streaming.

        Args:
            message: User message to process.
            system_prompt: Dynamic system prompt from AgentContextBuilder.
            history: Recent session history (prepended as summary to prompt).
            system_message: Legacy kwarg, superseded by system_prompt.
        """
        # Queue-aware admission for semaphore(1): emit hidden status heartbeats
        # while waiting so upstream first-chunk timeout does not trigger falsely.
        acquired = False
        queue_wait_started_at = time.monotonic()
        last_queue_heartbeat_at = 0.0

        try:
            while not acquired:
                try:
                    await asyncio.wait_for(
                        self._semaphore.acquire(),
                        timeout=OI_QUEUE_ACQUIRE_POLL_SECONDS,
                    )
                    acquired = True
                except TimeoutError:
                    waited = int(time.monotonic() - queue_wait_started_at)
                    if waited >= OI_QUEUE_WAIT_CAP_SECONDS:
                        logger.warning(
                            "Open Interpreter queue wait exceeded cap (wait=%ss cap=%ss)",
                            waited,
                            OI_QUEUE_WAIT_CAP_SECONDS,
                        )
                        yield {
                            "type": "error",
                            "content": (
                                "Agent backend is busy for too long. Please retry shortly."
                                if not _contains_arabic(message)
                                else "الباكند مشغول لفترة طويلة. حاول مرة ثانية بعد قليل."
                            ),
                        }
                        return

                    now = time.monotonic()
                    if (now - last_queue_heartbeat_at) >= OI_QUEUE_HEARTBEAT_SECONDS:
                        last_queue_heartbeat_at = now
                        yield {
                            "type": "status",
                            "content": "waiting_for_backend_slot",
                            "metadata": {
                                "wait_seconds": waited,
                                "wait_cap_seconds": OI_QUEUE_WAIT_CAP_SECONDS,
                            },
                        }

            self._stop_flag = False
            user_request = message
            is_arabic_request = _contains_arabic(user_request)
            quota_fallback_model = "gemini-2.5-flash"
            resolved_provider = str(getattr(self.settings, "llm_provider", "auto") or "auto")
            resolved_model = ""
            try:
                from Mudabbir.llm.client import resolve_llm_client

                resolved_llm = resolve_llm_client(self.settings)
                resolved_provider = resolved_llm.provider
                resolved_model = resolved_llm.model
            except Exception:
                if resolved_provider == "gemini":
                    resolved_model = str(getattr(self.settings, "gemini_model", "") or "")

            respond_mod = None
            original_display_markdown = None
            try:
                respond_mod = importlib.import_module("interpreter.core.respond")
                original_display_markdown = getattr(respond_mod, "display_markdown_message", None)
            except Exception:
                respond_mod = None
                original_display_markdown = None

            # Deterministic fast-path for frequent desktop requests that
            # were previously causing model loops.
            direct_result = await self._try_direct_task_manager_response(user_request)
            if direct_result is not None:
                yield direct_result
                return

            intent_map_result = await self._try_intent_map_desktop_response(user_request)
            if intent_map_result is not None:
                yield intent_map_result
                return
            if bool(getattr(self.settings, "windows_deterministic_hard_gate", True)) and self._is_gui_request(user_request):
                unresolved = resolve_windows_intent(user_request)
                if not unresolved.matched:
                    if is_arabic_request:
                        yield {
                            "type": "result",
                            "content": (
                                "لم أفهم أمر ويندوز بشكل كافٍ. جرّب صياغة مباشرة مثل: "
                                "'افتح مدير المهام'، 'اعرض سطح المكتب'، "
                                "'كم عملية Cursor وكلها سوا كم تستهلك'."
                            ),
                        }
                    else:
                        yield {
                            "type": "result",
                            "content": (
                                "I couldn't deterministically map that Windows command. "
                                "Try a direct form like: 'open task manager', 'show desktop', "
                                "or 'how many Cursor processes and total memory'."
                            ),
                        }
                    return

            ai_desktop_result = await self._try_ai_desktop_response(
                user_request, history=history
            )
            if ai_desktop_result is not None and ai_desktop_result.get("type") != "ai_noop":
                yield ai_desktop_result
                return

            # Deterministic fallback for resilience.
            # Important: this must also run when AI planner returns ai_noop,
            # otherwise requests fall through to raw OI and can leak execute JSON fragments.
            direct_desktop_result = await self._try_direct_desktop_response(
                user_request, history=history
            )
            if direct_desktop_result is not None:
                yield direct_desktop_result
                return

            if not self._interpreter:
                if is_arabic_request:
                    yield {"type": "message", "content": "Open Interpreter غير متاح حالياً."}
                else:
                    yield {"type": "message", "content": "❌ Open Interpreter not available."}
                return

            # Always include guardrails and prepend caller system prompt when present.
            base_system = system_prompt or system_message or ""
            task_guardrails = GUI_AUTOMATION_GUARDRAILS if self._is_gui_request(user_request) else ""
            guardrails = EXTRA_GUARDRAILS + task_guardrails
            effective_system = f"{base_system}\n\n{guardrails}" if base_system else guardrails
            original_system_message = self._interpreter.system_message
            self._interpreter.system_message = (
                f"{effective_system}\n\n{self._interpreter.system_message}"
            )

            # If history provided, prepend a conversation summary to the prompt
            if history:
                summary_lines = ["[Recent conversation context]"]
                for msg in history[-10:]:  # Last 10 messages to keep manageable
                    if not isinstance(msg, dict):
                        continue
                    role = str(msg.get("role", "user")).capitalize()
                    content = str(msg.get("content", "") or "")
                    if not content:
                        continue
                    if (
                        _looks_like_execute_noise(content)
                        or _looks_like_execute_payload_fragment(content)
                        or _looks_like_raw_command_leak(content)
                    ):
                        continue
                    content = re.sub(r"\s+", " ", content).strip()
                    if len(content) > 300:
                        content = content[:300] + "..."
                    summary_lines.append(f"{role}: {content}")
                summary_lines.append("[End of context]\n")
                message = "\n".join(summary_lines) + message

            # Use a queue to stream chunks from the sync thread to the async generator
            chunk_queue: asyncio.Queue = asyncio.Queue()

            def run_sync():
                """Run interpreter in a thread, push chunks to queue.

                Open Interpreter chunk types:
                - role: "assistant", type: "message" -> Text to show user
                - role: "assistant", type: "code" -> Code being written
                - role: "computer", type: "console", start: true -> Execution starting
                - role: "computer", type: "console", format: "output" -> Final output
                - role: "computer", type: "console", end: true -> Execution done
                """
                current_message = []
                current_language = None
                shown_running = False
                in_code_block = False
                current_code_parts: list[str] = []
                code_blocks_seen = 0
                command_counts: dict[str, int] = {}
                last_command_fingerprints: set[str] = set()
                recent_console_output = ""
                assistant_text_window = ""
                fence_markers_seen = 0
                saw_execution_complete = False
                failed_executions = 0
                structured_execute_done = False
                execute_payload_buffer = ""
                execute_capture_open = False
                max_code_blocks = max(
                    1, int(getattr(self.settings, "oi_max_code_blocks", DEFAULT_OI_MAX_CODE_BLOCKS))
                )
                repeat_limit = max(
                    1,
                    int(
                        getattr(
                            self.settings,
                            "oi_repeat_command_limit",
                            DEFAULT_OI_REPEAT_COMMAND_LIMIT,
                        )
                    ),
                )
                max_fix_retries = max(
                    0, int(getattr(self.settings, "oi_max_fix_retries", DEFAULT_OI_MAX_FIX_RETRIES))
                )
                stop_after_success = bool(getattr(self.settings, "oi_stop_after_success", True))
                fallback_attempted = False
                quota_state = {"detected": False, "text": ""}
                runtime_model_before_fallback = str(getattr(self._interpreter.llm, "model", "") or "")
                fallback_runtime_model = (
                    f"openai/{quota_fallback_model}"
                    if runtime_model_before_fallback.startswith("openai/")
                    else quota_fallback_model
                )

                def mark_quota_signal(text: str) -> None:
                    quota_state["detected"] = True
                    quota_state["text"] = str(text or quota_state.get("text") or "")

                if respond_mod is not None:

                    def _display_markdown_proxy(payload):
                        text = str(payload or "")
                        if _is_quota_or_rate_limit_message(text):
                            mark_quota_signal(text)
                        # Never emit raw Open Interpreter markdown blocks to terminal.
                        return None

                    respond_mod.display_markdown_message = _display_markdown_proxy

                def queue_chunk(payload: dict) -> None:
                    if isinstance(payload, dict) and payload.get("type") == "message":
                        content = str(payload.get("content", "") or "")
                        if (
                            _looks_like_execute_noise(content)
                            or _looks_like_execute_payload_fragment(content)
                            or _looks_like_raw_command_leak(content)
                        ):
                            return
                    asyncio.run_coroutine_threadsafe(chunk_queue.put(payload), loop)

                def trigger_stop(notice: str | None = None) -> None:
                    """Stop execution and reset interpreter promptly."""
                    if notice:
                        queue_chunk({"type": "message", "content": notice})
                    self._stop_flag = True
                    try:
                        self._interpreter.reset()
                    except Exception:
                        pass

                def stop_with_notice(arabic_msg: str, english_msg: str) -> tuple[bool, bool]:
                    trigger_stop(arabic_msg if is_arabic_request else english_msg)
                    return False, True

                def execute_structured_powershell(code: str, *, source: str) -> tuple[bool, bool]:
                    """Run one safe PowerShell command extracted from assistant text."""
                    nonlocal structured_execute_done, last_command_fingerprints
                    nonlocal execute_payload_buffer, execute_capture_open

                    if structured_execute_done:
                        return stop_with_notice(
                            "تم إيقاف الطلب: تم تكرار تنفيذ نفس الأمر.",
                            "Stopped: structured execute payload was repeated.",
                        )

                    fingerprints = _extract_command_fingerprints(code, allow_fallback=False)
                    allowed = bool(fingerprints) and all(
                        fp.startswith("start-process:")
                        or fp.startswith("get-process")
                        or fp == "tasklist"
                        or fp == "explorer"
                        for fp in fingerprints
                    )
                    if not allowed:
                        return stop_with_notice(
                            "تم إيقاف الطلب: الأمر غير مسموح ضمن قائمة الأمان.",
                            "Stopped structured execute payload: command not in safe allowlist.",
                        )

                    for fp in fingerprints:
                        count = command_counts.get(fp, 0) + 1
                        command_counts[fp] = count
                        if count > repeat_limit:
                            return stop_with_notice(
                                "تم إيقاف الطلب: تم اكتشاف تكرار نفس الأمر.",
                                "Stopped repeated execution: duplicate command pattern detected "
                                f"in execute payload (`{fp}`).",
                            )

                    last_command_fingerprints = fingerprints
                    queue_chunk(
                        {
                            "type": "tool_use",
                            "content": "جاري تنفيذ أمر PowerShell..." if is_arabic_request else "Running PowerShell...",
                            "metadata": {
                                "name": "run_powershell",
                                "input": {"code": code[:240], "source": source},
                            },
                        }
                    )
                    output = _run_powershell_once(code)
                    queue_chunk(
                        {
                            "type": "tool_result",
                            "content": "اكتمل تنفيذ أمر PowerShell" if is_arabic_request else "PowerShell execution completed",
                            "metadata": {"name": "run_powershell"},
                        }
                    )
                    summary = _summarize_structured_result(code, output, arabic=is_arabic_request)
                    if summary:
                        queue_chunk({"type": "message", "content": summary})

                    execute_payload_buffer = ""
                    execute_capture_open = False
                    structured_execute_done = True
                    trigger_stop(None)
                    return False, True

                def execute_structured_python(code: str, *, source: str) -> tuple[bool, bool]:
                    """Run one constrained Python desktop snippet from structured payload."""
                    nonlocal structured_execute_done
                    nonlocal execute_payload_buffer, execute_capture_open

                    if structured_execute_done:
                        return stop_with_notice(
                            "تم إيقاف الطلب: تم تكرار تنفيذ نفس الأمر.",
                            "Stopped: structured execute payload was repeated.",
                        )

                    fingerprints = _extract_command_fingerprints(code, allow_fallback=True)
                    for fp in fingerprints:
                        count = command_counts.get(fp, 0) + 1
                        command_counts[fp] = count
                        if count > repeat_limit:
                            return stop_with_notice(
                                "تم إيقاف الطلب: تم اكتشاف تكرار نفس الأمر.",
                                "Stopped repeated execution: duplicate Python command pattern detected.",
                            )

                    queue_chunk(
                        {
                            "type": "tool_use",
                            "content": "جاري تنفيذ أمر Python..." if is_arabic_request else "Running Python command...",
                            "metadata": {
                                "name": "run_python",
                                "input": {"code": code[:240], "source": source},
                            },
                        }
                    )
                    output = _run_python_desktop_once(code)
                    queue_chunk(
                        {
                            "type": "tool_result",
                            "content": "اكتمل تنفيذ أمر Python" if is_arabic_request else "Python command completed",
                            "metadata": {"name": "run_python"},
                        }
                    )
                    summary = _summarize_structured_result(code, output, arabic=is_arabic_request)
                    if summary:
                        queue_chunk({"type": "message", "content": summary})

                    execute_payload_buffer = ""
                    execute_capture_open = False
                    structured_execute_done = True
                    trigger_stop(None)
                    return False, True

                def track_message_commands(text: str) -> tuple[bool, bool]:
                    """Track command repetition from assistant text chunks."""
                    nonlocal fence_markers_seen
                    nonlocal execute_payload_buffer, execute_capture_open
                    if not text:
                        return True, False

                    maybe_execute = _is_execute_fragment(text) or _looks_like_execute_payload_fragment(text)
                    continuation = execute_capture_open and _looks_like_execute_continuation(text)
                    noisy_execute = _looks_like_execute_noise(text) or _looks_like_raw_command_leak(text)
                    if maybe_execute:
                        execute_capture_open = True

                    if maybe_execute or continuation or noisy_execute:
                        execute_payload_buffer = (execute_payload_buffer + text).strip()[-10000:]
                    elif execute_capture_open:
                        # We started capturing an execute payload but chunk stream switched
                        # back to plain text; clear stale buffer to avoid leaking junk output.
                        execute_capture_open = False
                        execute_payload_buffer = ""

                    payload = _extract_execute_payload(text) or _extract_execute_payload(
                        execute_payload_buffer
                    )
                    if payload is not None:
                        language, code = payload

                        if language in {"", "powershell", "pwsh", "ps1"}:
                            return execute_structured_powershell(code, source="structured_payload")
                        if language in {"python", "py"}:
                            return execute_structured_python(code, source="structured_payload")

                        return stop_with_notice(
                            "تم إيقاف الطلب: لغة التنفيذ غير مدعومة في هذا الوضع.",
                            "Stopped structured execute payload: "
                            f"unsupported language '{language}'.",
                        )

                    fallback_code = _extract_powershell_code_fallback(execute_payload_buffer)
                    if fallback_code:
                        return execute_structured_powershell(
                            fallback_code, source="malformed_payload_fallback"
                        )
                    fallback_py = _extract_python_code_fallback(execute_payload_buffer)
                    if fallback_py:
                        return execute_structured_python(
                            fallback_py, source="malformed_payload_fallback"
                        )

                    if maybe_execute or continuation or noisy_execute:
                        # Wait for more chunks and suppress noisy JSON fragments.
                        return True, True

                    if (
                        _looks_like_execute_noise(text)
                        or _looks_like_execute_payload_fragment(text)
                        or _looks_like_raw_command_leak(text)
                    ):
                        # Never stream malformed execute JSON snippets to the user.
                        return True, True

                    fence_markers_seen += text.count("```")
                    if fence_markers_seen > (max_code_blocks * 2 + 2):
                        trigger_stop(
                            "تم إيقاف المخرجات المزعجة: تم اكتشاف عدد كبير من كتل الكود."
                            if is_arabic_request
                            else "Stopped noisy output: too many markdown code fences were emitted."
                        )
                        return False, False

                    fingerprints = _extract_command_fingerprints(text, allow_fallback=False)
                    for fp in fingerprints:
                        count = command_counts.get(fp, 0) + 1
                        command_counts[fp] = count
                        if count > repeat_limit:
                            trigger_stop(
                                "تم إيقاف التنفيذ: تم اكتشاف تكرار نفس الأمر."
                                if is_arabic_request
                                else "Stopped repeated execution: duplicate command pattern "
                                f"detected in assistant output (`{fp}`)."
                            )
                            return False, False
                    return True, False

                def finalize_code_block() -> bool:
                    """Finalize a code block and stop if repeated too many times."""
                    nonlocal in_code_block, last_command_fingerprints
                    if not in_code_block:
                        return True
                    in_code_block = False
                    command = "".join(current_code_parts).strip()
                    current_code_parts.clear()
                    if not command:
                        last_command_fingerprints = set()
                        return True

                    if _is_execute_fragment(command) or _looks_like_execute_noise(command):
                        payload = _extract_execute_payload(command)
                        if payload is not None:
                            language, code = payload
                            if language in {"", "powershell", "pwsh", "ps1"}:
                                should_continue, _ = execute_structured_powershell(
                                    code, source="code_block_payload"
                                )
                                return should_continue
                            if language in {"python", "py"}:
                                should_continue, _ = execute_structured_python(
                                    code, source="code_block_payload"
                                )
                                return should_continue
                        fallback_code = _extract_powershell_code_fallback(command)
                        if fallback_code:
                            should_continue, _ = execute_structured_powershell(
                                fallback_code, source="code_block_fallback"
                            )
                            return should_continue
                        fallback_py = _extract_python_code_fallback(command)
                        if fallback_py:
                            should_continue, _ = execute_structured_python(
                                fallback_py, source="code_block_fallback"
                            )
                            return should_continue

                    fingerprints = _extract_command_fingerprints(command)
                    last_command_fingerprints = fingerprints
                    for fp in fingerprints:
                        count = command_counts.get(fp, 0) + 1
                        command_counts[fp] = count
                        if count > repeat_limit:
                            trigger_stop(
                                "⚠️ Stopped repeated execution: duplicate command pattern "
                                f"detected (`{fp}`)."
                            )
                            return False
                    return True

                try:
                    for chunk in self._interpreter.chat(message, stream=True):
                        if self._stop_flag:
                            break

                        if isinstance(chunk, dict):
                            chunk_role = chunk.get("role", "")
                            chunk_type = chunk.get("type", "")
                            content = chunk.get("content", "")
                            chunk_format = chunk.get("format", "")
                            is_start = chunk.get("start", False)
                            is_end = chunk.get("end", False)

                            # Handle computer/console chunks - emit tool events for Activity
                            if chunk_role == "computer":
                                if chunk_type == "console":
                                    if is_start and not finalize_code_block():
                                        break
                                    if is_start:
                                        recent_console_output = ""
                                    if is_start and current_language and not shown_running:
                                        # Emit tool_use event for Activity panel
                                        lang_display = current_language.title()
                                        queue_chunk(
                                            {
                                                "type": "tool_use",
                                                "content": f"Running {lang_display}...",
                                                "metadata": {
                                                    "name": f"run_{current_language}",
                                                    "input": {},
                                                },
                                            }
                                        )
                                        shown_running = True
                                    elif is_end:
                                        # Emit tool_result event for Activity panel
                                        lang_display = (
                                            current_language.title() if current_language else "Code"
                                        )
                                        queue_chunk(
                                            {
                                                "type": "tool_result",
                                                "content": f"{lang_display} execution completed",
                                                "metadata": {
                                                    "name": f"run_{current_language or 'code'}"
                                                },
                                            }
                                        )
                                        saw_execution_complete = True
                                        if _looks_like_error_output(recent_console_output):
                                            failed_executions += 1
                                            if failed_executions > max_fix_retries:
                                                trigger_stop(
                                                    "⚠️ Stopped after repeated execution "
                                                    "errors (retry limit reached)."
                                                )
                                                break
                                        elif stop_after_success and (
                                            (
                                                _looks_like_process_snapshot(recent_console_output)
                                                and any(
                                                    fp.startswith("get-process")
                                                    or fp == "tasklist"
                                                    for fp in last_command_fingerprints
                                                )
                                            )
                                            or _contains_done_signal(assistant_text_window)
                                        ):
                                            trigger_stop("✅ Auto-stopped after a successful result.")
                                            break
                                        # Reset for next code block
                                        shown_running = False
                                    elif content:
                                        recent_console_output = (
                                            f"{recent_console_output}\n{content}"
                                        ).strip()[-4000:]
                                    # Skip verbose active_line, intermediate output
                                continue

                            # Handle assistant chunks
                            if chunk_role == "assistant":
                                if chunk_type == "code":
                                    if not in_code_block:
                                        in_code_block = True
                                        code_blocks_seen += 1
                                        if code_blocks_seen > max_code_blocks:
                                            trigger_stop(
                                                "⚠️ Stopped execution after reaching the "
                                                f"limit of {max_code_blocks} code blocks "
                                                "for this message."
                                            )
                                            break
                                    if content:
                                        current_code_parts.append(str(content))
                                    # Capture language for progress indicator
                                    current_language = chunk_format or "code"
                                    # Flush any pending message
                                    if current_message:
                                        queue_chunk(
                                            {
                                                "type": "message",
                                                "content": "".join(current_message),
                                            }
                                        )
                                        current_message = []
                                    # Don't show raw code fragments
                                elif chunk_type == "message" and content:
                                    if not finalize_code_block():
                                        break
                                    should_continue, suppress_output = track_message_commands(
                                        str(content)
                                    )
                                    if not should_continue:
                                        break
                                    if (
                                        _looks_like_execute_noise(str(content))
                                        or _looks_like_execute_payload_fragment(str(content))
                                        or _looks_like_raw_command_leak(str(content))
                                    ):
                                        suppress_output = True
                                    # Stream message chunks
                                    if not suppress_output:
                                        queue_chunk({"type": "message", "content": content})
                                        assistant_text_window = (
                                            f"{assistant_text_window} {content}"
                                        ).strip()[-1200:]
                                    if saw_execution_complete and _contains_done_signal(
                                        assistant_text_window
                                    ):
                                        trigger_stop("✅ Auto-stopped after a successful result.")
                                        break
                        elif isinstance(chunk, str) and chunk:
                            if not finalize_code_block():
                                break
                            should_continue, suppress_output = track_message_commands(chunk)
                            if not should_continue:
                                break
                            if (
                                _looks_like_execute_noise(chunk)
                                or _looks_like_execute_payload_fragment(chunk)
                                or _looks_like_raw_command_leak(chunk)
                            ):
                                suppress_output = True
                            if not suppress_output:
                                current_message.append(chunk)
                                assistant_text_window = (
                                    f"{assistant_text_window} {chunk}"
                                ).strip()[-1200:]
                            if saw_execution_complete and _contains_done_signal(
                                assistant_text_window
                            ):
                                trigger_stop("✅ Auto-stopped after a successful result.")
                                break

                    if (
                        not self._stop_flag
                        and execute_payload_buffer
                        and not structured_execute_done
                    ):
                        fallback_code = _extract_powershell_code_fallback(execute_payload_buffer)
                        if fallback_code:
                            execute_structured_powershell(
                                fallback_code, source="end_of_stream_fallback"
                            )
                        else:
                            fallback_py = _extract_python_code_fallback(execute_payload_buffer)
                            if fallback_py:
                                execute_structured_python(
                                    fallback_py, source="end_of_stream_fallback"
                                )
                            else:
                                queue_chunk(
                                    {
                                        "type": "message",
                                        "content": (
                                            "تعذر تفسير أمر التنفيذ. أعد صياغة الطلب بشكل مباشر."
                                            if is_arabic_request
                                            else "Could not parse execution command. Please rephrase your request."
                                        ),
                                    }
                                )
                                execute_payload_buffer = ""

                    if not self._stop_flag:
                        finalize_code_block()

                    # Flush remaining message
                    if current_message:
                        queue_chunk({"type": "message", "content": "".join(current_message)})

                    if _should_retry_with_gemini_fallback(
                        resolved_provider,
                        fallback_attempted,
                        str(quota_state.get("text", "")),
                    ):
                        fallback_attempted = True
                        self._stop_flag = False
                        quota_state = {"detected": False, "text": ""}
                        queue_chunk(
                            {
                                "type": "status",
                                "content": "gemini_quota_fallback",
                                "metadata": {
                                    "from_model": resolved_model or runtime_model_before_fallback,
                                    "to_model": quota_fallback_model,
                                },
                            }
                        )
                        queue_chunk(
                            {
                                "type": "message",
                                "content": (
                                    f"🔁 تم تجاوز حد Gemini. أعيد المحاولة مرة واحدة عبر `{quota_fallback_model}`..."
                                    if is_arabic_request
                                    else f"🔁 Gemini quota reached. Retrying once with `{quota_fallback_model}`..."
                                ),
                            }
                        )
                        try:
                            self._interpreter.llm.model = fallback_runtime_model
                            for fallback_chunk in self._interpreter.chat(message, stream=True):
                                if self._stop_flag:
                                    break
                                if isinstance(fallback_chunk, dict):
                                    chunk_role = fallback_chunk.get("role", "")
                                    chunk_type = fallback_chunk.get("type", "")
                                    content = str(fallback_chunk.get("content", "") or "")
                                    if chunk_role == "assistant" and chunk_type == "message" and content:
                                        if not (
                                            _looks_like_execute_noise(content)
                                            or _looks_like_execute_payload_fragment(content)
                                            or _looks_like_raw_command_leak(content)
                                        ):
                                            queue_chunk({"type": "message", "content": content})
                                elif isinstance(fallback_chunk, str) and fallback_chunk:
                                    content = str(fallback_chunk)
                                    if not (
                                        _looks_like_execute_noise(content)
                                        or _looks_like_execute_payload_fragment(content)
                                        or _looks_like_raw_command_leak(content)
                                    ):
                                        queue_chunk({"type": "message", "content": content})
                        except Exception as fallback_err:
                            fallback_text = str(fallback_err)
                            if _is_quota_or_rate_limit_message(fallback_text):
                                mark_quota_signal(fallback_text)
                            else:
                                queue_chunk(
                                    {
                                        "type": "error",
                                        "content": f"Agent error: {fallback_text}",
                                    }
                                )
                        finally:
                            try:
                                self._interpreter.llm.model = runtime_model_before_fallback
                            except Exception:
                                pass

                    if quota_state.get("detected"):
                        queue_chunk(
                            {
                                "type": "error",
                                "content": _build_quota_error_message(
                                    resolved_provider,
                                    resolved_model or runtime_model_before_fallback,
                                    quota_fallback_model,
                                    arabic=is_arabic_request,
                                    fallback_attempted=fallback_attempted,
                                ),
                            }
                        )
                except Exception as e:
                    error_text = str(e)
                    if _is_quota_or_rate_limit_message(error_text):
                        mark_quota_signal(error_text)
                    else:
                        queue_chunk({"type": "error", "content": f"Agent error: {error_text}"})
                finally:
                    # Signal completion
                    asyncio.run_coroutine_threadsafe(chunk_queue.put(None), loop)

            try:
                loop = asyncio.get_event_loop()

                # Start the sync function in a thread
                executor_future = loop.run_in_executor(None, run_sync)

                # Yield chunks as they arrive
                while True:
                    try:
                        chunk = await asyncio.wait_for(chunk_queue.get(), timeout=60.0)
                        if chunk is None:  # End signal
                            break
                        if isinstance(chunk, dict) and chunk.get("type") == "message":
                            content = str(chunk.get("content", "") or "")
                            if (
                                _looks_like_execute_noise(content)
                                or _looks_like_execute_payload_fragment(content)
                                or _looks_like_raw_command_leak(content)
                            ):
                                continue
                        yield chunk
                    except TimeoutError:
                        yield {
                            "type": "status",
                            "content": "backend_processing",
                            "metadata": {"heartbeat_seconds": 60},
                        }

                # Wait for executor to finish
                await executor_future

            except Exception as e:
                logger.error(f"Open Interpreter error: {e}")
                yield {"type": "error", "content": f"❌ Agent error: {str(e)}"}
            finally:
                self._interpreter.system_message = original_system_message
                if respond_mod is not None:
                    try:
                        if original_display_markdown is not None:
                            respond_mod.display_markdown_message = original_display_markdown
                        elif hasattr(respond_mod, "display_markdown_message"):
                            delattr(respond_mod, "display_markdown_message")
                    except Exception:
                        pass
        finally:
            if acquired:
                self._semaphore.release()

    async def stop(self) -> None:
        """Stop the agent execution."""
        self._stop_flag = True
        if self._interpreter:
            try:
                self._interpreter.reset()
            except Exception:
                pass
