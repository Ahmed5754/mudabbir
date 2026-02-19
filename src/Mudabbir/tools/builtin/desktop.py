"""Desktop automation and system control helpers for Windows."""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import time
import wave
from pathlib import Path
from typing import Any

from Mudabbir.bus.media import get_media_dir
from Mudabbir.tools.protocol import BaseTool


class _NoisyDisplayWarningFilter(logging.Filter):
    """Drop known noisy EDID warnings from third-party brightness discovery."""

    _patterns = (
        "exception parsing edid str",
        "edidparseerror",
        "detected but not present in monitor_uids",
        "maybe it is asleep",
    )

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            message = str(record.getMessage() or "").lower()
        except Exception:
            message = ""
        if any(pattern in message for pattern in self._patterns):
            return False
        return True


_DISPLAY_WARNING_FILTER_INSTALLED = False


def _install_display_warning_filter_once() -> None:
    global _DISPLAY_WARNING_FILTER_INSTALLED
    if _DISPLAY_WARNING_FILTER_INSTALLED:
        return
    filt = _NoisyDisplayWarningFilter()
    for logger_name in (
        "screen_brightness_control",
        "screen_brightness_control.windows",
        "screen_brightness_control.helpers",
    ):
        logging.getLogger(logger_name).addFilter(filt)
    logging.getLogger().addFilter(filt)
    _DISPLAY_WARNING_FILTER_INSTALLED = True


def _json(data: dict[str, Any] | list[Any]) -> str:
    return json.dumps(data, ensure_ascii=False)


def _clamp(value: int | float, minimum: int, maximum: int) -> int:
    return max(minimum, min(maximum, int(value)))


def _timestamp_id() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


def _run_powershell(command: str, timeout: int = 15) -> tuple[bool, str]:
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
    except Exception as e:
        return False, str(e)

    out = (proc.stdout or "").strip()
    err = (proc.stderr or "").strip()
    if proc.returncode == 0:
        return True, out
    if err:
        return False, err
    if out:
        return False, out
    return False, f"PowerShell exited with code {proc.returncode}"


def _iter_search_roots() -> list[Path]:
    home = Path.home()
    roots = [
        home / "Desktop",
        home / "Documents",
        home / "Downloads",
        home / "Pictures",
        home / "Videos",
    ]
    existing = [p for p in roots if p.exists() and p.is_dir()]
    return existing or [home]


def _contains_text(haystack: str, needle: str) -> bool:
    return needle.casefold() in haystack.casefold()


def _normalize_query(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _serialize_windows(
    *,
    include_untitled: bool = True,
    limit: int = 200,
) -> list[dict[str, Any]]:
    windows: list[dict[str, Any]] = []
    try:
        import psutil
        import win32gui
        import win32process

        def enum_handler(hwnd: int, _: Any) -> None:
            if len(windows) >= limit:
                return
            if not win32gui.IsWindowVisible(hwnd):
                return

            title = (win32gui.GetWindowText(hwnd) or "").strip()
            if not include_untitled and not title:
                return

            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            width = max(0, right - left)
            height = max(0, bottom - top)
            if width < 40 or height < 40:
                return

            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process_name = ""
            try:
                process_name = psutil.Process(pid).name()
            except Exception:
                process_name = ""

            windows.append(
                {
                    "hwnd": int(hwnd),
                    "pid": int(pid),
                    "process_name": process_name,
                    "title": title,
                    "left": int(left),
                    "top": int(top),
                    "width": int(width),
                    "height": int(height),
                }
            )

        win32gui.EnumWindows(enum_handler, None)
    except Exception:
        try:
            import pygetwindow as gw

            for win in gw.getAllWindows():
                if len(windows) >= limit:
                    break
                title = str(getattr(win, "title", "") or "").strip()
                if not include_untitled and not title:
                    continue
                width = int(getattr(win, "width", 0) or 0)
                height = int(getattr(win, "height", 0) or 0)
                if width < 40 or height < 40:
                    continue
                windows.append(
                    {
                        "hwnd": 0,
                        "pid": 0,
                        "process_name": "",
                        "title": title,
                        "left": int(getattr(win, "left", 0) or 0),
                        "top": int(getattr(win, "top", 0) or 0),
                        "width": width,
                        "height": height,
                    }
                )
        except Exception:
            return []

    windows.sort(key=lambda item: (item.get("title") or "").casefold())
    return windows[:limit]


class DesktopTool(BaseTool):
    BLOCKED_AUTOMATION_PROCESSES = {
        "keepass.exe",
        "1password.exe",
        "lastpass.exe",
        "bitwarden.exe",
        "authy desktop.exe",
        "secpol.msc",
        "gpedit.msc",
        "regedit.exe",
    }
    BLOCKED_HOTKEY_COMBINATIONS = {
        frozenset({"ctrl", "alt", "delete"}),
        frozenset({"win", "r"}),
        frozenset({"win", "x"}),
    }
    BLOCKED_WINDOW_CLASSES = {
        "#32770",
        "credential dialog",
        "credential dialog xaml host",
        "windows security",
    }

    @property
    def name(self) -> str:
        return "desktop"

    @property
    def description(self) -> str:
        return "Desktop automation, app control, media capture, and Windows system actions."

    @property
    def trust_level(self) -> str:
        return "high"

    @property
    def parameters(self) -> dict[str, Any]:
        return {
            "type": "object",
            "properties": {
                "action": {"type": "string"},
                "mode": {"type": "string"},
                "query": {"type": "string"},
                "process_name": {"type": "string"},
                "x": {"type": "number"},
                "y": {"type": "number"},
                "button": {"type": "string"},
                "clicks": {"type": "integer"},
                "keys": {"type": "array", "items": {"type": "string"}},
                "key": {"type": "string"},
                "level": {"type": "integer"},
                "delta": {"type": "integer"},
                "max_results": {"type": "integer"},
                "seconds": {"type": "number"},
                "frames": {"type": "integer"},
                "interval_sec": {"type": "number"},
                "camera_index": {"type": "integer"},
                "only_windowed": {"type": "boolean"},
                "include_untitled": {"type": "boolean"},
                "limit": {"type": "integer"},
                "page": {"type": "string"},
                "force": {"type": "boolean"},
                "window_title": {"type": "string"},
                "control_name": {"type": "string"},
                "auto_id": {"type": "string"},
                "control_type": {"type": "string"},
                "text": {"type": "string"},
                "index": {"type": "integer"},
                "press_enter": {"type": "boolean"},
                "timeout_sec": {"type": "number"},
                "duration": {"type": "number"},
                "interval": {"type": "number"},
                "interaction": {"type": "string"},
            },
            "required": ["action"],
        }

    async def execute(self, action: str, **params: Any) -> str:
        action_normalized = (action or "").strip().lower()
        if not action_normalized:
            return self._error("Missing action")

        try:
            if action_normalized == "screen_snapshot":
                return self._screen_snapshot()
            if action_normalized == "screen_watch":
                return self._screen_watch(
                    frames=int(params.get("frames", 3) or 3),
                    interval_sec=float(params.get("interval_sec", 0.7) or 0.7),
                )
            if action_normalized == "desktop_overview":
                return self._desktop_overview()
            if action_normalized == "battery_status":
                return self._battery_status()
            if action_normalized == "list_windows":
                return self._list_windows(
                    include_untitled=bool(params.get("include_untitled", True)),
                    limit=int(params.get("limit", 120) or 120),
                )
            if action_normalized == "focus_window":
                return self._focus_window(
                    window_title=str(params.get("window_title", "") or params.get("query", "") or ""),
                    process_name=str(params.get("process_name", "") or ""),
                    timeout_sec=float(params.get("timeout_sec", 4.0) or 4.0),
                )
            if action_normalized == "ui_list_controls":
                return self._ui_list_controls(
                    window_title=str(params.get("window_title", "") or params.get("query", "") or ""),
                    process_name=str(params.get("process_name", "") or ""),
                    control_name=str(params.get("control_name", "") or params.get("query", "") or ""),
                    control_type=str(params.get("control_type", "") or ""),
                    max_results=int(params.get("max_results", 40) or 40),
                    timeout_sec=float(params.get("timeout_sec", 4.0) or 4.0),
                )
            if action_normalized == "ui_click":
                return self._ui_click(
                    window_title=str(params.get("window_title", "") or params.get("query", "") or ""),
                    process_name=str(params.get("process_name", "") or ""),
                    control_name=str(params.get("control_name", "") or ""),
                    auto_id=str(params.get("auto_id", "") or ""),
                    control_type=str(params.get("control_type", "") or ""),
                    index=int(params.get("index", 0) or 0),
                    timeout_sec=float(params.get("timeout_sec", 4.0) or 4.0),
                )
            if action_normalized == "ui_set_text":
                return self._ui_set_text(
                    text=str(params.get("text", "") or ""),
                    window_title=str(params.get("window_title", "") or params.get("query", "") or ""),
                    process_name=str(params.get("process_name", "") or ""),
                    control_name=str(params.get("control_name", "") or ""),
                    auto_id=str(params.get("auto_id", "") or ""),
                    control_type=str(params.get("control_type", "") or ""),
                    index=int(params.get("index", 0) or 0),
                    press_enter=bool(params.get("press_enter", False)),
                    timeout_sec=float(params.get("timeout_sec", 4.0) or 4.0),
                )
            if action_normalized in {"ui_target", "ui_move_to"}:
                interaction = str(params.get("interaction", "") or "").strip().lower()
                if action_normalized == "ui_move_to" and not interaction:
                    interaction = "move"
                return self._ui_target(
                    window_title=str(params.get("window_title", "") or params.get("query", "") or ""),
                    process_name=str(params.get("process_name", "") or ""),
                    control_name=str(params.get("control_name", "") or params.get("query", "") or ""),
                    auto_id=str(params.get("auto_id", "") or ""),
                    control_type=str(params.get("control_type", "") or ""),
                    index=int(params.get("index", 0) or 0),
                    interaction=interaction or "move",
                    duration=float(params.get("duration", 0.25) or 0.25),
                    timeout_sec=float(params.get("timeout_sec", 5.0) or 5.0),
                )
            if action_normalized == "list_processes":
                return self._list_processes(
                    only_windowed=bool(params.get("only_windowed", False)),
                    max_results=int(params.get("max_results", 25) or 25),
                )
            if action_normalized == "close_app":
                return self._close_app(
                    process_name=str(params.get("process_name", "") or ""),
                    force=bool(params.get("force", True)),
                )
            if action_normalized == "search_start_apps":
                return self._search_start_apps(
                    query=str(params.get("query", "") or ""),
                    max_results=int(params.get("max_results", 10) or 10),
                )
            if action_normalized == "launch_start_app":
                return self._launch_start_app(query=str(params.get("query", "") or ""))
            if action_normalized == "search_files":
                return self._search_files(
                    query=str(params.get("query", "") or ""),
                    max_results=int(params.get("max_results", 20) or 20),
                )
            if action_normalized == "move_mouse_to_desktop_file":
                return self._move_mouse_to_desktop_file(
                    query=str(params.get("query", "") or ""),
                    duration=float(params.get("duration", 0.25) or 0.25),
                    timeout_sec=float(params.get("timeout_sec", 4.0) or 4.0),
                )
            if action_normalized == "list_installed_apps":
                return self._list_installed_apps(
                    query=str(params.get("query", "") or ""),
                    max_results=int(params.get("max_results", 25) or 25),
                )
            if action_normalized == "volume":
                return self._volume_control(
                    mode=str(params.get("mode", "get") or "get"),
                    level=params.get("level"),
                    delta=params.get("delta"),
                )
            if action_normalized == "brightness":
                return self._brightness_control(
                    mode=str(params.get("mode", "get") or "get"),
                    level=params.get("level"),
                    delta=params.get("delta"),
                )
            if action_normalized == "media_control":
                return self._media_control(mode=str(params.get("mode", "play_pause") or "play_pause"))
            if action_normalized == "mouse_move":
                return self._mouse_move(
                    x=int(params.get("x", 0) or 0),
                    y=int(params.get("y", 0) or 0),
                    duration=float(params.get("duration", 0.2) or 0.2),
                )
            if action_normalized == "click":
                return self._click(
                    x=params.get("x"),
                    y=params.get("y"),
                    button=str(params.get("button", "left") or "left"),
                    clicks=int(params.get("clicks", 1) or 1),
                )
            if action_normalized == "press_key":
                return self._press_key(key=str(params.get("key", "") or ""))
            if action_normalized == "type_text":
                return self._type_text(
                    text=str(params.get("text", "") or ""),
                    press_enter=bool(params.get("press_enter", False)),
                    interval=float(params.get("interval", 0.01) or 0.01),
                )
            if action_normalized == "hotkey":
                keys = params.get("keys") or []
                keys = [str(k) for k in keys if str(k).strip()]
                return self._hotkey(keys=keys)
            if action_normalized == "camera_snapshot":
                return self._camera_snapshot(camera_index=int(params.get("camera_index", 0) or 0))
            if action_normalized == "microphone_record":
                return self._microphone_record(seconds=float(params.get("seconds", 3.0) or 3.0))
            if action_normalized == "open_settings_page":
                return self._open_settings_page(page=str(params.get("page", "") or ""))
            if action_normalized == "bluetooth_control":
                return self._bluetooth_control(mode=str(params.get("mode", "open_settings") or "open_settings"))
        except Exception as e:
            return self._error(str(e))

        return self._error(f"Unknown desktop action: {action_normalized}")

    def _screen_snapshot(self) -> str:
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            media_dir = get_media_dir()
            path = media_dir / f"screen_{_timestamp_id()}.png"
            screenshot = pyautogui.screenshot()
            screenshot.save(path)
            return f"Screenshot saved to {path}"
        except Exception as e:
            return self._error(f"screenshot failed: {e}")

    def _screen_watch(self, frames: int, interval_sec: float) -> str:
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            frames = _clamp(frames, 1, 12)
            interval = max(0.05, min(5.0, float(interval_sec)))
            media_dir = get_media_dir()
            paths: list[str] = []
            for idx in range(frames):
                shot = pyautogui.screenshot()
                path = media_dir / f"watch_{_timestamp_id()}_{idx + 1}.png"
                shot.save(path)
                paths.append(str(path))
                if idx < frames - 1:
                    time.sleep(interval)
            return _json(
                {
                    "captured_frames": len(paths),
                    "interval_sec": interval,
                    "paths": paths,
                }
            )
        except Exception as e:
            return self._error(f"screen watch failed: {e}")

    def _desktop_overview(self) -> str:
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            size = pyautogui.size()
            pos = pyautogui.position()
            windows = _serialize_windows(include_untitled=False, limit=40)
            return _json(
                {
                    "screen": {"width": int(size.width), "height": int(size.height)},
                    "mouse": {"x": int(pos.x), "y": int(pos.y)},
                    "windows": {"count": len(windows), "items": windows[:20]},
                    "timestamp": time.time(),
                }
            )
        except Exception as e:
            return self._error(f"desktop overview failed: {e}")

    def _battery_status(self) -> str:
        """Return battery status if available."""
        try:
            import psutil

            battery = psutil.sensors_battery()
            if battery is not None:
                secs_left_raw = getattr(battery, "secsleft", None)
                secs_left: int | None
                if secs_left_raw is None:
                    secs_left = None
                else:
                    try:
                        value = int(secs_left_raw)
                    except Exception:
                        value = -2
                    if value < 0 or value in {psutil.POWER_TIME_UNKNOWN, psutil.POWER_TIME_UNLIMITED}:
                        secs_left = None
                    else:
                        secs_left = max(0, value)

                return _json(
                    {
                        "available": True,
                        "percent": max(0.0, min(100.0, float(getattr(battery, "percent", 0.0) or 0.0))),
                        "plugged": bool(getattr(battery, "power_plugged", False)),
                        "secs_left": secs_left,
                        "source": "psutil",
                    }
                )
        except Exception:
            pass

        # Fallback for systems where psutil battery API is unavailable.
        cmd = (
            "Get-CimInstance Win32_Battery | Select-Object -First 1 "
            "EstimatedChargeRemaining,BatteryStatus,Name | ConvertTo-Json -Compress"
        )
        ok, out = _run_powershell(cmd, timeout=12)
        if ok and out and out.strip() not in {"null", ""}:
            try:
                data = json.loads(out)
                if isinstance(data, dict):
                    percent_raw = data.get("EstimatedChargeRemaining")
                    percent = None
                    try:
                        if percent_raw is not None:
                            percent = max(0.0, min(100.0, float(percent_raw)))
                    except Exception:
                        percent = None
                    status = int(data.get("BatteryStatus", 0) or 0)
                    plugged = status in {2, 6, 7, 8, 9, 11}
                    return _json(
                        {
                            "available": percent is not None,
                            "percent": percent,
                            "plugged": plugged,
                            "secs_left": None,
                            "status_code": status,
                            "name": str(data.get("Name", "") or "").strip(),
                            "source": "win32_battery",
                        }
                    )
            except Exception:
                pass

        return _json(
            {
                "available": False,
                "percent": None,
                "plugged": None,
                "secs_left": None,
                "message": "Battery information is not available on this machine/session.",
            }
        )

    def _list_windows(self, include_untitled: bool, limit: int) -> str:
        windows = _serialize_windows(include_untitled=include_untitled, limit=max(1, limit))
        return _json({"count": len(windows), "windows": windows})

    def _list_processes(self, only_windowed: bool, max_results: int) -> str:
        max_results = _clamp(max_results, 1, 200)
        if only_windowed:
            windows = _serialize_windows(include_untitled=False, limit=max_results * 3)
            rows: list[dict[str, Any]] = []
            seen: set[tuple[int, str]] = set()
            for w in windows:
                name = str(w.get("process_name") or "").strip()
                title = str(w.get("title") or "").strip()
                pid = int(w.get("pid") or 0)
                if not title:
                    continue
                key = (pid, title.casefold())
                if key in seen:
                    continue
                seen.add(key)
                rows.append({"Name": name or "unknown", "Id": pid, "MainWindowTitle": title})
                if len(rows) >= max_results:
                    break
            return _json({"count": len(rows), "processes": rows})

        try:
            import psutil

            rows = []
            for p in psutil.process_iter(["pid", "name", "memory_info"]):
                try:
                    mem = int(getattr(p.info.get("memory_info"), "rss", 0) or 0)
                    rows.append(
                        {
                            "Name": p.info.get("name") or "unknown",
                            "Id": int(p.info.get("pid") or 0),
                            "MemoryMB": round(mem / (1024 * 1024), 1),
                            "MainWindowTitle": "",
                        }
                    )
                except Exception:
                    continue
            rows.sort(key=lambda x: float(x.get("MemoryMB") or 0), reverse=True)
            return _json({"count": min(len(rows), max_results), "processes": rows[:max_results]})
        except Exception as e:
            return self._error(f"list_processes failed: {e}")

    def _close_app(self, process_name: str, force: bool) -> str:
        query = _normalize_query(process_name)
        if not query:
            return self._error("process_name is required")
        try:
            import psutil

            closed = 0
            pids: list[int] = []
            query_cf = query.casefold()
            query_noext = query_cf[:-4] if query_cf.endswith(".exe") else query_cf
            terms = [t for t in re.split(r"\s+", query_noext) if t]
            current_pid = os.getpid()
            for p in psutil.process_iter(["pid", "name", "exe", "cmdline"]):
                try:
                    pid = int(p.info.get("pid") or 0)
                    if pid <= 0 or pid == current_pid:
                        continue
                    name = str(p.info.get("name") or "").casefold()
                    exe = str(p.info.get("exe") or "")
                    exe_base = Path(exe).name.casefold() if exe else ""
                    name_noext = name[:-4] if name.endswith(".exe") else name
                    exe_noext = exe_base[:-4] if exe_base.endswith(".exe") else exe_base
                    candidates = [name, exe_base, name_noext, exe_noext]

                    matched = any(query_noext and query_noext in c for c in candidates if c)
                    if not matched and terms:
                        joined = " ".join([c for c in (name_noext, exe_noext) if c])
                        matched = all(t in joined for t in terms)
                    if not matched:
                        continue

                    if force:
                        p.kill()
                    else:
                        p.terminate()
                    closed += 1
                    pids.append(pid)
                except Exception:
                    continue
            return _json({"closed": closed, "matched_pids": pids})
        except Exception as e:
            return self._error(f"close_app failed: {e}")

    def _search_start_apps(self, query: str, max_results: int) -> str:
        max_results = _clamp(max_results, 1, 2000)
        query_norm = _normalize_query(query)
        query_cf = query_norm.casefold()
        query_terms = {query_cf} if query_cf else set()
        if query_cf and (
            "telegram" in query_cf
            or "تلجرام" in query_cf
            or "تيليجرام" in query_cf
            or "تليجرام" in query_cf
        ):
            query_terms.update({"telegram", "unigram", "تلجرام", "تيليجرام", "تليجرام"})

        command = "Get-StartApps | Sort-Object Name | ConvertTo-Json -Depth 4 -Compress"
        ok, output = _run_powershell(command, timeout=20)
        if not ok:
            return self._error(f"search_start_apps failed: {output}")
        try:
            parsed = json.loads(output) if output else []
            if isinstance(parsed, dict):
                parsed = [parsed]
            apps = []
            for item in parsed:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("Name", "") or "").strip()
                app_id = str(item.get("AppID", "") or "").strip()
                if not name:
                    continue
                if query_terms:
                    name_cf = name.casefold()
                    if not any(term in name_cf for term in query_terms if term):
                        continue
                apps.append({"Name": name, "AppID": app_id})
                if len(apps) >= max_results:
                    break
            return _json({"count": len(apps), "apps": apps})
        except Exception as e:
            return self._error(f"search_start_apps parse failed: {e}")

    @staticmethod
    def _parse_json_dict(raw: str) -> dict[str, Any] | None:
        try:
            parsed = json.loads(raw)
            return parsed if isinstance(parsed, dict) else None
        except Exception:
            return None

    def _try_focus_app(
        self,
        *,
        title_query: str = "",
        process_query: str = "",
        timeout_sec: float = 1.5,
    ) -> dict[str, Any] | None:
        raw = self._focus_window(
            window_title=str(title_query or ""),
            process_name=str(process_query or ""),
            timeout_sec=max(0.4, min(8.0, float(timeout_sec))),
        )
        if not isinstance(raw, str) or raw.lower().startswith("error:"):
            return None
        payload = self._parse_json_dict(raw)
        if isinstance(payload, dict) and payload.get("focused"):
            return payload
        return None

    def _launch_start_app(self, query: str) -> str:
        query_norm = _normalize_query(query)
        if not query_norm:
            return self._error("query is required")

        query_cf = query_norm.casefold()

        def _add_unique(values: list[str], extra: list[str]) -> None:
            for item in extra:
                value = _normalize_query(item)
                if value and value not in values:
                    values.append(value)

        alias_rules: dict[str, dict[str, list[str]]] = {
            "telegram": {
                "search": ["telegram", "unigram"],
                "focus_titles": ["telegram", "unigram"],
                "focus_processes": ["telegram.exe", "unigram.exe"],
                "uri": ["telegram:"],
                "shell": [],
            },
            "whatsapp": {
                "search": ["whatsapp"],
                "focus_titles": ["whatsapp"],
                "focus_processes": ["whatsapp.exe"],
                "uri": ["whatsapp:"],
                "shell": [],
            },
            "settings": {
                "search": ["settings"],
                "focus_titles": ["settings"],
                "focus_processes": ["systemsettings.exe"],
                "uri": ["ms-settings:"],
                "shell": ["ms-settings:"],
            },
            "bluetooth": {
                "search": ["bluetooth"],
                "focus_titles": ["settings", "bluetooth"],
                "focus_processes": ["systemsettings.exe"],
                "uri": ["ms-settings:bluetooth"],
                "shell": ["ms-settings:bluetooth"],
            },
            "calculator": {
                "search": ["calculator", "calc"],
                "focus_titles": ["calculator", "calc"],
                "focus_processes": ["calculatorapp.exe", "calculator.exe"],
                "uri": ["calculator:"],
                "shell": ["calc.exe"],
            },
            "calc": {
                "search": ["calculator", "calc"],
                "focus_titles": ["calculator", "calc"],
                "focus_processes": ["calculatorapp.exe", "calculator.exe"],
                "uri": ["calculator:"],
                "shell": ["calc.exe"],
            },
            "notepad": {
                "search": ["notepad"],
                "focus_titles": ["notepad"],
                "focus_processes": ["notepad.exe"],
                "uri": [],
                "shell": ["notepad.exe"],
            },
            "paint": {
                "search": ["paint", "mspaint"],
                "focus_titles": ["paint"],
                "focus_processes": ["mspaint.exe"],
                "uri": [],
                "shell": ["mspaint.exe"],
            },
            "cmd": {
                "search": ["command prompt", "cmd"],
                "focus_titles": ["command prompt", "cmd"],
                "focus_processes": ["cmd.exe"],
                "uri": [],
                "shell": ["cmd.exe"],
            },
            "command prompt": {
                "search": ["command prompt", "cmd"],
                "focus_titles": ["command prompt", "cmd"],
                "focus_processes": ["cmd.exe"],
                "uri": [],
                "shell": ["cmd.exe"],
            },
            "powershell": {
                "search": ["powershell", "windows powershell"],
                "focus_titles": ["powershell"],
                "focus_processes": ["powershell.exe", "pwsh.exe"],
                "uri": [],
                "shell": ["powershell.exe"],
            },
            "task manager": {
                "search": ["task manager", "taskmgr"],
                "focus_titles": ["task manager"],
                "focus_processes": ["taskmgr.exe"],
                "uri": [],
                "shell": ["taskmgr.exe"],
            },
            "taskmgr": {
                "search": ["task manager", "taskmgr"],
                "focus_titles": ["task manager"],
                "focus_processes": ["taskmgr.exe"],
                "uri": [],
                "shell": ["taskmgr.exe"],
            },
            "file explorer": {
                "search": ["file explorer", "explorer"],
                "focus_titles": ["file explorer", "explorer"],
                "focus_processes": ["explorer.exe"],
                "uri": [],
                "shell": ["explorer.exe"],
            },
            "explorer": {
                "search": ["file explorer", "explorer"],
                "focus_titles": ["file explorer", "explorer"],
                "focus_processes": ["explorer.exe"],
                "uri": [],
                "shell": ["explorer.exe"],
            },
        }

        search_terms: list[str] = [query_norm]
        focus_titles: list[str] = [query_norm]
        focus_processes: list[str] = []
        uri_targets: list[str] = []
        shell_targets: list[str] = []

        for token, bundle in alias_rules.items():
            if token in query_cf:
                _add_unique(search_terms, bundle.get("search", []))
                _add_unique(focus_titles, bundle.get("focus_titles", []))
                _add_unique(focus_processes, bundle.get("focus_processes", []))
                _add_unique(uri_targets, bundle.get("uri", []))
                _add_unique(shell_targets, bundle.get("shell", []))

        # If process-style query is provided directly.
        if query_cf.endswith(".exe"):
            _add_unique(focus_processes, [query_norm])
            _add_unique(shell_targets, [query_norm])

        # Focus already-open window first to avoid duplicate launches.
        for title in focus_titles[:4]:
            focused = self._try_focus_app(title_query=title, timeout_sec=1.3)
            if focused:
                return _json(
                    {
                        "launched": True,
                        "method": "focus_existing",
                        "query": query_norm,
                        "title": focused.get("title", ""),
                        "process_name": focused.get("process_name", ""),
                        "pid": focused.get("pid", 0),
                    }
                )
        for proc in focus_processes[:4]:
            focused = self._try_focus_app(process_query=proc, timeout_sec=1.2)
            if focused:
                return _json(
                    {
                        "launched": True,
                        "method": "focus_existing_process",
                        "query": query_norm,
                        "title": focused.get("title", ""),
                        "process_name": focused.get("process_name", ""),
                        "pid": focused.get("pid", 0),
                    }
                )

        # Search Start apps with ranking.
        search_terms_cf = [s.casefold() for s in search_terms if s]
        query_tokens = [tok for tok in re.split(r"\s+", query_cf) if tok]
        ranked_apps: list[dict[str, Any]] = []
        seen_app_ids: set[str] = set()
        for term in search_terms[:6]:
            raw_apps = self._search_start_apps(query=term, max_results=12)
            if not isinstance(raw_apps, str) or raw_apps.lower().startswith("error:"):
                continue
            payload = self._parse_json_dict(raw_apps)
            apps = payload.get("apps", []) if isinstance(payload, dict) else []
            for item in apps:
                if not isinstance(item, dict):
                    continue
                name = str(item.get("Name", "") or "").strip()
                app_id = str(item.get("AppID", "") or "").strip()
                if not app_id or app_id in seen_app_ids:
                    continue
                seen_app_ids.add(app_id)
                name_cf = name.casefold()
                score = 0
                if name_cf == query_cf:
                    score += 240
                if query_cf and query_cf in name_cf:
                    score += 140
                for token in search_terms_cf:
                    if token and token in name_cf:
                        score += 75
                for token in query_tokens:
                    if token in name_cf:
                        score += 20
                ranked_apps.append(
                    {
                        "score": score,
                        "Name": name,
                        "AppID": app_id,
                    }
                )
        ranked_apps.sort(key=lambda row: int(row.get("score", 0)), reverse=True)

        for candidate in ranked_apps[:8]:
            app_id = str(candidate.get("AppID", "") or "").strip()
            name = str(candidate.get("Name", "") or "").strip() or query_norm
            if not app_id:
                continue
            try:
                subprocess.Popen(
                    ["cmd", "/c", "start", "", f"shell:AppsFolder\\{app_id}"],
                    shell=False,
                )
                focused = (
                    self._try_focus_app(title_query=name, timeout_sec=3.8)
                    or self._try_focus_app(title_query=query_norm, timeout_sec=2.2)
                )
                if not focused:
                    for proc in focus_processes[:3]:
                        focused = self._try_focus_app(process_query=proc, timeout_sec=1.6)
                        if focused:
                            break
                payload: dict[str, Any] = {
                    "launched": True,
                    "method": "start_apps",
                    "Name": name,
                    "AppID": app_id,
                    "focused": bool(focused),
                }
                if focused:
                    payload.update(
                        {
                            "title": focused.get("title", ""),
                            "process_name": focused.get("process_name", ""),
                            "pid": focused.get("pid", 0),
                        }
                    )
                return _json(payload)
            except Exception:
                continue

        for uri in uri_targets:
            try:
                subprocess.Popen(["cmd", "/c", "start", "", uri], shell=False)
                focused = (
                    self._try_focus_app(title_query=query_norm, timeout_sec=2.4)
                    or self._try_focus_app(title_query="settings", timeout_sec=2.0)
                )
                return _json(
                    {
                        "launched": True,
                        "method": "uri",
                        "target": uri,
                        "focused": bool(focused),
                    }
                )
            except Exception:
                continue

        for shell_target in shell_targets:
            try:
                subprocess.Popen(["cmd", "/c", "start", "", shell_target], shell=False)
                focused = (
                    self._try_focus_app(title_query=query_norm, timeout_sec=2.8)
                    or self._try_focus_app(process_query=shell_target, timeout_sec=1.8)
                )
                return _json(
                    {
                        "launched": True,
                        "method": "shell_alias",
                        "target": shell_target,
                        "focused": bool(focused),
                    }
                )
            except Exception:
                continue

        try:
            subprocess.Popen(["cmd", "/c", "start", "", query_norm], shell=False)
            focused = self._try_focus_app(title_query=query_norm, timeout_sec=2.5)
            return _json(
                {
                    "launched": True,
                    "method": "shell_start",
                    "target": query_norm,
                    "focused": bool(focused),
                }
            )
        except Exception as e:
            return self._error(f"launch_start_app failed: {e}")

    def _search_files(self, query: str, max_results: int) -> str:
        query_norm = _normalize_query(query)
        if not query_norm:
            return self._error("query is required")
        max_results = _clamp(max_results, 1, 200)
        query_cf = query_norm.casefold()
        results: list[str] = []
        roots = _iter_search_roots()
        skip_dirs = {
            "node_modules",
            ".git",
            "__pycache__",
            ".venv",
            "venv",
            "appdata",
            "programdata",
            "$recycle.bin",
            "windows",
        }

        for root in roots:
            for current_root, dirs, files in os.walk(root, topdown=True):
                dirs[:] = [d for d in dirs if d.casefold() not in skip_dirs]
                for file_name in files:
                    if query_cf not in file_name.casefold():
                        continue
                    full = str(Path(current_root) / file_name)
                    results.append(full)
                    if len(results) >= max_results:
                        return _json({"count": len(results), "results": results})
        return _json({"count": len(results), "results": results})

    def _move_mouse_to_desktop_file(self, query: str, duration: float, timeout_sec: float) -> str:
        blocked = self._guard_interactive_action("move_mouse_to_desktop_file")
        if blocked:
            return blocked

        query_norm = _normalize_query(query)
        if not query_norm:
            return self._error("query is required")

        desktop_dir = Path.home() / "Desktop"
        if not desktop_dir.is_dir():
            return self._error("Desktop folder not found")

        query_cf = query_norm.casefold()
        candidates: list[Path] = []
        try:
            for item in desktop_dir.iterdir():
                name_cf = item.name.casefold()
                stem_cf = item.stem.casefold()
                if query_cf == name_cf or query_cf == stem_cf:
                    candidates.append(item)
                    continue
                if query_cf in name_cf or query_cf in stem_cf:
                    candidates.append(item)
        except Exception as e:
            return self._error(f"failed to inspect desktop items: {e}")

        if not candidates:
            return self._error(f"No desktop file matched: {query_norm}")

        candidates.sort(key=lambda p: (len(p.name), p.name.casefold()))
        target = candidates[0]
        label = target.stem or target.name

        attempts = [
            {"window_title": "Desktop", "control_type": "ListItem"},
            {"window_title": "Desktop", "control_type": ""},
            {"window_title": "", "control_type": "ListItem"},
        ]
        last_err = ""
        for attempt in attempts:
            raw = self._ui_target(
                window_title=attempt["window_title"],
                process_name="",
                control_name=label,
                auto_id="",
                control_type=attempt["control_type"],
                index=0,
                interaction="move",
                duration=duration,
                timeout_sec=timeout_sec,
            )
            if str(raw).lower().startswith("error:"):
                last_err = str(raw).replace("Error: ", "", 1)
                continue
            try:
                data = json.loads(raw)
            except Exception:
                last_err = "invalid response from ui_target"
                continue
            if isinstance(data, dict) and data.get("ok"):
                data["file_name"] = target.name
                data["file_path"] = str(target)
                return _json(data)
            last_err = "ui_target did not return success"

        try:
            subprocess.Popen(["explorer.exe", str(desktop_dir)], shell=False)
            time.sleep(0.8)
            raw = self._ui_target(
                window_title="Desktop",
                process_name="",
                control_name=label,
                auto_id="",
                control_type="ListItem",
                index=0,
                interaction="move",
                duration=duration,
                timeout_sec=max(4.0, timeout_sec),
            )
            if str(raw).lower().startswith("error:"):
                last_err = str(raw).replace("Error: ", "", 1)
            else:
                data = json.loads(raw)
                if isinstance(data, dict) and data.get("ok"):
                    data["file_name"] = target.name
                    data["file_path"] = str(target)
                    return _json(data)
        except Exception as e:
            last_err = f"fallback explorer path failed: {e}"

        return self._error(last_err or "Could not locate desktop icon on screen")

    def _list_installed_apps(self, query: str, max_results: int) -> str:
        query_norm = _normalize_query(query)
        max_results = _clamp(max_results, 1, 400)
        entries: dict[str, dict[str, Any]] = {}

        def add_entry(name: str, source: str) -> None:
            cleaned = name.strip()
            if not cleaned:
                return
            key = cleaned.casefold()
            if query_norm and query_norm.casefold() not in key:
                return
            if key not in entries:
                entries[key] = {"DisplayName": cleaned, "Source": source}

        raw_apps = self._search_start_apps(query="", max_results=600)
        if not raw_apps.lower().startswith("error:"):
            try:
                payload = json.loads(raw_apps)
                for item in (payload.get("apps", []) if isinstance(payload, dict) else []):
                    add_entry(str((item or {}).get("Name", "") or ""), "StartApps")
            except Exception:
                pass

        try:
            import winreg

            roots = [
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            ]
            for hive, key_path in roots:
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        index = 0
                        while True:
                            try:
                                sub_name = winreg.EnumKey(key, index)
                                index += 1
                            except OSError:
                                break
                            try:
                                with winreg.OpenKey(key, sub_name) as sub:
                                    name, _ = winreg.QueryValueEx(sub, "DisplayName")
                                    add_entry(str(name), "Registry")
                            except Exception:
                                continue
                except Exception:
                    continue
        except Exception:
            pass

        apps = sorted(entries.values(), key=lambda item: str(item.get("DisplayName", "")).casefold())
        return _json({"count": min(len(apps), max_results), "apps": apps[:max_results]})

    def _volume_state(self) -> dict[str, Any]:
        try:
            from pycaw.pycaw import AudioUtilities

            speakers = AudioUtilities.GetSpeakers()
            endpoint = getattr(speakers, "EndpointVolume", None)
            if endpoint is None:
                return self._legacy_volume_state()
            level = float(endpoint.GetMasterVolumeLevelScalar())
            muted = bool(endpoint.GetMute())
            return {"endpoint": endpoint, "level_percent": int(round(level * 100.0)), "muted": muted}
        except Exception as e:
            raise RuntimeError(f"volume backend unavailable: {e}") from e

    def _legacy_volume_state(self) -> dict[str, Any]:
        from ctypes import POINTER, cast

        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

        speakers = AudioUtilities.GetSpeakers()
        interface = speakers.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        endpoint = cast(interface, POINTER(IAudioEndpointVolume))
        level = float(endpoint.GetMasterVolumeLevelScalar())
        muted = bool(endpoint.GetMute())
        return {"endpoint": endpoint, "level_percent": int(round(level * 100.0)), "muted": muted}

    def _volume_control(self, mode: str, level: Any = None, delta: Any = None) -> str:
        mode_norm = (mode or "get").strip().lower()
        try:
            state = self._volume_state()
            endpoint = state["endpoint"]
            current = int(state["level_percent"])
            muted = bool(state["muted"])

            if mode_norm == "get":
                return _json({"level_percent": current, "muted": muted})
            if mode_norm in {"mute", "silence"}:
                endpoint.SetMute(1, None)
            elif mode_norm in {"unmute"}:
                endpoint.SetMute(0, None)
            elif mode_norm in {"max", "maximum"}:
                endpoint.SetMasterVolumeLevelScalar(1.0, None)
                endpoint.SetMute(0, None)
            elif mode_norm in {"min", "minimum"}:
                endpoint.SetMasterVolumeLevelScalar(0.0, None)
                endpoint.SetMute(0, None)
            elif mode_norm == "set":
                if level is None:
                    return self._error("level is required for mode=set")
                new_level = _clamp(int(level), 0, 100)
                endpoint.SetMasterVolumeLevelScalar(float(new_level) / 100.0, None)
                endpoint.SetMute(0 if new_level > 0 else int(muted), None)
            elif mode_norm == "up":
                d = _clamp(int(delta or 8), 1, 100)
                new_level = _clamp(current + d, 0, 100)
                endpoint.SetMasterVolumeLevelScalar(float(new_level) / 100.0, None)
                endpoint.SetMute(0, None)
            elif mode_norm == "down":
                d = _clamp(int(delta or 8), 1, 100)
                new_level = _clamp(current - d, 0, 100)
                endpoint.SetMasterVolumeLevelScalar(float(new_level) / 100.0, None)
            else:
                return self._error(f"unsupported volume mode: {mode_norm}")

            refreshed = self._volume_state()
            return _json(
                {
                    "level_percent": int(refreshed["level_percent"]),
                    "muted": bool(refreshed["muted"]),
                }
            )
        except Exception as e:
            return self._error(str(e))

    def _brightness_control(self, mode: str, level: Any = None, delta: Any = None) -> str:
        mode_norm = (mode or "get").strip().lower()
        try:
            _install_display_warning_filter_once()
            import screen_brightness_control as sbc

            def get_level() -> int:
                values = sbc.get_brightness()
                if isinstance(values, list):
                    if not values:
                        raise RuntimeError("No displays detected")
                    return int(round(sum(float(v) for v in values) / len(values)))
                return int(values)

            if mode_norm == "get":
                return _json({"brightness_percent": _clamp(get_level(), 0, 100)})

            current = _clamp(get_level(), 0, 100)
            if mode_norm in {"max", "maximum"}:
                target = 100
            elif mode_norm in {"min", "minimum"}:
                target = 0
            elif mode_norm == "set":
                if level is None:
                    return self._error("level is required for brightness set mode")
                target = _clamp(int(level), 0, 100)
            elif mode_norm == "up":
                target = _clamp(current + _clamp(int(delta or 10), 1, 100), 0, 100)
            elif mode_norm == "down":
                target = _clamp(current - _clamp(int(delta or 10), 1, 100), 0, 100)
            else:
                return self._error(f"unsupported brightness mode: {mode_norm}")

            sbc.set_brightness(target)
            return _json({"brightness_percent": _clamp(get_level(), 0, 100)})
        except Exception as e:
            return self._error(str(e))

    def _media_control(self, mode: str) -> str:
        mode_norm = (mode or "play_pause").strip().lower()
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            if mode_norm in {"play", "pause", "play_pause", "toggle"}:
                pyautogui.press("playpause")
                return "Media play/pause toggled."
            if mode_norm in {"next", "next_track"}:
                pyautogui.press("nexttrack")
                return "Skipped to next track."
            if mode_norm in {"previous", "prev", "previous_track"}:
                pyautogui.press("prevtrack")
                return "Returned to previous track."
            if mode_norm in {"stop"}:
                pyautogui.press("stop")
                return "Media playback stopped."
            return self._error(f"unsupported media mode: {mode_norm}")
        except Exception as e:
            return self._error(f"media_control failed: {e}")

    def _is_blocked_process(self, process_name: str) -> bool:
        name = str(process_name or "").strip().lower()
        if not name:
            return False
        return name in self.BLOCKED_AUTOMATION_PROCESSES

    def _is_blocked_window_class(self, class_name: str) -> bool:
        normalized = str(class_name or "").strip().casefold()
        if not normalized:
            return False
        return normalized in self.BLOCKED_WINDOW_CLASSES

    def _get_wrapper_class_name(self, wrapper: Any) -> str:
        try:
            info = wrapper.element_info
            return str(getattr(info, "class_name", "") or "").strip()
        except Exception:
            return ""

    def _foreground_process(self) -> tuple[str, int]:
        try:
            import psutil
            import win32gui
            import win32process

            hwnd = int(win32gui.GetForegroundWindow() or 0)
            if hwnd <= 0:
                return "", 0
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if not pid:
                return "", 0
            name = psutil.Process(pid).name()
            return str(name or ""), int(pid)
        except Exception:
            return "", 0

    def _guard_interactive_action(self, action: str, keys: list[str] | None = None) -> str | None:
        name, pid = self._foreground_process()
        if self._is_blocked_process(name):
            return self._error(f"interactive action blocked on protected process: {name} (PID {pid})")

        if action == "hotkey" and keys:
            keyset = frozenset(str(k).strip().lower() for k in keys if str(k).strip())
            if keyset in self.BLOCKED_HOTKEY_COMBINATIONS:
                combo = " + ".join(sorted(keyset))
                return self._error(f"blocked hotkey combination: {combo}")
        return None

    def _resolve_window(
        self,
        *,
        window_title: str = "",
        process_name: str = "",
        timeout_sec: float = 4.0,
    ) -> dict[str, Any]:
        from pywinauto import Desktop

        title_query = _normalize_query(window_title).casefold()
        process_query = _normalize_query(process_name).casefold()
        timeout_sec = max(0.3, min(20.0, float(timeout_sec)))
        deadline = time.time() + timeout_sec

        last_err = "No matching window found."
        while time.time() < deadline:
            try:
                import psutil
                import win32gui
                import win32process

                desktop = Desktop(backend="uia")

                if not title_query and not process_query:
                    hwnd = int(win32gui.GetForegroundWindow() or 0)
                    if hwnd > 0:
                        try:
                            _, pid = win32process.GetWindowThreadProcessId(hwnd)
                            p_name = psutil.Process(pid).name() if pid else ""
                            if self._is_blocked_process(p_name):
                                raise RuntimeError(f"target process is protected: {p_name}")
                            wrapper = desktop.window(handle=hwnd)
                            class_name = self._get_wrapper_class_name(wrapper)
                            if self._is_blocked_window_class(class_name):
                                raise RuntimeError(
                                    f"target window class is blocked: {class_name or 'unknown'}"
                                )
                            return {
                                "wrapper": wrapper,
                                "title": str(wrapper.window_text() or ""),
                                "process_name": str(p_name or ""),
                                "pid": int(pid or 0),
                                "hwnd": hwnd,
                                "class_name": class_name,
                            }
                        except Exception:
                            pass

                best: dict[str, Any] | None = None
                best_score = -1
                for wrapper in desktop.windows():
                    try:
                        title = str(wrapper.window_text() or "").strip()
                        if not title:
                            continue
                        info = wrapper.element_info
                        pid = int(getattr(info, "process_id", 0) or 0)
                        class_name = str(getattr(info, "class_name", "") or "").strip()
                        if self._is_blocked_window_class(class_name):
                            continue
                        p_name = ""
                        if pid:
                            try:
                                p_name = str(psutil.Process(pid).name() or "")
                            except Exception:
                                p_name = ""
                        if self._is_blocked_process(p_name):
                            continue

                        title_cf = title.casefold()
                        process_cf = p_name.casefold()
                        if title_query and title_query not in title_cf:
                            continue
                        if process_query and process_query not in process_cf:
                            continue

                        score = 0
                        if title_query:
                            if title_cf == title_query:
                                score += 120
                            else:
                                score += 80
                        if process_query:
                            if process_cf == process_query:
                                score += 80
                            else:
                                score += 50
                        if not title_query and not process_query:
                            score += 10
                        if score > best_score:
                            best_score = score
                            hwnd = int(getattr(info, "handle", 0) or 0)
                            best = {
                                "wrapper": wrapper,
                                "title": title,
                                "process_name": p_name,
                                "pid": pid,
                                "hwnd": hwnd,
                                "class_name": class_name,
                            }
                    except Exception:
                        continue

                if best is not None:
                    return best
            except Exception as e:
                last_err = str(e)

            time.sleep(0.15)

        raise RuntimeError(last_err)

    def _enumerate_window_controls(self, window_wrapper: Any, max_items: int = 400) -> list[dict[str, Any]]:
        controls: list[dict[str, Any]] = []
        seen: set[tuple[str, str, str]] = set()

        candidates: list[Any] = [window_wrapper]
        try:
            descendants = window_wrapper.descendants()
            candidates.extend(descendants[: max(1, max_items * 2)])
        except Exception:
            pass

        for wrapper in candidates:
            try:
                info = wrapper.element_info
                name = str(wrapper.window_text() or "").strip()
                auto_id = str(getattr(info, "automation_id", "") or "").strip()
                control_type = str(getattr(info, "control_type", "") or "").strip()
                class_name = str(getattr(info, "class_name", "") or "").strip()
                left = top = right = bottom = width = height = 0
                try:
                    rect = wrapper.rectangle()
                    left = int(getattr(rect, "left", 0) or 0)
                    top = int(getattr(rect, "top", 0) or 0)
                    right = int(getattr(rect, "right", 0) or 0)
                    bottom = int(getattr(rect, "bottom", 0) or 0)
                    width = max(0, right - left)
                    height = max(0, bottom - top)
                except Exception:
                    pass
                if not any((name, auto_id, control_type, class_name)):
                    continue
                key = (name.casefold(), auto_id.casefold(), control_type.casefold())
                if key in seen:
                    continue
                seen.add(key)
                controls.append(
                    {
                        "wrapper": wrapper,
                        "name": name,
                        "auto_id": auto_id,
                        "control_type": control_type,
                        "class_name": class_name,
                        "left": left,
                        "top": top,
                        "right": right,
                        "bottom": bottom,
                        "width": width,
                        "height": height,
                    }
                )
                if len(controls) >= max_items:
                    break
            except Exception:
                continue
        return controls

    def _pick_control(
        self,
        controls: list[dict[str, Any]],
        *,
        control_name: str = "",
        auto_id: str = "",
        control_type: str = "",
        index: int = 0,
    ) -> dict[str, Any]:
        if not controls:
            raise RuntimeError("No controls found in target window.")

        name_q = _normalize_query(control_name).casefold()
        auto_q = _normalize_query(auto_id).casefold()
        type_q = _normalize_query(control_type).casefold()
        name_tokens = [t for t in re.split(r"\s+", name_q) if t]

        ranked: list[tuple[int, dict[str, Any]]] = []
        for item in controls:
            name = str(item.get("name", "") or "")
            aid = str(item.get("auto_id", "") or "")
            ctype = str(item.get("control_type", "") or "")
            klass = str(item.get("class_name", "") or "")
            left = int(item.get("left", 0) or 0)
            top = int(item.get("top", 0) or 0)
            width = int(item.get("width", 0) or 0)
            height = int(item.get("height", 0) or 0)

            if width <= 1 or height <= 1:
                continue
            if abs(left) > 10000 or abs(top) > 10000:
                continue

            name_cf = name.casefold()
            aid_cf = aid.casefold()
            ctype_cf = ctype.casefold()
            klass_cf = klass.casefold()

            token_ratio = 0.0
            if name_q:
                if "\n" in name or len(name) > 120:
                    continue
                if not type_q and ctype_cf in {"text", "document"}:
                    continue
                exact_match = name_cf == name_q
                contains_match = bool(name_cf) and (name_q in name_cf or name_cf in name_q)
                token_hits = sum(1 for t in name_tokens if t in name_cf)
                token_ratio = (token_hits / len(name_tokens)) if name_tokens else 0.0
                if not (exact_match or contains_match or token_ratio >= 0.5):
                    continue
            if auto_q and auto_q not in aid_cf:
                continue
            if type_q and type_q not in ctype_cf and type_q not in klass_cf:
                continue

            score = 0
            if name_q:
                if name_cf == name_q:
                    score += 120
                elif name_q in name_cf or name_cf in name_q:
                    score += 90
                else:
                    score += int(60 * token_ratio)
            if auto_q:
                score += 90 if aid_cf == auto_q else 55
            if type_q:
                score += 40
            if name and not name_q:
                score += 5
            if "\n" in name:
                score -= 40
            if len(name) > 180:
                score -= 30
            clickable_types = {
                "button",
                "hyperlink",
                "menuitem",
                "tabitem",
                "listitem",
                "treeitem",
                "checkbox",
                "radiobutton",
                "splitbutton",
            }
            if ctype_cf in clickable_types:
                score += 25
            if ctype_cf in {"text", "document"}:
                score -= 15

            ranked.append((score, item))

        if not ranked:
            raise RuntimeError("No matching control found.")

        ranked.sort(key=lambda pair: pair[0], reverse=True)
        safe_index = _clamp(index, 0, len(ranked) - 1)
        return ranked[safe_index][1]

    def _control_center(self, wrapper: Any) -> tuple[int, int]:
        try:
            rect = wrapper.rectangle()
            left = int(getattr(rect, "left", 0) or 0)
            top = int(getattr(rect, "top", 0) or 0)
            right = int(getattr(rect, "right", 0) or 0)
            bottom = int(getattr(rect, "bottom", 0) or 0)
            if right <= left or bottom <= top:
                raise RuntimeError("control has invalid bounds")
            x = left + ((right - left) // 2)
            y = top + ((bottom - top) // 2)
            return x, y
        except Exception as e:
            raise RuntimeError(f"failed to resolve control bounds: {e}") from e

    def _focus_window(self, window_title: str, process_name: str, timeout_sec: float) -> str:
        try:
            target = self._resolve_window(
                window_title=window_title,
                process_name=process_name,
                timeout_sec=timeout_sec,
            )
            wrapper = target["wrapper"]
            try:
                if hasattr(wrapper, "is_minimized") and wrapper.is_minimized():
                    wrapper.restore()
            except Exception:
                pass
            try:
                wrapper.set_focus()
            except Exception:
                try:
                    wrapper.click_input()
                except Exception:
                    pass
            return _json(
                {
                    "focused": True,
                    "title": target.get("title", ""),
                    "process_name": target.get("process_name", ""),
                    "pid": target.get("pid", 0),
                }
            )
        except Exception as e:
            return self._error(f"focus_window failed: {e}")

    def _ui_list_controls(
        self,
        *,
        window_title: str,
        process_name: str,
        control_name: str,
        control_type: str,
        max_results: int,
        timeout_sec: float,
    ) -> str:
        try:
            target = self._resolve_window(
                window_title=window_title,
                process_name=process_name,
                timeout_sec=timeout_sec,
            )
            controls = self._enumerate_window_controls(target["wrapper"], max_items=max(50, max_results * 5))

            name_q = _normalize_query(control_name).casefold()
            type_q = _normalize_query(control_type).casefold()
            filtered: list[dict[str, Any]] = []
            for item in controls:
                name = str(item.get("name", "") or "")
                ctype = str(item.get("control_type", "") or "")
                klass = str(item.get("class_name", "") or "")
                if name_q and name_q not in name.casefold():
                    continue
                if type_q and type_q not in ctype.casefold() and type_q not in klass.casefold():
                    continue
                filtered.append(
                    {
                        "name": name,
                        "auto_id": str(item.get("auto_id", "") or ""),
                        "control_type": ctype,
                        "class_name": klass,
                    }
                )
                if len(filtered) >= _clamp(max_results, 1, 200):
                    break

            return _json(
                {
                    "window_title": target.get("title", ""),
                    "process_name": target.get("process_name", ""),
                    "count": len(filtered),
                    "controls": filtered,
                }
            )
        except Exception as e:
            return self._error(f"ui_list_controls failed: {e}")

    def _ui_click(
        self,
        *,
        window_title: str,
        process_name: str,
        control_name: str,
        auto_id: str,
        control_type: str,
        index: int,
        timeout_sec: float,
    ) -> str:
        blocked = self._guard_interactive_action("ui_click")
        if blocked:
            return blocked
        try:
            target = self._resolve_window(
                window_title=window_title,
                process_name=process_name,
                timeout_sec=timeout_sec,
            )
            controls = self._enumerate_window_controls(target["wrapper"])
            selected = self._pick_control(
                controls,
                control_name=control_name,
                auto_id=auto_id,
                control_type=control_type,
                index=index,
            )
            wrapper = selected["wrapper"]
            try:
                if hasattr(wrapper, "invoke"):
                    wrapper.invoke()
                else:
                    wrapper.click_input()
            except Exception:
                wrapper.click_input()
            return _json(
                {
                    "clicked": True,
                    "window_title": target.get("title", ""),
                    "process_name": target.get("process_name", ""),
                    "control_name": selected.get("name", ""),
                    "auto_id": selected.get("auto_id", ""),
                    "control_type": selected.get("control_type", ""),
                }
            )
        except Exception as e:
            return self._error(f"ui_click failed: {e}")

    def _ui_set_text(
        self,
        *,
        text: str,
        window_title: str,
        process_name: str,
        control_name: str,
        auto_id: str,
        control_type: str,
        index: int,
        press_enter: bool,
        timeout_sec: float,
    ) -> str:
        blocked = self._guard_interactive_action("ui_set_text")
        if blocked:
            return blocked
        text_value = str(text or "")
        if not text_value:
            return self._error("text is required")
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            target = self._resolve_window(
                window_title=window_title,
                process_name=process_name,
                timeout_sec=timeout_sec,
            )
            controls = self._enumerate_window_controls(target["wrapper"])
            selected = self._pick_control(
                controls,
                control_name=control_name,
                auto_id=auto_id,
                control_type=control_type,
                index=index,
            )
            wrapper = selected["wrapper"]
            used_method = "typewrite"
            try:
                if hasattr(wrapper, "set_edit_text"):
                    wrapper.set_edit_text(text_value)
                    used_method = "set_edit_text"
                else:
                    wrapper.click_input()
                    pyautogui.hotkey("ctrl", "a")
                    pyautogui.press("backspace")
                    pyautogui.write(text_value, interval=0.01)
            except Exception:
                wrapper.click_input()
                pyautogui.hotkey("ctrl", "a")
                pyautogui.press("backspace")
                pyautogui.write(text_value, interval=0.01)
            if press_enter:
                pyautogui.press("enter")
            return _json(
                {
                    "text_set": True,
                    "method": used_method,
                    "window_title": target.get("title", ""),
                    "process_name": target.get("process_name", ""),
                    "control_name": selected.get("name", ""),
                    "chars": len(text_value),
                    "pressed_enter": bool(press_enter),
                }
            )
        except Exception as e:
            return self._error(f"ui_set_text failed: {e}")

    def _ui_target(
        self,
        *,
        window_title: str,
        process_name: str,
        control_name: str,
        auto_id: str,
        control_type: str,
        index: int,
        interaction: str,
        duration: float,
        timeout_sec: float,
    ) -> str:
        blocked = self._guard_interactive_action("ui_target")
        if blocked:
            return blocked
        try:
            import pyautogui

            pyautogui.FAILSAFE = False

            interaction_norm = (interaction or "move").strip().lower()
            interaction_aliases = {
                "hover": "move",
                "point": "move",
                "double": "double_click",
                "double-click": "double_click",
                "right": "right_click",
                "context": "right_click",
            }
            interaction_norm = interaction_aliases.get(interaction_norm, interaction_norm)
            if interaction_norm not in {"move", "click", "double_click", "right_click"}:
                return self._error(f"unsupported ui_target interaction: {interaction_norm}")

            target = self._resolve_window(
                window_title=window_title,
                process_name=process_name,
                timeout_sec=timeout_sec,
            )
            controls = self._enumerate_window_controls(target["wrapper"])
            screen = pyautogui.size()
            selected: dict[str, Any] | None = None
            x = y = 0
            limit = min(len(controls), 20)
            start_index = max(0, int(index))
            coord_error = ""
            for idx in range(start_index, limit):
                candidate = self._pick_control(
                    controls,
                    control_name=control_name,
                    auto_id=auto_id,
                    control_type=control_type,
                    index=idx,
                )
                candidate_name = str(candidate.get("name", "") or "").strip()
                if control_name and not candidate_name and not str(candidate.get("auto_id", "") or "").strip():
                    continue
                cx, cy = self._control_center(candidate["wrapper"])
                if abs(int(cx)) > int(screen.width * 3) or abs(int(cy)) > int(screen.height * 3):
                    coord_error = f"target coordinates out of expected bounds: ({cx}, {cy})"
                    continue
                selected = candidate
                x, y = int(cx), int(cy)
                break

            if selected is None:
                if coord_error:
                    raise RuntimeError(coord_error)
                raise RuntimeError("No viable control match found.")

            move_duration = max(0.0, min(3.0, float(duration)))

            if interaction_norm == "move":
                pyautogui.moveTo(x, y, duration=move_duration)
            elif interaction_norm == "click":
                pyautogui.click(x=x, y=y, button="left", clicks=1)
            elif interaction_norm == "double_click":
                pyautogui.click(x=x, y=y, button="left", clicks=2)
            elif interaction_norm == "right_click":
                pyautogui.click(x=x, y=y, button="right", clicks=1)

            return _json(
                {
                    "ok": True,
                    "interaction": interaction_norm,
                    "x": x,
                    "y": y,
                    "window_title": target.get("title", ""),
                    "process_name": target.get("process_name", ""),
                    "control_name": selected.get("name", ""),
                    "auto_id": selected.get("auto_id", ""),
                    "control_type": selected.get("control_type", ""),
                }
            )
        except Exception as e:
            return self._error(f"ui_target failed: {e}")

    def _mouse_move(self, x: int, y: int, duration: float) -> str:
        blocked = self._guard_interactive_action("mouse_move")
        if blocked:
            return blocked
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            pyautogui.moveTo(int(x), int(y), duration=max(0.0, min(3.0, float(duration))))
            return _json({"ok": True, "x": int(x), "y": int(y)})
        except Exception as e:
            return self._error(f"mouse_move failed: {e}")

    def _click(self, x: Any, y: Any, button: str, clicks: int) -> str:
        blocked = self._guard_interactive_action("click")
        if blocked:
            return blocked
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            btn = (button or "left").strip().lower()
            if btn not in {"left", "right", "middle"}:
                btn = "left"
            c = _clamp(clicks, 1, 5)
            if x is not None and y is not None:
                pyautogui.click(x=int(x), y=int(y), clicks=c, button=btn)
                return _json({"ok": True, "x": int(x), "y": int(y), "button": btn, "clicks": c})
            pyautogui.click(clicks=c, button=btn)
            pos = pyautogui.position()
            return _json({"ok": True, "x": int(pos.x), "y": int(pos.y), "button": btn, "clicks": c})
        except Exception as e:
            return self._error(f"click failed: {e}")

    def _press_key(self, key: str) -> str:
        key_norm = (key or "").strip().lower()
        if not key_norm:
            return self._error("key is required")
        blocked = self._guard_interactive_action("press_key")
        if blocked:
            return blocked
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            pyautogui.press(key_norm)
            return _json({"ok": True, "key": key_norm})
        except Exception as e:
            return self._error(f"press_key failed: {e}")

    def _type_text(self, text: str, press_enter: bool, interval: float) -> str:
        text_value = str(text or "")
        if not text_value:
            return self._error("text is required")
        blocked = self._guard_interactive_action("type_text")
        if blocked:
            return blocked
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            interval_sec = max(0.0, min(0.4, float(interval)))
            pyautogui.write(text_value, interval=interval_sec)
            if press_enter:
                pyautogui.press("enter")
            return _json(
                {
                    "ok": True,
                    "typed": True,
                    "chars": len(text_value),
                    "pressed_enter": bool(press_enter),
                }
            )
        except Exception as e:
            return self._error(f"type_text failed: {e}")

    def _hotkey(self, keys: list[str]) -> str:
        usable = [k.strip().lower() for k in keys if str(k).strip()]
        if not usable:
            return self._error("keys are required")
        blocked = self._guard_interactive_action("hotkey", keys=usable)
        if blocked:
            return blocked
        try:
            import pyautogui

            pyautogui.FAILSAFE = False
            pyautogui.hotkey(*usable)
            return _json({"ok": True, "keys": usable})
        except Exception as e:
            return self._error(f"hotkey failed: {e}")

    def _camera_snapshot(self, camera_index: int) -> str:
        try:
            import cv2

            media_dir = get_media_dir()
            out_path = media_dir / f"camera_{_timestamp_id()}.jpg"
            cap = cv2.VideoCapture(camera_index, cv2.CAP_DSHOW)
            if not cap or not cap.isOpened():
                cap = cv2.VideoCapture(camera_index)
            if not cap or not cap.isOpened():
                return self._error("camera is not available")
            try:
                time.sleep(0.2)
                ok, frame = cap.read()
                if not ok or frame is None:
                    return self._error("failed to read camera frame")
                cv2.imwrite(str(out_path), frame)
            finally:
                cap.release()
            return f"Camera snapshot saved to {out_path}"
        except Exception as e:
            return self._error(f"camera_snapshot failed: {e}")

    def _microphone_record(self, seconds: float) -> str:
        sec = max(0.4, min(30.0, float(seconds or 3.0)))
        try:
            import numpy as np
            import sounddevice as sd

            samplerate = 16000
            total_samples = int(sec * samplerate)
            if total_samples <= 0:
                return self._error("recording duration must be positive")
            media_dir = get_media_dir()
            out_path = media_dir / f"mic_{_timestamp_id()}.wav"
            audio = sd.rec(total_samples, samplerate=samplerate, channels=1, dtype="int16")
            sd.wait()
            if audio is None:
                return self._error("no audio data captured")
            if not isinstance(audio, np.ndarray):
                return self._error("audio capture returned invalid data")
            with wave.open(str(out_path), "wb") as wf:
                wf.setnchannels(1)
                wf.setsampwidth(2)
                wf.setframerate(samplerate)
                wf.writeframes(audio.tobytes())
            return f"Microphone recording saved to {out_path}"
        except Exception as e:
            return self._error(f"microphone_record failed: {e}")

    def _open_settings_page(self, page: str) -> str:
        page_norm = _normalize_query(page).casefold()
        mapping = {
            "": "ms-settings:",
            "settings": "ms-settings:",
            "bluetooth": "ms-settings:bluetooth",
            "wifi": "ms-settings:network-wifi",
            "network": "ms-settings:network",
            "display": "ms-settings:display",
            "sound": "ms-settings:sound",
            "volume": "ms-settings:sound",
            "apps": "ms-settings:appsfeatures",
            "default apps": "ms-settings:defaultapps",
            "defaultapps": "ms-settings:defaultapps",
            "privacy": "ms-settings:privacy",
            "update": "ms-settings:windowsupdate",
            "language": "ms-settings:regionlanguage",
            "keyboard": "ms-settings:typing",
        }
        uri = mapping.get(page_norm, None)
        if uri is None:
            if "default" in page_norm and "app" in page_norm:
                uri = "ms-settings:defaultapps"
            elif "bluetooth" in page_norm:
                uri = "ms-settings:bluetooth"
            elif "app" in page_norm:
                uri = "ms-settings:appsfeatures"
            else:
                uri = "ms-settings:"
        try:
            subprocess.Popen(["cmd", "/c", "start", "", uri], shell=False)
            return _json({"opened": True, "uri": uri})
        except Exception as e:
            return self._error(f"open_settings_page failed: {e}")

    def _bluetooth_control(self, mode: str) -> str:
        mode_norm = (mode or "open_settings").strip().lower()
        if mode_norm in {"open", "open_settings", "settings"}:
            return self._open_settings_page("bluetooth")

        desired = None
        if mode_norm in {"on", "enable", "start"}:
            desired = "On"
        elif mode_norm in {"off", "disable", "stop"}:
            desired = "Off"
        elif mode_norm in {"toggle"}:
            desired = "Toggle"

        if desired is None:
            return self._error(f"unsupported bluetooth mode: {mode_norm}")

        set_script = f"""
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{
    $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1
}} | Select-Object -First 1)
if (-not $asTask) {{ throw 'AsTask bridge unavailable' }}
$null = [Windows.Devices.Radios.Radio, Windows.System.Devices, ContentType=WindowsRuntime]
$accessOp = [Windows.Devices.Radios.Radio]::RequestAccessAsync()
$accessTask = $asTask.MakeGenericMethod([Windows.Devices.Radios.RadioAccessStatus]).Invoke($null, @($accessOp))
$accessTask.Wait()
if ($accessTask.Result -ne [Windows.Devices.Radios.RadioAccessStatus]::Allowed) {{
    throw 'Bluetooth access denied'
}}
$radiosOp = [Windows.Devices.Radios.Radio]::GetRadiosAsync()
$radiosTask = $asTask.MakeGenericMethod([System.Collections.Generic.IReadOnlyList[Windows.Devices.Radios.Radio]]).Invoke($null, @($radiosOp))
$radiosTask.Wait()
$radio = $radiosTask.Result | Where-Object {{ $_.Kind -eq [Windows.Devices.Radios.RadioKind]::Bluetooth }} | Select-Object -First 1
if (-not $radio) {{ throw 'Bluetooth radio not found' }}
$current = $radio.State.ToString()
$target = '{desired}'
if ($target -eq 'Toggle') {{
    if ($current -eq 'On') {{ $target = 'Off' }} else {{ $target = 'On' }}
}}
$setOp = $radio.SetStateAsync([Windows.Devices.Radios.RadioState]::$target)
$setTask = $asTask.MakeGenericMethod([Windows.Devices.Radios.RadioAccessStatus]).Invoke($null, @($setOp))
$setTask.Wait()
$stateNow = $radio.State.ToString()
$obj = [ordered]@{{ ok = $true; requested = $target.ToLower(); state = $stateNow.ToLower() }}
$obj | ConvertTo-Json -Compress
"""
        ok, output = _run_powershell(set_script, timeout=25)
        if ok and output:
            try:
                data = json.loads(output)
                if isinstance(data, dict):
                    return _json(data)
            except Exception:
                return _json({"ok": True, "state": output.strip().casefold(), "requested": mode_norm})

        fallback = self._open_settings_page("bluetooth")
        if fallback.lower().startswith("error:"):
            return self._error(f"bluetooth control failed: {output}")
        return _json(
            {
                "ok": False,
                "requested": mode_norm,
                "error": output or "Bluetooth state change failed",
                "opened_settings": True,
            }
        )


class ScreenshotTool(BaseTool):
    @property
    def name(self) -> str:
        return "screenshot"

    @property
    def description(self) -> str:
        return "Capture one screenshot of the current desktop."

    async def execute(self, **params: Any) -> str:
        return await DesktopTool().execute(action="screen_snapshot")


class StatusTool(BaseTool):
    @property
    def name(self) -> str:
        return "desktop_status"

    @property
    def description(self) -> str:
        return "Return a lightweight desktop overview."

    async def execute(self, **params: Any) -> str:
        return await DesktopTool().execute(action="desktop_overview")
