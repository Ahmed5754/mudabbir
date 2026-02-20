"""Capability registry and lookup helpers."""

from __future__ import annotations

from collections import defaultdict

from Mudabbir.tools.capabilities.schema import CapabilitySpec
from Mudabbir.tools.capabilities.windows_catalog import WINDOWS_CAPABILITIES


class CapabilityRegistry:
    """In-memory catalog for planner/discovery."""

    def __init__(self, capabilities: tuple[CapabilitySpec, ...] = WINDOWS_CAPABILITIES):
        self._capabilities = capabilities
        self._by_id = {cap.id: cap for cap in capabilities}
        self._by_stage: dict[str, list[CapabilitySpec]] = defaultdict(list)
        for cap in capabilities:
            self._by_stage[cap.stage].append(cap)

    def all(self) -> tuple[CapabilitySpec, ...]:
        return self._capabilities

    def by_id(self, capability_id: str) -> CapabilitySpec | None:
        return self._by_id.get(capability_id)

    def by_stage(self, stage: str) -> list[CapabilitySpec]:
        return list(self._by_stage.get(stage, []))

    def allowed_actions_stage_a(self) -> set[str]:
        """Map Stage A capabilities to current DesktopTool action names."""
        return {
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


DEFAULT_CAPABILITY_REGISTRY = CapabilityRegistry()
