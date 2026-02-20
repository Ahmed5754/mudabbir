"""Agent backend registry with lazy imports and alias normalization."""

from __future__ import annotations

import importlib
from dataclasses import dataclass
from typing import Any

from Mudabbir.agents.protocol import BackendInfo, Capability


@dataclass(frozen=True)
class _BackendSpec:
    info: BackendInfo
    module_path: str
    class_name: str


_BACKEND_SPECS: dict[str, _BackendSpec] = {
    "claude_agent_sdk": _BackendSpec(
        info=BackendInfo(
            name="claude_agent_sdk",
            display_name="Claude SDK",
            description=(
                "Official Claude Agent SDK backend with built-in tools and robust streaming."
            ),
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MCP
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("claude", "claude_sdk", "default", "claude_code"),
            install_hint="pip install claude-agent-sdk",
        ),
        module_path="Mudabbir.agents.claude_sdk",
        class_name="ClaudeAgentSDKWrapper",
    ),
    "Mudabbir_native": _BackendSpec(
        info=BackendInfo(
            name="Mudabbir_native",
            display_name="Mudabbir Native",
            description="Mudabbir's native orchestrator with first-party tool integration.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("mudabbir", "native", "mudabbir_native"),
            install_hint="pip install 'Mudabbir[native]'",
        ),
        module_path="Mudabbir.agents.mudabbir_native",
        class_name="MudabbirOrchestrator",
    ),
    "open_interpreter": _BackendSpec(
        info=BackendInfo(
            name="open_interpreter",
            display_name="Open Interpreter",
            description="Standalone Open Interpreter backend (experimental).",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("oi", "gemini_cli"),
            install_hint="pip install 'Mudabbir[native]'",
        ),
        module_path="Mudabbir.agents.open_interpreter",
        class_name="OpenInterpreterAgent",
    ),
}

_BACKEND_LOOKUP: dict[str, str] = {name.lower(): name for name in _BACKEND_SPECS}
for canonical, spec in _BACKEND_SPECS.items():
    for alias in spec.info.aliases:
        _BACKEND_LOOKUP[alias.lower()] = canonical


def list_backend_names() -> list[str]:
    """Return all registered backend canonical names."""
    return list(_BACKEND_SPECS.keys())


def list_backend_infos() -> list[BackendInfo]:
    """Return static metadata for all registered backends."""
    return [spec.info for spec in _BACKEND_SPECS.values()]


def normalize_backend_name(name: str | None, *, fallback: str = "claude_agent_sdk") -> str:
    """Normalize a backend name/alias to a canonical registered backend."""
    key = (name or "").strip().lower()
    if key and key in _BACKEND_LOOKUP:
        return _BACKEND_LOOKUP[key]
    return fallback


def get_backend_info(name: str | None) -> BackendInfo:
    """Get backend metadata for a name/alias."""
    canonical = normalize_backend_name(name)
    return _BACKEND_SPECS[canonical].info


def load_backend_class(name: str | None):
    """Lazily import and return the backend class for a name/alias."""
    canonical = normalize_backend_name(name)
    spec = _BACKEND_SPECS[canonical]
    module = importlib.import_module(spec.module_path)
    return getattr(module, spec.class_name)


def create_backend(name: str | None, settings: Any):
    """Instantiate a backend for a name/alias."""
    cls = load_backend_class(name)
    return cls(settings)

