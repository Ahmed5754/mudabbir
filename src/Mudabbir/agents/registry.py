"""Agent backend registry with lazy imports, aliases, and availability checks."""

from __future__ import annotations

import importlib
import importlib.util
import logging
import shutil
from dataclasses import dataclass
from typing import Any

from Mudabbir.agents.protocol import BackendInfo, Capability

logger = logging.getLogger(__name__)


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
            builtin_tools=("bash", "read", "write", "edit", "glob", "grep", "web_search"),
            required_keys=("anthropic_api_key",),
            supported_providers=("anthropic", "ollama", "openai_compatible"),
            install_hint={
                "verify_import": "claude_agent_sdk",
                "external_cmd": "npm install -g @anthropic-ai/claude-code",
                "binary": "claude",
                "docs_url": "https://docs.anthropic.com/en/docs/claude-code",
            },
        ),
        module_path="Mudabbir.agents.claude_sdk",
        class_name="ClaudeAgentSDKWrapper",
    ),
    "Mudabbir_native": _BackendSpec(
        info=BackendInfo(
            name="Mudabbir_native",
            display_name="Mudabbir Native",
            description="Mudabbir native orchestrator with first-party tool integration.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("mudabbir", "native", "mudabbir_native"),
            builtin_tools=("shell", "read_file", "write_file", "list_dir", "memory"),
            required_keys=("anthropic_api_key",),
            supported_providers=("anthropic", "openai", "ollama", "gemini", "openai_compatible"),
            install_hint={
                "pip_package": "open-interpreter",
                "pip_spec": "Mudabbir[native]",
                "verify_import": "interpreter",
            },
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
            builtin_tools=("python", "shell", "desktop_automation"),
            supported_providers=("ollama", "gemini", "openai", "openai_compatible"),
            install_hint={
                "pip_package": "open-interpreter",
                "pip_spec": "Mudabbir[native]",
                "verify_import": "interpreter",
            },
            beta=True,
        ),
        module_path="Mudabbir.agents.open_interpreter",
        class_name="OpenInterpreterAgent",
    ),
    "openai_agents": _BackendSpec(
        info=BackendInfo(
            name="openai_agents",
            display_name="OpenAI Agents",
            description="OpenAI Agents SDK backend with runtime tool bridge support.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("openai_agent", "openai_agents_sdk"),
            builtin_tools=("function_tools",),
            required_keys=("openai_api_key",),
            supported_providers=("openai", "ollama", "openai_compatible"),
            install_hint={
                "pip_package": "openai-agents",
                "pip_spec": "Mudabbir[openai-agents]",
                "verify_import": "agents",
            },
            beta=True,
        ),
        module_path="Mudabbir.agents.openai_agents",
        class_name="OpenAIAgentsBackend",
    ),
    "google_adk": _BackendSpec(
        info=BackendInfo(
            name="google_adk",
            display_name="Google ADK",
            description="Google Agent Development Kit backend for Gemini and MCP tools.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MCP
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("adk", "gemini"),
            builtin_tools=("google_search", "code_execution"),
            required_keys=("google_api_key",),
            supported_providers=("gemini",),
            install_hint={
                "pip_package": "google-adk",
                "pip_spec": "Mudabbir[google-adk]",
                "verify_import": "google.adk",
            },
            beta=True,
        ),
        module_path="Mudabbir.agents.google_adk",
        class_name="GoogleADKBackend",
    ),
    "codex_cli": _BackendSpec(
        info=BackendInfo(
            name="codex_cli",
            display_name="Codex CLI",
            description="OpenAI Codex CLI subprocess backend.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("codex",),
            builtin_tools=("shell", "file_edit", "web_search"),
            required_keys=("openai_api_key",),
            supported_providers=("openai",),
            install_hint={
                "external_cmd": "npm install -g @openai/codex",
                "binary": "codex",
                "docs_url": "https://github.com/openai/codex",
            },
            beta=True,
        ),
        module_path="Mudabbir.agents.codex_cli",
        class_name="CodexCLIBackend",
    ),
    "opencode": _BackendSpec(
        info=BackendInfo(
            name="opencode",
            display_name="OpenCode",
            description="OpenCode server backend via REST API.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("open_code",),
            builtin_tools=(),
            supported_providers=(),
            install_hint={
                "external_cmd": "go install github.com/opencode-ai/opencode@latest",
                "binary": "opencode",
                "docs_url": "https://github.com/opencode-ai/opencode",
            },
            beta=True,
        ),
        module_path="Mudabbir.agents.opencode",
        class_name="OpenCodeBackend",
    ),
    "copilot_sdk": _BackendSpec(
        info=BackendInfo(
            name="copilot_sdk",
            display_name="Copilot SDK",
            description="GitHub Copilot SDK backend with optional BYOK provider routing.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            aliases=("copilot",),
            builtin_tools=("shell", "file_ops", "git", "web_search"),
            supported_providers=("copilot", "openai", "azure", "anthropic"),
            install_hint={
                "pip_package": "github-copilot-sdk",
                "pip_spec": "Mudabbir[copilot-sdk]",
                "verify_import": "copilot",
                "verify_attr": "CopilotClient",
                "external_cmd": "Install Copilot CLI from https://github.com/github/copilot-sdk",
                "binary": "copilot",
                "docs_url": "https://github.com/github/copilot-sdk",
            },
            beta=True,
        ),
        module_path="Mudabbir.agents.copilot_sdk",
        class_name="CopilotSDKBackend",
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


def _is_hint_available(info: BackendInfo) -> bool:
    """Check dependency hints (import + binary) for availability."""
    hint = info.install_hint or {}
    verify_import = hint.get("verify_import")
    if verify_import:
        try:
            if importlib.util.find_spec(verify_import) is None:
                return False
            verify_attr = hint.get("verify_attr")
            if verify_attr:
                mod = importlib.import_module(verify_import)
                if not hasattr(mod, verify_attr):
                    return False
        except Exception:
            return False

    binary = hint.get("binary")
    if binary and shutil.which(binary) is None:
        return False

    return True


def is_backend_available(name: str | None) -> bool:
    """Return True when a backend is importable and runtime hints are satisfied."""
    canonical = normalize_backend_name(name)
    try:
        load_backend_class(canonical)
    except Exception:
        return False
    return _is_hint_available(_BACKEND_SPECS[canonical].info)


def backend_summary(name: str | None) -> dict[str, Any]:
    """Return frontend-safe backend metadata with availability."""
    info = get_backend_info(name)
    capabilities = []
    for cap in Capability:
        if cap is Capability.NONE:
            continue
        if cap in info.capabilities:
            capabilities.append(cap.name.lower())

    return {
        "name": info.name,
        "displayName": info.display_name,
        "description": info.description,
        "available": is_backend_available(info.name),
        "capabilities": capabilities,
        "builtinTools": list(info.builtin_tools),
        "requiredKeys": list(info.required_keys),
        "supportedProviders": list(info.supported_providers),
        "installHint": info.install_hint,
        "beta": info.beta,
    }


def list_backend_summaries() -> list[dict[str, Any]]:
    """Return backend metadata for frontend/API consumers."""
    return [backend_summary(name) for name in list_backend_names()]


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


def install_hint_text(info: BackendInfo) -> str:
    """Render a compact install hint string for logs/errors."""
    hint = info.install_hint or {}
    pip_spec = hint.get("pip_spec")
    external_cmd = hint.get("external_cmd")
    if pip_spec and external_cmd:
        return f"pip install {pip_spec} and run: {external_cmd}"
    if pip_spec:
        return f"pip install {pip_spec}"
    if external_cmd:
        return external_cmd
    return "Install required backend dependencies"
