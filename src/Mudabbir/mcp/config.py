"""MCP server configuration — load/save from ~/.Mudabbir/mcp_servers.json.

Created: 2026-02-07
"""

from __future__ import annotations

import json
import logging
import os
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from Mudabbir.config import get_config_dir

logger = logging.getLogger(__name__)

MCP_CONFIG_FILENAME = "mcp_servers.json"


@dataclass
class MCPServerConfig:
    """Configuration for a single MCP server."""

    name: str
    transport: str = "stdio"  # "stdio", "http" (SSE), or "streamable-http"
    command: str = ""  # For stdio: executable command
    args: list[str] = field(default_factory=list)  # For stdio: command arguments
    url: str = ""  # For http: server URL
    env: dict[str, str] = field(default_factory=dict)
    enabled: bool = True
    timeout: int = 30  # Connection timeout in seconds
    # Legacy: original registry identifier (e.g. "@cmd8/excalidraw-mcp@0.1.4").
    # Kept for backward compatibility with servers installed from the now-removed
    # MCP Registry tab. Not used by new installations.
    registry_ref: str = ""
    oauth: bool = False  # True if server uses OAuth authentication

    def to_dict(self) -> dict:
        d = {
            "name": self.name,
            "transport": self.transport,
            "command": self.command,
            "args": self.args,
            "url": self.url,
            "env": self.env,
            "enabled": self.enabled,
            "timeout": self.timeout,
        }
        if self.registry_ref:
            d["registry_ref"] = self.registry_ref
        if self.oauth:
            d["oauth"] = True
        return d

    @classmethod
    def from_dict(cls, data: dict) -> MCPServerConfig:
        return cls(
            name=data.get("name", ""),
            transport=data.get("transport", "stdio"),
            command=data.get("command", ""),
            args=data.get("args", []),
            url=data.get("url", ""),
            env=data.get("env", {}),
            enabled=data.get("enabled", True),
            timeout=data.get("timeout", 30),
            registry_ref=data.get("registry_ref", ""),
            oauth=data.get("oauth", False),
        )


def _get_mcp_config_path() -> Path:
    return get_config_dir() / MCP_CONFIG_FILENAME


def load_mcp_config() -> list[MCPServerConfig]:
    """Load MCP server configs from disk."""
    path = _get_mcp_config_path()
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text())
        servers = data.get("servers", [])
        return [MCPServerConfig.from_dict(s) for s in servers]
    except (json.JSONDecodeError, Exception) as e:
        logger.warning("Failed to load MCP config: %s", e)
        return []


def save_mcp_config(configs: list[MCPServerConfig]) -> None:
    """Save MCP server configs to disk."""
    path = _get_mcp_config_path()
    data = {"servers": [c.to_dict() for c in configs]}
    path.write_text(json.dumps(data, indent=2))
    logger.info("Saved %d MCP server configs", len(configs))


def cleanup_invalid_mcp_configs() -> dict[str, int]:
    """Auto-clean invalid MCP configs.

    - Deduplicates by server name (keeps latest definition).
    - Removes stdio servers with missing/invalid command.
    - Removes preset configs missing required env keys.
    - Saves only when changes are detected.
    """
    configs = load_mcp_config()
    if not configs:
        return {"removed": 0, "deduplicated": 0, "invalid_stdio": 0, "invalid_env": 0}

    original_count = len(configs)

    # Deduplicate by name (latest wins)
    by_name: dict[str, MCPServerConfig] = {}
    for cfg in configs:
        by_name[cfg.name] = cfg
    deduped = list(by_name.values())
    deduplicated = original_count - len(deduped)

    cleaned: list[MCPServerConfig] = []
    invalid_stdio = 0
    invalid_env = 0

    # Lazy import to avoid module cycle (presets -> config).
    get_preset = None
    try:
        from Mudabbir.mcp.presets import get_preset as _get_preset

        get_preset = _get_preset
    except Exception:
        get_preset = None

    for cfg in deduped:
        if cfg.transport == "stdio":
            cmd = (cfg.command or "").strip()
            if not cmd or shutil.which(cmd) is None:
                invalid_stdio += 1
                logger.warning(
                    "Removing invalid MCP stdio config '%s' (missing command: %r)",
                    cfg.name,
                    cfg.command,
                )
                continue

        if get_preset is not None:
            preset = get_preset(cfg.name)
            if preset and preset.env_keys:
                missing_required = [
                    ek.key
                    for ek in preset.env_keys
                    if ek.required
                    and not (cfg.env or {}).get(ek.key)
                    and not os.getenv(ek.key)
                ]
                if missing_required:
                    invalid_env += 1
                    logger.warning(
                        "Removing invalid MCP config '%s' (missing required env vars: %s)",
                        cfg.name,
                        ", ".join(missing_required),
                    )
                    continue
        cleaned.append(cfg)

    removed = deduplicated + invalid_stdio + invalid_env
    if removed > 0:
        save_mcp_config(cleaned)
        logger.info(
            "MCP config cleanup complete: removed=%d (deduplicated=%d, invalid_stdio=%d, invalid_env=%d)",
            removed,
            deduplicated,
            invalid_stdio,
            invalid_env,
        )

    return {
        "removed": removed,
        "deduplicated": deduplicated,
        "invalid_stdio": invalid_stdio,
        "invalid_env": invalid_env,
    }
