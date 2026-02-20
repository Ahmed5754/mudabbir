"""OpenCode REST backend for Mudabbir."""

from __future__ import annotations

import logging
from collections.abc import AsyncIterator
from typing import Any

import httpx

from Mudabbir.agents.protocol import BackendInfo, Capability
from Mudabbir.config import Settings

logger = logging.getLogger(__name__)


class OpenCodeBackend:
    """Backend that talks to a running OpenCode server over REST."""

    @staticmethod
    def info() -> BackendInfo:
        return BackendInfo(
            name="opencode",
            display_name="OpenCode",
            description="OpenCode server backend via REST API.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            install_hint={
                "external_cmd": "go install github.com/opencode-ai/opencode@latest",
                "binary": "opencode",
                "docs_url": "https://github.com/opencode-ai/opencode",
            },
            beta=True,
        )

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._base_url = (settings.opencode_base_url or "http://localhost:4096").rstrip("/")
        self._stop_flag = False
        self._client: httpx.AsyncClient | None = None
        self._session_map: dict[str, str] = {}

    def _client_instance(self) -> httpx.AsyncClient:
        if self._client is None or self._client.is_closed:
            self._client = httpx.AsyncClient(base_url=self._base_url, timeout=None)
        return self._client

    async def _check_health(self) -> bool:
        try:
            resp = await self._client_instance().get("/")
            return resp.status_code < 500
        except Exception:
            return False

    async def _get_or_create_session(self, key: str = "_default") -> str:
        if key in self._session_map:
            return self._session_map[key]

        resp = await self._client_instance().post("/session")
        resp.raise_for_status()
        data = resp.json()
        session_id = data["id"]
        self._session_map[key] = session_id
        return session_id

    async def run(
        self,
        message: str,
        *,
        system_prompt: str | None = None,
        history: list[dict] | None = None,
        session_key: str | None = None,
    ) -> AsyncIterator[dict]:
        del history
        self._stop_flag = False

        if not await self._check_health():
            yield {
                "type": "error",
                "content": (
                    f"OpenCode server is unreachable at {self._base_url}.\n\n"
                    "Start it with: opencode --server"
                ),
            }
            yield {"type": "done", "content": ""}
            return

        try:
            key = session_key or "_default"
            session_id = await self._get_or_create_session(key)

            if self._stop_flag:
                yield {"type": "done", "content": ""}
                return

            effective_system = system_prompt or ""
            try:
                from Mudabbir.agents.tool_bridge import get_tool_instructions_compact

                tool_section = get_tool_instructions_compact(self.settings)
                if tool_section:
                    effective_system = (effective_system + "\n" + tool_section).strip()
            except Exception:
                pass

            payload: dict[str, Any] = {"parts": [{"type": "text", "text": message}]}
            if effective_system:
                payload["system"] = effective_system
            if self.settings.opencode_model:
                payload["model"] = self.settings.opencode_model

            resp = await self._client_instance().post(f"/session/{session_id}/message", json=payload)
            resp.raise_for_status()
            data = resp.json()

            parts = data.get("parts", []) if isinstance(data, dict) else []
            for part in parts:
                if self._stop_flag:
                    break
                part_type = part.get("type", "text")
                if part_type == "text":
                    text = part.get("text", "")
                    if text:
                        yield {"type": "message", "content": text}
                elif part_type == "tool":
                    tool = part.get("tool", {})
                    tool_name = tool.get("name", "tool") if isinstance(tool, dict) else str(tool)
                    yield {
                        "type": "tool_use",
                        "content": f"Using {tool_name}...",
                        "metadata": {"name": tool_name},
                    }
                    state = part.get("state", {})
                    if isinstance(state, dict) and state.get("output"):
                        yield {
                            "type": "tool_result",
                            "content": str(state["output"])[:200],
                            "metadata": {"name": tool_name},
                        }

            if not parts and isinstance(data, dict):
                fallback = data.get("content", "") or data.get("text", "")
                if fallback:
                    yield {"type": "message", "content": fallback}

            yield {"type": "done", "content": ""}

        except httpx.HTTPStatusError as exc:
            logger.error("OpenCode HTTP error: %s", exc)
            yield {"type": "error", "content": f"OpenCode server error: {exc.response.status_code}"}
            yield {"type": "done", "content": ""}
        except Exception as exc:
            logger.error("OpenCode backend error: %s", exc)
            yield {"type": "error", "content": f"OpenCode backend error: {exc}"}
            yield {"type": "done", "content": ""}

    async def stop(self) -> None:
        self._stop_flag = True
        if self._client and not self._client.is_closed:
            await self._client.aclose()
            self._client = None

    async def get_status(self) -> dict[str, Any]:
        return {
            "backend": "opencode",
            "base_url": self._base_url,
            "model": self.settings.opencode_model or "server-default",
            "active_sessions": len(self._session_map),
        }
