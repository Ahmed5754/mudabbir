"""GitHub Copilot SDK backend for Mudabbir."""

from __future__ import annotations

import asyncio
import logging
import shutil
from collections.abc import AsyncIterator
from typing import Any

from Mudabbir.agents.protocol import BackendInfo, Capability
from Mudabbir.config import Settings

logger = logging.getLogger(__name__)


class CopilotSDKBackend:
    """Python SDK backend for GitHub Copilot CLI."""

    @staticmethod
    def info() -> BackendInfo:
        return BackendInfo(
            name="copilot_sdk",
            display_name="Copilot SDK",
            description="GitHub Copilot SDK backend with optional BYOK providers.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
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
        )

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._stop_flag = False
        self._client: Any = None
        self._sessions: dict[str, Any] = {}
        self._cli_available = shutil.which("copilot") is not None
        self._sdk_available = False
        try:
            import copilot  # noqa: F401

            self._sdk_available = True
        except ImportError:
            self._sdk_available = False

    @staticmethod
    def _inject_history(instruction: str, history: list[dict]) -> str:
        lines = ["# Recent Conversation"]
        for msg in history:
            role = msg.get("role", "user").capitalize()
            content = str(msg.get("content", ""))
            if len(content) > 500:
                content = content[:500] + "..."
            lines.append(f"**{role}**: {content}")
        return instruction + "\n\n" + "\n".join(lines)

    async def _ensure_client(self) -> Any:
        if self._client is not None:
            return self._client
        from copilot import CopilotClient

        self._client = CopilotClient()
        await self._client.start()
        return self._client

    def _provider_config(self) -> dict[str, Any] | None:
        provider = self.settings.copilot_sdk_provider
        if provider == "openai":
            cfg: dict[str, Any] = {"type": "openai"}
            if self.settings.openai_compatible_base_url:
                cfg["base_url"] = self.settings.openai_compatible_base_url
            cfg["api_key"] = self.settings.openai_api_key or self.settings.openai_compatible_api_key
            return cfg
        if provider == "azure":
            cfg = {"type": "azure"}
            if self.settings.openai_compatible_base_url:
                cfg["base_url"] = self.settings.openai_compatible_base_url
            if self.settings.openai_api_key:
                cfg["api_key"] = self.settings.openai_api_key
            return cfg
        if provider == "anthropic":
            cfg = {"type": "anthropic"}
            if self.settings.anthropic_api_key:
                cfg["api_key"] = self.settings.anthropic_api_key
            return cfg
        return None

    async def run(
        self,
        message: str,
        *,
        system_prompt: str | None = None,
        history: list[dict] | None = None,
        session_key: str | None = None,
    ) -> AsyncIterator[dict]:
        if not self._cli_available:
            yield {
                "type": "error",
                "content": (
                    "Copilot CLI not found on PATH.\n\n"
                    "Install from: https://github.com/github/copilot-sdk"
                ),
            }
            yield {"type": "done", "content": ""}
            return
        if not self._sdk_available:
            yield {
                "type": "error",
                "content": (
                    "Copilot SDK is not installed.\n\n"
                    "Install with: pip install github-copilot-sdk"
                ),
            }
            yield {"type": "done", "content": ""}
            return

        self._stop_flag = False
        try:
            client = await self._ensure_client()

            prompt_parts: list[str] = []
            if system_prompt:
                prompt_parts.append(f"[System Instructions]\n{system_prompt}\n")
            if history:
                prompt_parts.append(self._inject_history("", history).strip())

            try:
                from Mudabbir.agents.tool_bridge import get_tool_instructions_compact

                tool_section = get_tool_instructions_compact(self.settings)
                if tool_section:
                    prompt_parts.append(tool_section)
            except Exception:
                pass

            prompt_parts.append(message)
            full_prompt = "\n\n".join(prompt_parts)

            model = self.settings.copilot_sdk_model or "gpt-4o"
            provider_config = self._provider_config()

            if session_key and session_key in self._sessions:
                session = self._sessions[session_key]
            else:
                session_opts: dict[str, Any] = {"model": model, "streaming": True}
                if system_prompt:
                    session_opts["system_message"] = system_prompt
                if provider_config:
                    session_opts["provider"] = provider_config
                session = await client.create_session(session_opts)
                if session_key:
                    self._sessions[session_key] = session

            queue: asyncio.Queue[dict | None] = asyncio.Queue()
            streamed_delta = False

            def on_event(event: Any) -> None:
                nonlocal streamed_delta
                event_type = _event_type(event)
                data = getattr(event, "data", event)

                if event_type == "assistant.message_delta":
                    delta = getattr(data, "delta_content", "") or ""
                    if delta:
                        streamed_delta = True
                        queue.put_nowait({"type": "message", "content": delta})
                elif event_type == "assistant.reasoning_delta":
                    delta = getattr(data, "delta_content", "") or ""
                    if delta:
                        queue.put_nowait({"type": "thinking", "content": delta})
                elif event_type == "assistant.message":
                    if not streamed_delta:
                        content = getattr(data, "content", "") or ""
                        if content:
                            queue.put_nowait({"type": "message", "content": content})
                    streamed_delta = False
                elif event_type == "tool.call":
                    name = getattr(data, "name", "tool")
                    args = getattr(data, "arguments", {}) or {}
                    queue.put_nowait(
                        {
                            "type": "tool_use",
                            "content": f"Using: {name}",
                            "metadata": {"name": name, "input": args},
                        }
                    )
                elif event_type == "tool.result":
                    name = getattr(data, "name", "tool")
                    output = getattr(data, "output", "")
                    queue.put_nowait(
                        {
                            "type": "tool_result",
                            "content": str(output)[:200],
                            "metadata": {"name": name},
                        }
                    )
                elif event_type == "error":
                    queue.put_nowait(
                        {
                            "type": "error",
                            "content": getattr(data, "message", "Copilot SDK error"),
                        }
                    )
                elif event_type == "session.idle":
                    queue.put_nowait(None)

            session.on(on_event)
            await session.send({"prompt": full_prompt})

            max_turns = self.settings.copilot_sdk_max_turns or 0
            turn_count = 0
            while not self._stop_flag:
                item = await queue.get()
                if item is None:
                    break
                yield item
                if item.get("type") == "tool_result":
                    turn_count += 1
                    if max_turns and turn_count >= max_turns:
                        yield {"type": "error", "content": f"Reached max turns ({max_turns})"}
                        break

            yield {"type": "done", "content": ""}

        except Exception as exc:
            logger.error("Copilot SDK backend error: %s", exc)
            yield {"type": "error", "content": f"Copilot SDK backend error: {exc}"}
            yield {"type": "done", "content": ""}

    async def stop(self) -> None:
        self._stop_flag = True
        for session in self._sessions.values():
            try:
                await session.destroy()
            except Exception:
                pass
        self._sessions.clear()

        if self._client is not None:
            try:
                await self._client.stop()
            except Exception:
                pass
            self._client = None

    async def get_status(self) -> dict[str, Any]:
        return {
            "backend": "copilot_sdk",
            "cli_available": self._cli_available,
            "sdk_available": self._sdk_available,
            "running": self._client is not None,
            "active_sessions": len(self._sessions),
        }


def _event_type(event: Any) -> str:
    raw = getattr(event, "type", "")
    return raw.value if hasattr(raw, "value") else str(raw)
