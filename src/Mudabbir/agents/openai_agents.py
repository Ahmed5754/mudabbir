"""OpenAI Agents SDK backend for Mudabbir."""

from __future__ import annotations

import logging
from collections.abc import AsyncIterator
from typing import Any

from Mudabbir.agents.protocol import BackendInfo, Capability
from Mudabbir.config import Settings

logger = logging.getLogger(__name__)


class OpenAIAgentsBackend:
    """OpenAI Agents SDK backend with optional tool bridge integration."""

    @staticmethod
    def info() -> BackendInfo:
        return BackendInfo(
            name="openai_agents",
            display_name="OpenAI Agents",
            description="OpenAI Agents SDK backend with function tool bridge support.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            builtin_tools=("function_tools",),
            required_keys=("openai_api_key",),
            supported_providers=("openai", "ollama", "openai_compatible"),
            install_hint={
                "pip_package": "openai-agents",
                "pip_spec": "Mudabbir[openai-agents]",
                "verify_import": "agents",
            },
            beta=True,
        )

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._stop_flag = False
        self._sdk_available = False
        self._custom_tools: list | None = None
        self._initialize()

    def _initialize(self) -> None:
        try:
            import agents  # noqa: F401

            self._sdk_available = True
            logger.info("OpenAI Agents SDK backend ready")
        except ImportError:
            logger.warning(
                "OpenAI Agents SDK is not installed. Install with: pip install 'Mudabbir[openai-agents]'"
            )

    @staticmethod
    def _inject_history(instructions: str, history: list[dict]) -> str:
        lines = ["# Recent Conversation"]
        for msg in history:
            role = msg.get("role", "user").capitalize()
            content = str(msg.get("content", ""))
            if len(content) > 500:
                content = content[:500] + "..."
            lines.append(f"**{role}**: {content}")
        return instructions + "\n\n" + "\n".join(lines)

    def _build_custom_tools(self) -> list:
        if self._custom_tools is not None:
            return self._custom_tools
        try:
            from Mudabbir.agents.tool_bridge import build_openai_function_tools

            self._custom_tools = build_openai_function_tools(self.settings)
        except Exception as exc:
            logger.debug("Failed to build OpenAI function tools: %s", exc)
            self._custom_tools = []
        return self._custom_tools

    def _build_model(self) -> Any:
        model_name = self.settings.openai_agents_model or self.settings.openai_model or "gpt-4o"
        provider = (
            getattr(self.settings, "openai_agents_provider", "") or self.settings.llm_provider
        )

        if provider in ("ollama", "openai_compatible") or (
            provider == "auto" and self.settings.openai_compatible_base_url
        ):
            from agents.models.openai_chatcompletions import OpenAIChatCompletionsModel
            from openai import AsyncOpenAI

            if provider == "ollama":
                base_url = self.settings.ollama_host.rstrip("/") + "/v1"
                model_name = self.settings.ollama_model or model_name
                client = AsyncOpenAI(base_url=base_url, api_key="ollama")
            else:
                base_url = self.settings.openai_compatible_base_url
                api_key = self.settings.openai_compatible_api_key or "none"
                model_name = self.settings.openai_compatible_model or model_name
                client = AsyncOpenAI(base_url=base_url, api_key=api_key)

            return OpenAIChatCompletionsModel(model=model_name, openai_client=client)

        return model_name

    async def run(
        self,
        message: str,
        *,
        system_prompt: str | None = None,
        history: list[dict] | None = None,
        session_key: str | None = None,
    ) -> AsyncIterator[dict]:
        del session_key  # reserved for future native session support

        if not self._sdk_available:
            yield {
                "type": "error",
                "content": (
                    "OpenAI Agents SDK is not installed.\n\n"
                    "Install with: pip install 'Mudabbir[openai-agents]'"
                ),
            }
            yield {"type": "done", "content": ""}
            return

        self._stop_flag = False

        try:
            from agents import Agent, Runner

            model = self._build_model()
            instructions = system_prompt or "You are Mudabbir, a helpful AI assistant."
            if history:
                instructions = self._inject_history(instructions, history)

            agent = Agent(
                name="Mudabbir",
                instructions=instructions,
                model=model,
                tools=self._build_custom_tools(),
            )

            run_kwargs: dict[str, Any] = {"input": message}
            max_turns = getattr(self.settings, "openai_agents_max_turns", 0) or 0
            if max_turns > 0:
                run_kwargs["max_turns"] = max_turns

            result = Runner.run_streamed(agent, **run_kwargs)

            async for event in result.stream_events():
                if self._stop_flag:
                    break

                event_type = getattr(event, "type", "")
                if event_type == "raw_response_event":
                    data = getattr(event, "data", None)
                    delta = getattr(data, "delta", None) or getattr(data, "text_delta", None)
                    if delta:
                        yield {"type": "message", "content": str(delta)}

                elif event_type == "run_item_stream_event":
                    item = getattr(event, "item", None)
                    item_type = getattr(item, "type", "")
                    if item_type == "tool_call_item":
                        yield {
                            "type": "tool_use",
                            "content": f"Using {getattr(item, 'name', 'tool')}...",
                            "metadata": {
                                "name": getattr(item, "name", "tool"),
                                "input": getattr(item, "arguments", {}) or {},
                            },
                        }
                    elif item_type == "tool_call_output_item":
                        output = getattr(item, "output", "")
                        yield {
                            "type": "tool_result",
                            "content": str(output)[:200],
                            "metadata": {"name": "tool"},
                        }
                    elif item_type == "message_output_item":
                        text = getattr(item, "text", "")
                        if text:
                            yield {"type": "message", "content": str(text)}

            yield {"type": "done", "content": ""}

        except Exception as exc:
            logger.error("OpenAI Agents backend error: %s", exc)
            yield {"type": "error", "content": f"OpenAI Agents backend error: {exc}"}
            yield {"type": "done", "content": ""}

    async def stop(self) -> None:
        self._stop_flag = True

    async def get_status(self) -> dict[str, Any]:
        return {
            "backend": "openai_agents",
            "available": self._sdk_available,
            "running": not self._stop_flag,
        }
