"""Google ADK backend for Mudabbir."""

from __future__ import annotations

import logging
import os
from collections.abc import AsyncIterator
from typing import Any

from Mudabbir.agents.protocol import BackendInfo, Capability
from Mudabbir.config import Settings
from Mudabbir.tools.policy import ToolPolicy

logger = logging.getLogger(__name__)

_APP_NAME = "Mudabbir"


class GoogleADKBackend:
    """Google ADK backend with optional MCP and tool bridge support."""

    @staticmethod
    def info() -> BackendInfo:
        return BackendInfo(
            name="google_adk",
            display_name="Google ADK",
            description="Google Agent Development Kit backend with Gemini and MCP integration.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MCP
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            builtin_tools=("google_search", "code_execution"),
            required_keys=("google_api_key",),
            supported_providers=("gemini",),
            install_hint={
                "pip_package": "google-adk",
                "pip_spec": "Mudabbir[google-adk]",
                "verify_import": "google.adk",
            },
            beta=True,
        )

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._stop_flag = False
        self._sdk_available = False
        self._sessions: dict[str, str] = {}
        self._custom_tools: list | None = None
        self._initialize()

    def _initialize(self) -> None:
        try:
            import google.adk  # noqa: F401

            self._sdk_available = True
            logger.info("Google ADK backend ready")
        except ImportError:
            logger.warning(
                "Google ADK is not installed. Install with: pip install 'Mudabbir[google-adk]'"
            )
            return

        if self.settings.google_api_key:
            os.environ["GOOGLE_API_KEY"] = self.settings.google_api_key
        os.environ["GOOGLE_GENAI_USE_VERTEXAI"] = "FALSE"

    def _build_custom_tools(self) -> list:
        if self._custom_tools is not None:
            return self._custom_tools
        try:
            from Mudabbir.agents.tool_bridge import build_adk_function_tools

            self._custom_tools = build_adk_function_tools(self.settings)
        except Exception as exc:
            logger.debug("Failed to build ADK tools: %s", exc)
            self._custom_tools = []
        return self._custom_tools

    def _build_mcp_toolsets(self) -> list:
        try:
            from google.adk.tools.mcp_tool import McpToolset
            from google.adk.tools.mcp_tool.mcp_session_manager import (
                SseConnectionParams,
                StdioConnectionParams,
            )
            from mcp import StdioServerParameters
        except ImportError:
            return []

        try:
            from Mudabbir.mcp.config import load_mcp_config
        except Exception:
            return []

        policy = ToolPolicy(
            profile=self.settings.tool_profile,
            allow=self.settings.tools_allow,
            deny=self.settings.tools_deny,
        )
        toolsets: list = []
        for cfg in load_mcp_config():
            if not policy.is_mcp_server_allowed(cfg.name):
                continue
            try:
                if cfg.transport == "stdio":
                    toolsets.append(
                        McpToolset(
                            connection_params=StdioConnectionParams(
                                server_params=StdioServerParameters(
                                    command=cfg.command,
                                    args=cfg.args or [],
                                    env=cfg.env,
                                )
                            )
                        )
                    )
                elif cfg.transport in ("sse", "http") and cfg.url:
                    toolsets.append(
                        McpToolset(
                            connection_params=SseConnectionParams(
                                url=cfg.url,
                                headers=cfg.headers or {},
                            )
                        )
                    )
            except Exception as exc:
                logger.debug("Skipping MCP server %s: %s", cfg.name, exc)
        return toolsets

    def _get_runner(self, instruction: str, tools: list):
        from google.adk.agents import LlmAgent
        from google.adk.runners import InMemoryRunner

        model = self.settings.google_adk_model or "gemini-2.5-flash"
        agent = LlmAgent(
            name="Mudabbir",
            model=model,
            instruction=instruction,
            tools=tools,
        )
        return InMemoryRunner(agent=agent, app_name=_APP_NAME)

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

    async def run(
        self,
        message: str,
        *,
        system_prompt: str | None = None,
        history: list[dict] | None = None,
        session_key: str | None = None,
    ) -> AsyncIterator[dict]:
        if not self._sdk_available:
            yield {
                "type": "error",
                "content": (
                    "Google ADK is not installed.\n\n"
                    "Install with: pip install 'Mudabbir[google-adk]'"
                ),
            }
            yield {"type": "done", "content": ""}
            return

        self._stop_flag = False

        try:
            from google.genai import types

            instruction = system_prompt or "You are Mudabbir, a helpful AI assistant."
            if history:
                instruction = self._inject_history(instruction, history)

            tools = self._build_custom_tools() + self._build_mcp_toolsets()
            runner = self._get_runner(instruction, tools)

            user_id = "mudabbir_user"
            if session_key and session_key in self._sessions:
                adk_session_id = self._sessions[session_key]
            else:
                import uuid

                adk_session_id = str(uuid.uuid4())
                if session_key:
                    self._sessions[session_key] = adk_session_id

            await runner.session_service.create_session(
                app_name=_APP_NAME,
                user_id=user_id,
                session_id=adk_session_id,
            )

            user_message = types.Content(role="user", parts=[types.Part(text=message)])

            run_kwargs: dict[str, Any] = {
                "user_id": user_id,
                "session_id": adk_session_id,
                "new_message": user_message,
            }
            try:
                from google.adk.agents.run_config import RunConfig, StreamingMode

                run_kwargs["run_config"] = RunConfig(streaming_mode=StreamingMode.SSE)
            except Exception:
                pass

            streamed_partial = False
            max_turns = getattr(self.settings, "google_adk_max_turns", 0) or 0
            turn_count = 0

            async for event in runner.run_async(**run_kwargs):
                if self._stop_flag:
                    break

                if max_turns and turn_count >= max_turns:
                    yield {"type": "error", "content": f"Max turns ({max_turns}) reached"}
                    break

                content = getattr(event, "content", None)
                parts = getattr(content, "parts", None) if content else None
                if not parts:
                    continue

                is_partial = getattr(event, "partial", None) is True
                for part in parts:
                    if self._stop_flag:
                        break

                    if getattr(part, "text", None):
                        if is_partial:
                            streamed_partial = True
                            yield {"type": "message", "content": part.text}
                        elif not streamed_partial:
                            yield {"type": "message", "content": part.text}
                        else:
                            streamed_partial = False

                    elif getattr(part, "function_call", None):
                        turn_count += 1
                        fc = part.function_call
                        yield {
                            "type": "tool_use",
                            "content": f"Using {fc.name}...",
                            "metadata": {
                                "name": fc.name,
                                "input": dict(fc.args) if fc.args else {},
                            },
                        }

                    elif getattr(part, "function_response", None):
                        fr = part.function_response
                        response = fr.response if fr.response is not None else ""
                        yield {
                            "type": "tool_result",
                            "content": str(response)[:200],
                            "metadata": {"name": fr.name or "tool"},
                        }

            yield {"type": "done", "content": ""}

        except Exception as exc:
            logger.error("Google ADK backend error: %s", exc)
            yield {"type": "error", "content": f"Google ADK backend error: {exc}"}
            yield {"type": "done", "content": ""}

    async def stop(self) -> None:
        self._stop_flag = True

    async def get_status(self) -> dict[str, Any]:
        return {
            "backend": "google_adk",
            "available": self._sdk_available,
            "running": not self._stop_flag,
            "active_sessions": len(self._sessions),
        }
