"""Codex CLI backend for Mudabbir."""

from __future__ import annotations

import asyncio
import json
import logging
import shutil
from collections.abc import AsyncIterator
from typing import Any

from Mudabbir.agents.protocol import BackendInfo, Capability
from Mudabbir.config import Settings

logger = logging.getLogger(__name__)


class CodexCLIBackend:
    """Subprocess wrapper for OpenAI Codex CLI JSON stream output."""

    @staticmethod
    def info() -> BackendInfo:
        return BackendInfo(
            name="codex_cli",
            display_name="Codex CLI",
            description="Codex CLI subprocess backend with JSON event parsing.",
            capabilities=(
                Capability.STREAMING
                | Capability.TOOLS
                | Capability.MULTI_TURN
                | Capability.CUSTOM_SYSTEM_PROMPT
            ),
            builtin_tools=("shell", "file_edit", "web_search"),
            required_keys=("openai_api_key",),
            supported_providers=("openai",),
            install_hint={
                "external_cmd": "npm install -g @openai/codex",
                "binary": "codex",
                "docs_url": "https://github.com/openai/codex",
            },
            beta=True,
        )

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self._stop_flag = False
        self._process: asyncio.subprocess.Process | None = None
        self._cli_available = shutil.which("codex") is not None

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
        del session_key

        if not self._cli_available:
            yield {
                "type": "error",
                "content": (
                    "Codex CLI not found on PATH.\n\n"
                    "Install with: npm install -g @openai/codex"
                ),
            }
            yield {"type": "done", "content": ""}
            return

        self._stop_flag = False

        try:
            prompt_parts: list[str] = []
            if system_prompt:
                prompt_parts.append(f"[System Instructions]\n{system_prompt}\n")
            if history:
                prompt_parts.append(self._inject_history("", history).strip())
            prompt_parts.append(message)
            full_prompt = "\n\n".join(prompt_parts)

            model = self.settings.codex_cli_model or "gpt-4o"
            cmd = ["codex", "exec", "--json", "--full-auto", "--model", model, full_prompt]

            self._process = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
            )
            if self._process.stdout is None:
                yield {"type": "error", "content": "Failed to capture Codex CLI output"}
                yield {"type": "done", "content": ""}
                return

            async for raw_line in self._process.stdout:
                if self._stop_flag:
                    break

                line = raw_line.decode("utf-8", errors="replace").strip()
                if not line:
                    continue
                try:
                    event_data = json.loads(line)
                except json.JSONDecodeError:
                    continue

                event_type = event_data.get("type", "")
                if event_type == "item.started":
                    item = event_data.get("item", {})
                    item_type = item.get("type", "")
                    if item_type == "command_execution":
                        cmd_str = item.get("command", "")
                        yield {
                            "type": "tool_use",
                            "content": f"Running: {cmd_str}",
                            "metadata": {"name": "shell", "input": {"command": cmd_str}},
                        }
                    elif item_type == "file_change":
                        filename = item.get("filename", "unknown")
                        yield {
                            "type": "tool_use",
                            "content": f"Editing: {filename}",
                            "metadata": {"name": "file_edit", "input": {"filename": filename}},
                        }
                    elif item_type == "web_search":
                        query = item.get("query", "")
                        yield {
                            "type": "tool_use",
                            "content": f"Searching: {query}",
                            "metadata": {"name": "web_search", "input": {"query": query}},
                        }
                elif event_type == "item.completed":
                    item = event_data.get("item", {})
                    item_type = item.get("type", "")
                    if item_type == "agent_message":
                        text = item.get("text", "")
                        if text:
                            yield {"type": "message", "content": text}
                    elif item_type == "command_execution":
                        output = item.get("output", "")
                        yield {
                            "type": "tool_result",
                            "content": str(output)[:200],
                            "metadata": {"name": "shell"},
                        }
                    elif item_type == "file_change":
                        filename = item.get("filename", "unknown")
                        yield {
                            "type": "tool_result",
                            "content": f"Updated {filename}",
                            "metadata": {"name": "file_edit"},
                        }
                    elif item_type == "web_search":
                        output = item.get("output", "")
                        yield {
                            "type": "tool_result",
                            "content": str(output)[:200],
                            "metadata": {"name": "web_search"},
                        }
                    elif item_type == "reasoning":
                        thought = item.get("text", "")
                        if thought:
                            yield {"type": "thinking", "content": thought}
                elif event_type == "error":
                    yield {"type": "error", "content": event_data.get("message", "Codex error")}

            await self._process.wait()
            if self._process.returncode not in (None, 0) and not self._stop_flag:
                stderr_output = ""
                if self._process.stderr:
                    stderr_output = (
                        (await self._process.stderr.read())
                        .decode("utf-8", errors="replace")
                        .strip()
                    )
                msg = f"Codex CLI exited with code {self._process.returncode}"
                if stderr_output:
                    msg += f": {stderr_output[:300]}"
                yield {"type": "error", "content": msg}

            self._process = None
            yield {"type": "done", "content": ""}

        except Exception as exc:
            logger.error("Codex CLI backend error: %s", exc)
            yield {"type": "error", "content": f"Codex CLI backend error: {exc}"}
            yield {"type": "done", "content": ""}

    async def stop(self) -> None:
        self._stop_flag = True
        if self._process and self._process.returncode is None:
            try:
                self._process.terminate()
            except ProcessLookupError:
                pass

    async def get_status(self) -> dict[str, Any]:
        return {
            "backend": "codex_cli",
            "available": self._cli_available,
            "running": self._process is not None and self._process.returncode is None,
            "model": self.settings.codex_cli_model or "gpt-4o",
        }
