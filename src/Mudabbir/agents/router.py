"""Agent Router - routes to the selected agent backend.

Changes:
  - 2026-02-14: Graceful ImportError handling when a backend dep is missing.
  - 2026-02-02: Added claude_agent_sdk_full for 2-layer architecture.
  - 2026-02-02: Simplified - removed 2-layer mode (SDK has built-in execution).
  - 2026-02-02: Added Mudabbir_native - custom orchestrator with OI executor.
  - 2026-02-02: RE-ENABLED claude_agent_sdk - now uses official SDK properly!
                claude_code still disabled (homebrew pyautogui approach).
"""

import logging
from collections.abc import AsyncIterator

from Mudabbir.agents.registry import (
    create_backend,
    get_backend_info,
    install_hint_text,
    normalize_backend_name,
)
from Mudabbir.config import Settings

logger = logging.getLogger(__name__)


class AgentRouter:
    """Routes agent requests to the selected backend.

    ACTIVE backends:
    - claude_agent_sdk: Official Claude Agent SDK with all built-in tools (RECOMMENDED)
    - Mudabbir_native: Mudabbir's own brain + Open Interpreter hands
    - open_interpreter: Standalone Open Interpreter (local/cloud LLMs)

    DISABLED backends (for future use):
    - claude_code: Homebrew Claude + pyautogui (needs work)
    """

    def __init__(self, settings: Settings):
        self.settings = settings
        self._agent = None
        self._initialize_agent()

    def _initialize_agent(self) -> None:
        """Initialize the selected agent backend."""
        from Mudabbir.llm.client import resolve_llm_client

        requested_backend = self.settings.agent_backend
        backend = normalize_backend_name(requested_backend)
        self.settings.agent_backend = backend
        info = get_backend_info(backend)
        if requested_backend and requested_backend != backend:
            logger.warning("⚠️ Backend '%s' normalized to '%s'", requested_backend, backend)

        # Log Ollama usage
        llm = resolve_llm_client(self.settings)
        if llm.is_ollama:
            logger.info(
                "🦙 Ollama provider detected (%s) — using %s backend with local model",
                llm.ollama_host,
                backend,
            )
        if llm.is_openai_compatible:
            logger.info(
                "🔗 OpenAI-compatible provider detected (%s) — using %s backend",
                llm.openai_compatible_base_url,
                backend,
            )

        try:
            self._agent = create_backend(backend, self.settings)
            logger.info("🚀 Backend loaded: %s (%s)", info.display_name, info.name)
        except ImportError as exc:
            hint_text = install_hint_text(info)
            logger.error(
                f"Could not load '{backend}' backend — missing dependency: {exc}. "
                f"Install hint: {hint_text}"
            )
            if backend != "claude_agent_sdk":
                fallback = "claude_agent_sdk"
                fallback_info = get_backend_info(fallback)
                logger.warning("Falling back to '%s'", fallback)
                self.settings.agent_backend = fallback
                try:
                    self._agent = create_backend(fallback, self.settings)
                    logger.info("🚀 Backend loaded: %s (%s)", fallback_info.display_name, fallback)
                except ImportError as fallback_exc:
                    fallback_hint = install_hint_text(fallback_info)
                    logger.error(
                        "Fallback backend '%s' failed to load: %s. Install hint: %s",
                        fallback,
                        fallback_exc,
                        fallback_hint,
                    )
        except Exception as exc:
            logger.error("Could not initialize backend '%s': %s", backend, exc)

    async def run(
        self,
        message: str,
        *,
        system_prompt: str | None = None,
        history: list[dict] | None = None,
    ) -> AsyncIterator[dict]:
        """Run the agent with the given message.

        Args:
            message: User message to process.
            system_prompt: Dynamic system prompt from AgentContextBuilder.
            history: Recent session history as list of {"role": ..., "content": ...} dicts.

        Yields dicts with:
          - type: "message", "tool_use", "tool_result", "error", "done"
          - content: string content
          - metadata: optional dict with tool info (name, input)
        """
        if not self._agent:
            yield {"type": "error", "content": "❌ No agent initialized"}
            yield {"type": "done", "content": ""}
            return

        async for chunk in self._agent.run(message, system_prompt=system_prompt, history=history):
            yield chunk

    async def stop(self) -> None:
        """Stop the agent."""
        if self._agent:
            await self._agent.stop()
