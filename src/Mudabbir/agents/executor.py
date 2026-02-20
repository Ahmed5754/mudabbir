"""Open Interpreter Executor - The "Hands" layer for OS control."""

from __future__ import annotations

import asyncio
import importlib
import logging
import os
import re
from dataclasses import dataclass
from pathlib import Path

from Mudabbir.config import Settings

logger = logging.getLogger(__name__)

_MAX_COMPLEX_CHUNKS = 220
_MAX_COMPLEX_TEXT_CHARS = 6000
_EXECUTOR_GUARDRAILS = """
Execution output policy:
- Do not include markdown code fences or planning prose.
- Do not include JSON execute payload fragments in user-facing text.
- Return concise execution result only.
- If action fails, return one short error line with the root cause.
"""
_PERCENT_RE = re.compile(r"(\d{1,3})\s*%")


@dataclass
class ExecutionResult:
    """Structured result contract for computer execution."""

    status: str
    action: str
    evidence: str = ""
    final_value: str = ""
    error: str = ""
    error_code: str = ""


def _looks_like_planning_text(text: str) -> bool:
    lowered = (text or "").strip().lower()
    if not lowered:
        return False
    patterns = (
        "here's my plan",
        "here is my plan",
        "let's start by",
        "step 1",
        "step 2",
        "my apologies",
        "single block",
        "please copy",
        "خطة",
        "لنبدأ",
        "الخطوة 1",
        "سأقوم",
        "اعتذر",
    )
    if any(p in lowered for p in patterns):
        return True
    if re.match(r"^\d+\.\s+\*\*?.+\*\*?:", lowered):
        return True
    return False


def _extract_percent_value(text: str) -> str:
    if not text:
        return ""
    match = _PERCENT_RE.search(text)
    if not match:
        return ""
    try:
        value = int(match.group(1))
    except Exception:
        return ""
    if 0 <= value <= 100:
        return f"{value}%"
    return ""


def _format_execution_summary(result: ExecutionResult) -> str:
    """Backward-compatible text summary for existing callers."""
    if result.status == "ok":
        suffix = f" (final: {result.final_value})" if result.final_value else ""
        evidence = f" | {result.evidence}" if result.evidence else ""
        return f"✅ {result.action} succeeded{suffix}{evidence}"
    if result.error_code:
        return f"❌ {result.action} failed [{result.error_code}]: {result.error or 'unknown error'}"
    return f"❌ {result.action} failed: {result.error or 'unknown error'}"


class OpenInterpreterExecutor:
    """Open Interpreter as the executor layer.

    Implements ExecutorProtocol - handles actual OS operations:
    - Shell commands
    - File read/write
    - Directory listing

    Used by orchestrators (Claude Agent SDK) to execute tool calls.
    """

    def __init__(self, settings: Settings):
        self.settings = settings
        self._interpreter = None
        self._initialize()

    def _initialize(self) -> None:
        """Initialize Open Interpreter instance."""
        try:
            from interpreter import interpreter

            from Mudabbir.llm.client import resolve_llm_client

            # Configure for execution mode (minimal LLM usage)
            interpreter.auto_run = True
            interpreter.loop = False  # Single command execution
            interpreter.system_message = (
                f"{getattr(interpreter, 'system_message', '')}\n{_EXECUTOR_GUARDRAILS}".strip()
            )
            try:
                interpreter.display_message = lambda *args, **kwargs: None
            except Exception:
                pass
            # Compatibility shim for OI builds where respond references display_markdown_message.
            try:
                respond_mod = importlib.import_module("interpreter.core.respond")
                if not hasattr(respond_mod, "display_markdown_message"):
                    try:
                        from interpreter.terminal_interface.utils.display_markdown_message import (
                            display_markdown_message as _display_markdown_message,
                        )
                    except Exception:

                        def _display_markdown_message(msg):
                            interpreter.display_message(str(msg))

                    respond_mod.display_markdown_message = _display_markdown_message
            except Exception as patch_err:
                logger.debug("Open Interpreter compatibility shim skipped: %s", patch_err)

            # Set LLM for any reasoning needed
            llm = resolve_llm_client(self.settings)
            if llm.is_ollama:
                interpreter.llm.model = f"ollama/{llm.model}"
                interpreter.llm.api_base = llm.ollama_host
            elif llm.is_gemini and llm.api_key:
                model_name = llm.model if llm.model.startswith("openai/") else f"openai/{llm.model}"
                interpreter.llm.model = model_name
                interpreter.llm.api_key = llm.api_key
                interpreter.llm.api_base = llm.openai_compatible_base_url
                if str(llm.model).lower().startswith("gemini-3"):
                    # Gemini 3 models are more stable at temperature=1.0.
                    try:
                        interpreter.llm.temperature = 1.0
                    except Exception:
                        pass
                os.environ.setdefault("GOOGLE_API_KEY", llm.api_key)
                os.environ.setdefault("GEMINI_API_KEY", llm.api_key)
            elif llm.api_key:
                interpreter.llm.model = llm.model
                interpreter.llm.api_key = llm.api_key

            self._interpreter = interpreter
            logger.info("=" * 50)
            logger.info("🔧 EXECUTOR: Open Interpreter initialized")
            logger.info("   └─ Role: Shell, files, system commands")
            logger.info("=" * 50)

        except ImportError:
            logger.error("❌ Open Interpreter not installed")
            self._interpreter = None
        except Exception as e:
            logger.error(f"❌ Failed to initialize executor: {e}")
            self._interpreter = None

    async def run_shell(self, command: str) -> str:
        """Execute a shell command and return output.

        Uses DIRECT subprocess execution for speed (not OI chat).
        This makes simple commands like 'ls', 'git status' ~10x faster.
        """
        logger.info(f"🔧 EXECUTOR: run_shell({command[:50]}...)")

        try:
            # Direct async subprocess - FAST
            proc = await asyncio.create_subprocess_shell(
                command,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                cwd=str(Path.home()),  # Default to home directory
            )

            # Wait with timeout
            try:
                stdout, stderr = await asyncio.wait_for(
                    proc.communicate(),
                    timeout=60.0,  # 60 second timeout
                )
            except asyncio.TimeoutError:
                proc.kill()
                return "Error: Command timed out after 60 seconds"

            # Combine output
            output = stdout.decode("utf-8", errors="replace")
            if stderr:
                err_text = stderr.decode("utf-8", errors="replace")
                if err_text.strip():
                    output += f"\n[stderr]: {err_text}"

            return output if output.strip() else "(no output)"

        except Exception as e:
            logger.error(f"Shell execution error: {e}")
            return f"Error: {str(e)}"

    async def run_complex_task(self, task: str) -> str:
        """Execute a complex multi-step task using Open Interpreter.

        Use this for tasks that need:
        - Multi-step reasoning
        - Code generation
        - Browser automation
        - AppleScript/Python for app queries

        For simple shell commands, use run_shell() instead.
        """
        result = await self.run_complex_task_struct(task)
        return _format_execution_summary(result)

    async def run_complex_task_struct(self, task: str) -> ExecutionResult:
        """Execute a complex task and return a structured summary."""
        if not self._interpreter:
            return ExecutionResult(
                status="error",
                action="computer task",
                error="Open Interpreter not available",
                error_code="oi_unavailable",
            )

        logger.info("🤖 EXECUTOR: run_complex_task(%s...)", task[:80])
        try:
            loop = asyncio.get_event_loop()
            result = await loop.run_in_executor(None, self._run_interpreter_sync, task)
            return result
        except Exception as e:
            logger.error("Complex task error: %s", e)
            return ExecutionResult(
                status="error",
                action="computer task",
                error=str(e),
                error_code="executor_exception",
            )

    def _run_interpreter_sync(self, task: str) -> ExecutionResult:
        """Synchronous Open Interpreter execution for complex tasks."""
        from Mudabbir.agents.open_interpreter import (
            _is_quota_or_rate_limit_message,
            is_noisy_execution_text,
        )

        output_parts: list[str] = []
        raw_error = ""
        chunks_seen = 0

        for chunk in self._interpreter.chat(task, stream=True):
            chunks_seen += 1
            if chunks_seen > _MAX_COMPLEX_CHUNKS:
                raw_error = "execution stream exceeded safety chunk limit"
                break

            if isinstance(chunk, dict):
                content = str(chunk.get("content", "") or "")
            else:
                content = str(chunk or "")

            if not content.strip():
                continue
            if is_noisy_execution_text(content) or _looks_like_planning_text(content):
                continue
            if "```" in content:
                continue

            output_parts.append(content.strip())
            merged = " ".join(output_parts)
            if len(merged) > _MAX_COMPLEX_TEXT_CHARS:
                output_parts = [merged[-_MAX_COMPLEX_TEXT_CHARS:]]

            lowered = content.lower()
            if "error" in lowered or "traceback" in lowered:
                raw_error = content.strip()

        merged_output = " ".join(output_parts).strip()
        if _is_quota_or_rate_limit_message(merged_output) or _is_quota_or_rate_limit_message(raw_error):
            return ExecutionResult(
                status="error",
                action="computer task",
                evidence="",
                error="quota/rate limit reached",
                error_code="quota_exhausted",
            )

        if raw_error and not merged_output:
            return ExecutionResult(
                status="error",
                action="computer task",
                evidence="",
                error=raw_error[:220],
                error_code="execution_error",
            )

        if not merged_output:
            return ExecutionResult(
                status="error",
                action="computer task",
                evidence="",
                error="no reliable execution output",
                error_code="no_output",
            )

        final_value = _extract_percent_value(merged_output)
        return ExecutionResult(
            status="ok",
            action="computer task",
            evidence=merged_output[:220],
            final_value=final_value,
        )

    async def read_file(self, path: str) -> str:
        """Read file contents."""
        logger.info(f"🔧 EXECUTOR: read_file({path})")

        try:
            # Direct file read - no need for interpreter
            with open(path, "r") as f:
                return f.read()
        except Exception as e:
            logger.error(f"File read error: {e}")
            return f"Error reading file: {str(e)}"

    async def write_file(self, path: str, content: str) -> None:
        """Write content to file."""
        logger.info(f"🔧 EXECUTOR: write_file({path})")

        try:
            with open(path, "w") as f:
                f.write(content)
        except Exception as e:
            logger.error(f"File write error: {e}")
            raise

    async def list_directory(self, path: str) -> list[str]:
        """List directory contents."""
        logger.info(f"🔧 EXECUTOR: list_directory({path})")

        import os

        try:
            return os.listdir(path)
        except Exception as e:
            logger.error(f"Directory list error: {e}")
            return []
