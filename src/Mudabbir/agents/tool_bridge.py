"""Tool bridge for adapting Mudabbir tools to external SDK function-tool APIs."""

from __future__ import annotations

import json
import logging
from typing import Any

from Mudabbir.tools.policy import ToolPolicy
from Mudabbir.tools.protocol import BaseTool
from Mudabbir.tools.registry import ToolRegistry

logger = logging.getLogger(__name__)

_EXCLUDED_TOOLS = frozenset(
    {
        "ShellTool",
        "ReadFileTool",
        "WriteFileTool",
        "ListDirTool",
        "BrowserTool",
        "DesktopTool",
        "ScreenshotTool",
        "StatusTool",
    }
)


def _instantiate_all_tools() -> list[BaseTool]:
    """Discover and instantiate builtin tools except SDK/native-equivalent tools."""
    from Mudabbir.tools.builtin import _LAZY_IMPORTS

    tools: list[BaseTool] = []
    for class_name, (module_path, attr_name) in _LAZY_IMPORTS.items():
        if class_name in _EXCLUDED_TOOLS:
            continue
        try:
            import importlib

            mod = importlib.import_module(module_path, "Mudabbir.tools.builtin")
            cls = getattr(mod, attr_name)
            tools.append(cls())
        except Exception as exc:
            logger.debug("Skipping tool %s: %s", class_name, exc)
    return tools


def _build_registry(settings: Any) -> ToolRegistry:
    policy = ToolPolicy(
        profile=settings.tool_profile,
        allow=settings.tools_allow,
        deny=settings.tools_deny,
    )
    registry = ToolRegistry(policy=policy)
    for tool in _instantiate_all_tools():
        registry.register(tool)
    return registry


def _make_invoke_callback(tool: Any):
    """Create an async callback per tool to avoid closure capture bugs."""

    async def callback(_ctx: Any, args: str) -> str:
        try:
            params = json.loads(args) if args else {}
        except (json.JSONDecodeError, TypeError):
            return f"Error: invalid JSON arguments for {tool.name}: {args!r}"

        if not isinstance(params, dict):
            return f"Error: arguments must be a JSON object, got {type(params).__name__}"

        try:
            return await tool.execute(**params)
        except Exception as exc:
            logger.error("Tool %s execution error: %s", tool.name, exc)
            return f"Error executing {tool.name}: {exc}"

    return callback


def build_openai_function_tools(settings: Any) -> list:
    """Build OpenAI Agents SDK ``FunctionTool`` wrappers from Mudabbir tools."""
    try:
        from agents import FunctionTool
    except ImportError:
        logger.debug("OpenAI Agents SDK is not installed")
        return []

    registry = _build_registry(settings)
    function_tools: list[FunctionTool] = []
    for tool_name in registry.allowed_tool_names:
        tool = registry.get(tool_name)
        if tool is None:
            continue
        defn = tool.definition
        function_tools.append(
            FunctionTool(
                name=defn.name,
                description=defn.description,
                params_json_schema=defn.parameters,
                on_invoke_tool=_make_invoke_callback(tool),
            )
        )
    logger.info("Built %d OpenAI function tools", len(function_tools))
    return function_tools


def _make_adk_wrapper(tool: Any):
    """Build an ADK-compatible callable with introspectable signature."""
    import inspect

    defn = tool.definition
    props = (defn.parameters or {}).get("properties", {})
    param_names = list(props.keys())

    async def _adk_tool_wrapper(**kwargs: str) -> str:
        try:
            return await tool.execute(**kwargs)
        except Exception as exc:
            logger.error("ADK tool %s execution error: %s", tool.name, exc)
            return f"Error executing {tool.name}: {exc}"

    _adk_tool_wrapper.__name__ = defn.name
    _adk_tool_wrapper.__qualname__ = defn.name
    _adk_tool_wrapper.__doc__ = defn.description
    _adk_tool_wrapper.__signature__ = inspect.Signature(
        parameters=[
            inspect.Parameter(name, inspect.Parameter.KEYWORD_ONLY, annotation=str)
            for name in param_names
        ],
        return_annotation=str,
    )
    _adk_tool_wrapper.__annotations__ = {name: str for name in param_names}
    _adk_tool_wrapper.__annotations__["return"] = str
    return _adk_tool_wrapper


def build_adk_function_tools(settings: Any) -> list:
    """Build Google ADK ``FunctionTool`` wrappers from Mudabbir tools."""
    try:
        from google.adk.tools import FunctionTool
    except ImportError:
        logger.debug("Google ADK is not installed")
        return []

    registry = _build_registry(settings)
    function_tools: list = []
    for tool_name in registry.allowed_tool_names:
        tool = registry.get(tool_name)
        if tool is None:
            continue
        function_tools.append(FunctionTool(func=_make_adk_wrapper(tool)))
    logger.info("Built %d Google ADK function tools", len(function_tools))
    return function_tools


def get_tool_instructions_compact(settings: Any) -> str:
    """Return a compact markdown tool guide for prompt injection."""
    registry = _build_registry(settings)
    allowed = sorted(registry.allowed_tool_names)
    if not allowed:
        return ""

    lines = [
        "# Mudabbir Tools",
        "",
        "Use tools through the CLI wrapper:",
        "`python -m Mudabbir.tools.cli <tool_name> '<json_args>'`",
        "",
    ]
    for tool_name in allowed:
        tool = registry.get(tool_name)
        if not tool:
            continue
        desc = tool.definition.description.split(".")[0]
        lines.append(f"- `{tool_name}` - {desc}")
    lines.append("")
    lines.append(f"Total tools: {len(allowed)}")
    return "\n".join(lines)
