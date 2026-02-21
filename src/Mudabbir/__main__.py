"""Mudabbir entry point.

Changes:
  - 2026-02-17: Run startup health checks after settings load (prints colored summary).
  - 2026-02-16: Add startup version check against PyPI (cached daily, silent on error).
  - 2026-02-14: Dashboard deps moved to core — `pip install Mudabbir` just works.
  - 2026-02-12: Fixed --version to read dynamically from package metadata.
  - 2026-02-06: Web dashboard is now the default mode (no flags needed).
  - 2026-02-06: Added --telegram flag for legacy Telegram-only mode.
  - 2026-02-06: Added --discord, --slack, --whatsapp CLI modes.
  - 2026-02-02: Added Rich logging for beautiful console output.
  - 2026-02-03: Handle port-in-use gracefully with automatic port finding.
"""

import argparse
import asyncio
import importlib.util
import json
import logging
import os
import subprocess
import sys
import webbrowser
import warnings
from importlib.metadata import version as get_version
from importlib.metadata import PackageNotFoundError
from pathlib import Path

# Dependency noise filters (do not affect runtime behavior).
warnings.filterwarnings(
    "ignore",
    message=r"pkg_resources is deprecated as an API.*",
    category=UserWarning,
)
warnings.filterwarnings(
    "ignore",
    message=r".*asyncio\.iscoroutinefunction.*deprecated.*",
    category=DeprecationWarning,
    module=r"backoff\._decorator",
)
warnings.filterwarnings(
    "ignore",
    message=r"remove second argument of ws_handler",
    category=DeprecationWarning,
    module=r"websockets\.legacy\.server",
)


def _sanitize_sys_path() -> None:
    """Avoid cwd package-shadowing when launched inside Mudabbir package dir.

    If current working directory is `.../site-packages/Mudabbir`, top-level
    imports like `import mcp` can accidentally resolve to `Mudabbir/mcp`
    instead of the external `mcp` dependency package.
    """
    try:
        package_dir = Path(__file__).resolve().parent
        cwd = Path.cwd().resolve()
        if cwd != package_dir:
            return
        cleaned: list[str] = []
        for item in sys.path:
            try:
                if item in ("", "."):
                    continue
                if Path(item).resolve() == cwd:
                    continue
            except Exception:
                pass
            cleaned.append(item)
        sys.path[:] = cleaned
    except Exception:
        return


_sanitize_sys_path()

from Mudabbir.config import Settings, get_config_dir, get_config_path, get_settings  # noqa: E402
from Mudabbir.logging_setup import setup_logging  # noqa: E402


def _runtime_version() -> str:
    """Return installed package version with compatibility fallback."""
    for dist_name in ("Mudabbir", "mudabbir"):
        try:
            return get_version(dist_name)
        except PackageNotFoundError:
            continue
    return "0.0.0"


def _print_mem0_config_error(errors: list[str]) -> None:
    """Print a clear terminal diagnostic and fix hints, then exit."""
    config_path = get_config_path()
    print("\n" + "=" * 64)
    print("MEM0 CONFIG ERROR")
    print("=" * 64)
    print("Mudabbir stopped because memory_backend=mem0 has invalid settings.\n")
    for idx, err in enumerate(errors, start=1):
        print(f"{idx}. {err}")
    print("\nHow to fix:")
    print(f"- Open: {config_path}")
    print("- Set mem0_ollama_base_url to a valid endpoint URL")
    print("  Example: http://localhost:11434")
    print("- Keep model names in model fields only:")
    print("  mem0_llm_model / mem0_embedder_model")
    print("=" * 64 + "\n")

# Setup beautiful logging with Rich
setup_logging(level="INFO")
logger = logging.getLogger(__name__)

_AUTO_UPDATE_GUARD_ENV = "MUDABBIR_AUTO_UPDATE_DONE"
_AUTO_UPDATE_DISABLE_ENV = "MUDABBIR_DISABLE_AUTO_UPDATE"


def _pid_file_path() -> Path:
    return get_config_dir() / "runtime.pid"


def _write_runtime_pid() -> None:
    try:
        p = _pid_file_path()
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(str(os.getpid()), encoding="utf-8")
    except Exception:
        pass


def _clear_runtime_pid() -> None:
    try:
        p = _pid_file_path()
        if not p.exists():
            return
        content = p.read_text(encoding="utf-8").strip()
        if content == str(os.getpid()):
            p.unlink(missing_ok=True)
    except Exception:
        pass


def _stop_running_instances() -> int:
    """Best-effort stop for stuck Mudabbir processes.

    Returns number of processes targeted for stop.
    """
    current_pid = os.getpid()
    target_pids: set[int] = set()

    # 1) PID file (most reliable for normal single-instance usage).
    try:
        p = _pid_file_path()
        if p.exists():
            raw = p.read_text(encoding="utf-8").strip()
            pid = int(raw)
            if pid > 0 and pid != current_pid:
                target_pids.add(pid)
    except Exception:
        pass

    # 2) Windows process scan by commandline (covers stale/no pid-file cases).
    if os.name == "nt":
        ps = rf"""
$ErrorActionPreference='SilentlyContinue'
$self={current_pid}
$rows = Get-CimInstance Win32_Process | Where-Object {{
  $_.Name -eq 'python.exe' -and
  $_.ProcessId -ne $self -and
  ($_.CommandLine -match '(-m\s+Mudabbir\b|\\Mudabbir\\__main__\.py|\bmudabbir(\.exe)?\b)')
}} | Select-Object -ExpandProperty ProcessId
$rows | ConvertTo-Json -Compress
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps],
                capture_output=True,
                text=True,
                timeout=10,
            )
            out = (proc.stdout or "").strip()
            if out:
                parsed = json.loads(out)
                if isinstance(parsed, int):
                    if parsed != current_pid:
                        target_pids.add(parsed)
                elif isinstance(parsed, list):
                    for item in parsed:
                        try:
                            pid = int(item)
                            if pid != current_pid:
                                target_pids.add(pid)
                        except Exception:
                            continue
        except Exception:
            pass

    stopped = 0
    for pid in sorted(target_pids):
        try:
            if os.name == "nt":
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    capture_output=True,
                    text=True,
                    timeout=8,
                )
            else:
                os.kill(pid, 15)
            stopped += 1
        except Exception:
            continue

    try:
        if stopped > 0:
            _pid_file_path().unlink(missing_ok=True)
    except Exception:
        pass
    return stopped


def _should_auto_update_on_boot(argv: list[str]) -> bool:
    """Auto-update only for plain terminal launch: `Mudabbir`."""
    if len(argv) != 1:
        return False
    if os.environ.get(_AUTO_UPDATE_GUARD_ENV) == "1":
        return False
    if os.environ.get(_AUTO_UPDATE_DISABLE_ENV) == "1":
        return False
    return True


def _bootstrap_latest_and_restart(argv: list[str]) -> None:
    """Update to latest published version (if needed) then restart once."""
    from Mudabbir.config import get_config_dir
    from Mudabbir.update_check import check_for_updates

    info = check_for_updates(_runtime_version(), get_config_dir(), use_cache=False)
    if not info or not info.get("update_available"):
        return

    current = info.get("current", _runtime_version())
    latest = info.get("latest", "?")
    logger.info("🔄 Updating Mudabbir %s -> %s", current, latest)

    # Never run pip upgrade in-process on Windows; mudabbir.exe is locked while running.
    if os.name == "nt":
        logger.info("🔁 Switching to Windows handoff updater...")
        _handoff_windows_update_and_restart(argv)
        raise SystemExit(0)

    install_cmd = [
        sys.executable,
        "-m",
        "pip",
        "install",
        "--upgrade",
        "Mudabbir",
        "--disable-pip-version-check",
        "--no-input",
        "--quiet",
    ]
    install = subprocess.run(install_cmd, capture_output=True, text=True)
    if install.returncode != 0:
        last_error_line = (
            (install.stderr or install.stdout or "").strip().splitlines()[-1:]
        )
        compact_error = last_error_line[0] if last_error_line else "pip install failed"
        logger.error(
            "Auto-update failed: %s. Launching current installed version.",
            compact_error,
        )
        return

    env = os.environ.copy()
    env[_AUTO_UPDATE_GUARD_ENV] = "1"
    restart_cmd = [sys.executable, "-m", "Mudabbir", *argv[1:]]
    logger.info("✅ Update complete. Restarting Mudabbir...")
    restarted = subprocess.run(restart_cmd, env=env)
    raise SystemExit(restarted.returncode)


def _handoff_windows_update_and_restart(argv: list[str]) -> None:
    """On Windows, finish update in a child process after current process exits."""
    helper = r"""
import os
import subprocess
import sys
import time

args = sys.argv[1:]
time.sleep(0.8)

install_cmd = [
    sys.executable,
    "-m",
    "pip",
    "install",
    "--upgrade",
    "Mudabbir",
    "--disable-pip-version-check",
    "--no-input",
    "--quiet",
]
res = subprocess.run(install_cmd, capture_output=True, text=True)
if res.returncode == 0:
    print("[Mudabbir] Update complete. Starting Mudabbir...")
else:
    print("[Mudabbir] Auto-update failed in handoff mode. Starting current version...")

env = os.environ.copy()
env["MUDABBIR_AUTO_UPDATE_DONE"] = "1"
subprocess.run([sys.executable, "-m", "Mudabbir", *args], env=env)
"""
    subprocess.Popen([sys.executable, "-c", helper, *argv[1:]], close_fds=True)


async def run_telegram_mode(settings: Settings) -> None:
    """Run in Telegram bot mode."""
    from Mudabbir.bot_gateway import run_bot
    from Mudabbir.web_server import find_available_port, run_pairing_server

    # Check if we need to run pairing flow
    if not settings.telegram_bot_token or not settings.allowed_user_id:
        logger.info("🔧 First-time setup: Starting pairing server...")

        # Find available port before showing instructions
        try:
            port = find_available_port(settings.web_port)
        except OSError:
            logger.error(
                "❌ Could not find an available port. Please close other applications and try again."
            )
            return

        print("\n" + "=" * 50)
        print("Mudabbir SETUP")
        print("=" * 50)
        print("\n1. Create a Telegram bot via @BotFather")
        print("2. Copy the bot token")
        print(f"3. Open http://localhost:{port} in your browser")
        print("4. Paste the token and scan the QR code\n")

        # Open browser automatically with correct port
        webbrowser.open(f"http://localhost:{port}")

        # Run pairing server (blocks until pairing complete)
        await run_pairing_server(settings)

        # Reload settings after pairing
        settings = get_settings(force_reload=True)

    # Start the bot
    logger.info("🚀 Starting Mudabbir (Beta)...")
    await run_bot(settings)


async def run_multi_channel_mode(settings: Settings, args: argparse.Namespace) -> None:
    """Run one or more channel adapters sharing a single bus and AgentLoop."""
    from Mudabbir.agents.loop import AgentLoop
    from Mudabbir.bus import get_message_bus

    bus = get_message_bus()
    adapters = []

    if args.discord:
        if not settings.discord_bot_token:
            logger.error("Discord bot token not configured. Set Mudabbir_DISCORD_BOT_TOKEN.")
        else:
            from Mudabbir.bus.adapters.discord_adapter import DiscordAdapter

            adapters.append(
                DiscordAdapter(
                    token=settings.discord_bot_token,
                    allowed_guild_ids=settings.discord_allowed_guild_ids,
                    allowed_user_ids=settings.discord_allowed_user_ids,
                )
            )

    if args.slack:
        if not settings.slack_bot_token or not settings.slack_app_token:
            logger.error(
                "Slack tokens not configured. Set Mudabbir_SLACK_BOT_TOKEN "
                "and Mudabbir_SLACK_APP_TOKEN."
            )
        else:
            from Mudabbir.bus.adapters.slack_adapter import SlackAdapter

            adapters.append(
                SlackAdapter(
                    bot_token=settings.slack_bot_token,
                    app_token=settings.slack_app_token,
                    allowed_channel_ids=settings.slack_allowed_channel_ids,
                )
            )

    if args.whatsapp:
        if not settings.whatsapp_access_token or not settings.whatsapp_phone_number_id:
            logger.error(
                "WhatsApp not configured. Set Mudabbir_WHATSAPP_ACCESS_TOKEN "
                "and Mudabbir_WHATSAPP_PHONE_NUMBER_ID."
            )
        else:
            from Mudabbir.bus.adapters.whatsapp_adapter import WhatsAppAdapter

            adapters.append(
                WhatsAppAdapter(
                    access_token=settings.whatsapp_access_token,
                    phone_number_id=settings.whatsapp_phone_number_id,
                    verify_token=settings.whatsapp_verify_token or "",
                    allowed_phone_numbers=settings.whatsapp_allowed_phone_numbers,
                )
            )

    if getattr(args, "signal", False):
        if not settings.signal_phone_number:
            logger.error("Signal not configured. Set Mudabbir_SIGNAL_PHONE_NUMBER.")
        else:
            from Mudabbir.bus.adapters.signal_adapter import SignalAdapter

            adapters.append(
                SignalAdapter(
                    api_url=settings.signal_api_url,
                    phone_number=settings.signal_phone_number,
                    allowed_phone_numbers=settings.signal_allowed_phone_numbers,
                )
            )

    if getattr(args, "matrix", False):
        if not settings.matrix_homeserver or not settings.matrix_user_id:
            logger.error(
                "Matrix not configured. Set Mudabbir_MATRIX_HOMESERVER "
                "and Mudabbir_MATRIX_USER_ID."
            )
        else:
            from Mudabbir.bus.adapters.matrix_adapter import MatrixAdapter

            adapters.append(
                MatrixAdapter(
                    homeserver=settings.matrix_homeserver,
                    user_id=settings.matrix_user_id,
                    access_token=settings.matrix_access_token,
                    password=settings.matrix_password,
                    allowed_room_ids=settings.matrix_allowed_room_ids,
                    device_id=settings.matrix_device_id,
                )
            )

    if getattr(args, "teams", False):
        if not settings.teams_app_id or not settings.teams_app_password:
            logger.error(
                "Teams not configured. Set Mudabbir_TEAMS_APP_ID and Mudabbir_TEAMS_APP_PASSWORD."
            )
        else:
            from Mudabbir.bus.adapters.teams_adapter import TeamsAdapter

            adapters.append(
                TeamsAdapter(
                    app_id=settings.teams_app_id,
                    app_password=settings.teams_app_password,
                    allowed_tenant_ids=settings.teams_allowed_tenant_ids,
                    webhook_port=settings.teams_webhook_port,
                )
            )

    if getattr(args, "gchat", False):
        if not settings.gchat_service_account_key:
            logger.error("Google Chat not configured. Set Mudabbir_GCHAT_SERVICE_ACCOUNT_KEY.")
        else:
            from Mudabbir.bus.adapters.gchat_adapter import GoogleChatAdapter

            adapters.append(
                GoogleChatAdapter(
                    mode=settings.gchat_mode,
                    service_account_key=settings.gchat_service_account_key,
                    project_id=settings.gchat_project_id,
                    subscription_id=settings.gchat_subscription_id,
                    allowed_space_ids=settings.gchat_allowed_space_ids,
                )
            )

    if not adapters:
        logger.error("No channel adapters could be started. Check your configuration.")
        return

    agent_loop = AgentLoop()

    for adapter in adapters:
        await adapter.start(bus)
        logger.info(f"Started {adapter.channel.value} adapter")

    loop_task = asyncio.create_task(agent_loop.start())

    # If WhatsApp is one of the adapters, start a minimal webhook server
    whatsapp_server = None
    if args.whatsapp:
        import uvicorn

        import Mudabbir.whatsapp_gateway as wa_gw
        from Mudabbir.whatsapp_gateway import create_whatsapp_app

        # Point the gateway module at our adapter
        for a in adapters:
            if a.channel.value == "whatsapp":
                wa_gw._whatsapp_adapter = a
                break

        wa_app = create_whatsapp_app(settings)
        config = uvicorn.Config(
            wa_app, host=settings.web_host, port=settings.web_port, log_level="info"
        )
        whatsapp_server = uvicorn.Server(config)
        asyncio.create_task(whatsapp_server.serve())

    try:
        await loop_task
    except asyncio.CancelledError:
        logger.info("Stopping channels...")
    finally:
        await agent_loop.stop()
        for adapter in adapters:
            await adapter.stop()


def _is_headless() -> bool:
    """Detect headless server (no display)."""
    import os

    if sys.platform == "darwin":
        return False  # macOS always has a display
    return not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY")


def run_dashboard_mode(settings: Settings, host: str, port: int, dev: bool = False) -> None:
    """Run in web dashboard mode."""
    from Mudabbir.dashboard import run_dashboard

    run_dashboard(host=host, port=port, open_browser=not _is_headless() and not dev, dev=dev)


async def check_ollama(settings: Settings) -> int:
    """Check Ollama connectivity, model availability, and tool calling support.

    Returns 0 on success, 1 on failure.
    """
    import httpx
    from rich.console import Console

    from Mudabbir.llm.client import resolve_llm_client

    console = Console()
    llm = resolve_llm_client(settings, force_provider="ollama")
    ollama_host = llm.ollama_host
    ollama_model = llm.model
    failures = 0

    # 1. Check server connectivity
    console.print(f"\n  Checking Ollama at [bold]{ollama_host}[/] ...")
    try:
        async with httpx.AsyncClient(timeout=10.0) as client:
            resp = await client.get(f"{ollama_host}/api/tags")
            resp.raise_for_status()
            tags_data = resp.json()
        models = [m.get("name", "") for m in tags_data.get("models", [])]
        console.print(f"  [green]\\[OK][/]  Server reachable — {len(models)} model(s) available")
    except Exception as e:
        console.print(f"  [red]\\[FAIL][/] Cannot reach Ollama server: {e}")
        console.print("         Make sure Ollama is running: [bold]ollama serve[/]")
        return 1

    # 2. Check configured model is available
    # Ollama model names may or may not include ":latest" tag
    model_found = any(m == ollama_model or m.startswith(f"{ollama_model}:") for m in models)
    if model_found:
        console.print(f"  [green]\\[OK][/]  Model '{ollama_model}' is available")
    else:
        console.print(f"  [yellow]\\[WARN][/] Model '{ollama_model}' not found locally")
        if models:
            console.print(f"         Available: {', '.join(models[:10])}")
        console.print(f"         Pull it with: [bold]ollama pull {ollama_model}[/]")
        failures += 1

    # 3. Test Anthropic-compatible endpoint (basic completion)
    console.print("  Testing Anthropic Messages API compatibility ...")
    try:
        ac = llm.create_anthropic_client(timeout=60.0, max_retries=1)
        response = await ac.messages.create(
            model=ollama_model,
            max_tokens=32,
            messages=[{"role": "user", "content": "Say hi"}],
        )
        text = response.content[0].text if response.content else ""
        console.print(f"  [green]\\[OK][/]  Messages API works — response: {text[:60]}")
    except Exception as e:
        console.print(f"  [red]\\[FAIL][/] Messages API failed: {e}")
        console.print("         Ollama v0.14.0+ is required for Anthropic API compatibility")
        failures += 1
        # Skip tool test if basic API fails
        console.print(f"\n  Result: {2 - (1 if model_found else 0)}/3 checks passed")
        return 1

    # 4. Test tool calling
    console.print("  Testing tool calling support ...")
    try:
        tool_response = await ac.messages.create(
            model=ollama_model,
            max_tokens=256,
            messages=[{"role": "user", "content": "What is 2 + 2?"}],
            tools=[
                {
                    "name": "calculator",
                    "description": "Performs arithmetic calculations",
                    "input_schema": {
                        "type": "object",
                        "properties": {
                            "expression": {
                                "type": "string",
                                "description": "Math expression to evaluate",
                            }
                        },
                        "required": ["expression"],
                    },
                }
            ],
        )
        has_tool_use = any(b.type == "tool_use" for b in tool_response.content)
        if has_tool_use:
            console.print("  [green]\\[OK][/]  Tool calling works")
        else:
            console.print("  [yellow]\\[WARN][/] Model responded but did not use the tool")
            console.print("         Tool calling quality varies by model. Try a larger model.")
            failures += 1
    except Exception as e:
        console.print(f"  [yellow]\\[WARN][/] Tool calling test failed: {e}")
        console.print("         Some models may not support tool calling reliably.")
        failures += 1

    passed = 4 - failures
    console.print(f"\n  Result: [bold]{passed}/4[/] checks passed")
    if failures == 0:
        console.print("  [green]Ollama is ready to use with Mudabbir![/]")
        console.print(
            "  Set [bold]llm_provider=ollama[/] in settings"
            " or [bold]Mudabbir_LLM_PROVIDER=ollama[/]\n"
        )
    return 1 if failures > 1 else 0


async def check_openai_compatible(settings: Settings) -> int:
    """Check OpenAI-compatible endpoint connectivity and tool calling support.

    Returns 0 on success, 1 on failure.
    """
    from rich.console import Console

    from Mudabbir.llm.client import resolve_llm_client

    console = Console()
    llm = resolve_llm_client(settings, force_provider="openai_compatible")
    base_url = llm.openai_compatible_base_url
    model = llm.model

    if not base_url:
        console.print("\n  [red]\\[FAIL][/] No base URL configured.")
        console.print(
            "         Set [bold]Mudabbir_OPENAI_COMPATIBLE_BASE_URL[/] or configure in Settings.\n"
        )
        return 1

    if not model:
        console.print("\n  [red]\\[FAIL][/] No model configured.")
        console.print(
            "         Set [bold]Mudabbir_OPENAI_COMPATIBLE_MODEL[/] or configure in Settings.\n"
        )
        return 1

    failures = 0

    # 1. Test OpenAI Chat Completions API
    console.print(f"\n  Checking endpoint at [bold]{base_url}[/] ...")
    console.print(f"  Model: [bold]{model}[/]")
    console.print("  Testing Chat Completions API ...")
    try:
        oc = llm.create_openai_client(timeout=60.0, max_retries=1)
        response = await oc.chat.completions.create(
            model=model,
            max_tokens=32,
            messages=[{"role": "user", "content": "Say hi"}],
        )
        text = response.choices[0].message.content or ""
        console.print(f"  [green]\\[OK][/]  Chat Completions API works — response: {text[:60]}")
    except Exception as e:
        console.print(f"  [red]\\[FAIL][/] Chat Completions API failed: {e}")
        console.print("\n  Result: 0/2 checks passed")
        return 1

    # 2. Test tool calling
    console.print("  Testing tool calling support ...")
    try:
        tool_response = await oc.chat.completions.create(
            model=model,
            max_tokens=256,
            messages=[{"role": "user", "content": "What is 2 + 2?"}],
            tools=[
                {
                    "type": "function",
                    "function": {
                        "name": "calculator",
                        "description": "Performs arithmetic calculations",
                        "parameters": {
                            "type": "object",
                            "properties": {
                                "expression": {
                                    "type": "string",
                                    "description": "Math expression to evaluate",
                                }
                            },
                            "required": ["expression"],
                        },
                    },
                }
            ],
        )
        has_tool_use = bool(tool_response.choices[0].message.tool_calls)
        if has_tool_use:
            console.print("  [green]\\[OK][/]  Tool calling works")
        else:
            console.print("  [yellow]\\[WARN][/] Model responded but did not use the tool")
            console.print("         Tool calling quality varies by model.")
            failures += 1
    except Exception as e:
        console.print(f"  [yellow]\\[WARN][/] Tool calling test failed: {e}")
        failures += 1

    passed = 2 - failures
    console.print(f"\n  Result: [bold]{passed}/2[/] checks passed")
    if failures == 0:
        console.print("  [green]Endpoint is ready to use with Mudabbir![/]")
        console.print(
            "  Set [bold]llm_provider=openai_compatible[/] in settings"
            " or [bold]Mudabbir_LLM_PROVIDER=openai_compatible[/]\n"
        )
    return 1 if failures > 1 else 0


async def run_doctor(settings: Settings) -> int:
    """Run diagnostic checks and print categorized output."""
    from collections import defaultdict

    from Mudabbir.config import get_config_dir
    from Mudabbir.health import get_health_engine
    from Mudabbir.update_check import check_for_updates

    engine = get_health_engine()
    startup = engine.run_startup_checks()
    connectivity = await engine.run_connectivity_checks()
    all_results = startup + connectivity

    grouped: dict[str, list] = defaultdict(list)
    for result in all_results:
        grouped[result.category].append(result)

    update_info = check_for_updates(_runtime_version(), get_config_dir())
    update_rows: list[tuple[str, str, str]] = []
    if update_info and update_info.get("update_available"):
        latest = update_info.get("latest", "unknown")
        update_rows.append(
            (
                "warning",
                f"Update available: {latest}",
                "Run: pip install --upgrade Mudabbir",
            )
        )
    else:
        update_rows.append(("ok", "You are on the latest known version", ""))

    def _icon(status: str) -> str:
        return {"ok": "[OK]", "warning": "[WARN]", "critical": "[FAIL]"}.get(status, "[?]")

    print("\n" + "=" * 64)
    print("Mudabbir Doctor")
    print("=" * 64)

    sections = (
        ("Config", grouped.get("config", [])),
        ("Storage", grouped.get("storage", [])),
        ("Connectivity", grouped.get("connectivity", [])),
    )
    for title, rows in sections:
        print(f"\n{title}:")
        if not rows:
            print("  [OK] No checks in this category")
            continue
        for row in rows:
            print(f"  {_icon(row.status)} {row.name}: {row.message}")
            if row.fix_hint and row.status != "ok":
                print(f"       Fix: {row.fix_hint}")

    print("\nUpdates:")
    for status, message, hint in update_rows:
        print(f"  {_icon(status)} {message}")
        if hint:
            print(f"       Fix: {hint}")

    status = engine.overall_status
    print(f"\nOverall: {status.upper()}")
    print("=" * 64 + "\n")

    has_critical = any(r.status == "critical" for r in all_results)
    return 1 if has_critical else 0


def _check_extras_installed(args: argparse.Namespace) -> None:
    """Check that required optional dependencies are installed for the chosen mode.

    Exits with a helpful message if something is missing.
    """
    missing: list[tuple[str, str, str]] = []  # (package, import_name, extra)

    # Dashboard deps are now in core — no need to check for them.

    if args.telegram:
        if importlib.util.find_spec("telegram") is None:
            missing.append(("python-telegram-bot", "telegram", "telegram"))

    channel_checks = {
        "discord": ("discord.py", "discord", "discord"),
        "slack": ("slack-bolt", "slack_bolt", "slack"),
    }
    for flag, (pkg, mod, extra) in channel_checks.items():
        if getattr(args, flag, False) and importlib.util.find_spec(mod) is None:
            missing.append((pkg, mod, extra))

    if not missing:
        return

    print("\n  Missing dependencies detected:\n")
    extras = set()
    for pkg, _mod, extra in missing:
        print(f"    - {pkg}  (extra: {extra})")
        extras.add(extra)
    extras_str = ",".join(sorted(extras))
    print(f"\n  Install with:  pip install 'Mudabbir[{extras_str}]'\n")
    sys.exit(1)


def main() -> None:
    """Main entry point."""
    if _should_auto_update_on_boot(sys.argv):
        _bootstrap_latest_and_restart(sys.argv)

    parser = argparse.ArgumentParser(
        description="Mudabbir (Beta) - The AI agent that runs on your laptop",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Mudabbir                          Start web dashboard (default)
  Mudabbir --telegram               Start in Telegram-only mode
  Mudabbir --discord                Start headless Discord bot
  Mudabbir --slack                  Start headless Slack bot (Socket Mode)
  Mudabbir --whatsapp               Start headless WhatsApp webhook server
  Mudabbir --discord --slack        Run Discord + Slack simultaneously
  Mudabbir --doctor                 Run categorized diagnostics and exit
  Mudabbir --dev                    Start dashboard with auto-reload (dev mode)
""",
    )

    parser.add_argument(
        "--web",
        "-w",
        action="store_true",
        help="Run web dashboard (same as default, kept for compatibility)",
    )
    parser.add_argument(
        "--telegram",
        action="store_true",
        help="Run Telegram-only mode (legacy pairing flow)",
    )
    parser.add_argument("--discord", action="store_true", help="Run headless Discord bot")
    parser.add_argument("--slack", action="store_true", help="Run headless Slack bot (Socket Mode)")
    parser.add_argument(
        "--whatsapp", action="store_true", help="Run headless WhatsApp webhook server"
    )
    parser.add_argument("--signal", action="store_true", help="Run headless Signal bot")
    parser.add_argument("--matrix", action="store_true", help="Run headless Matrix bot")
    parser.add_argument("--teams", action="store_true", help="Run headless Teams bot")
    parser.add_argument("--gchat", action="store_true", help="Run headless Google Chat bot")
    parser.add_argument(
        "--security-audit",
        action="store_true",
        help="Run security audit and print report",
    )
    parser.add_argument(
        "--fix",
        action="store_true",
        help="Auto-fix fixable issues found by --security-audit",
    )
    parser.add_argument(
        "--host",
        type=str,
        default=None,
        help="Host to bind web server (default: auto-detect; 0.0.0.0 on headless servers)",
    )
    parser.add_argument(
        "--port", "-p", type=int, default=8888, help="Port for web server (default: 8888)"
    )
    parser.add_argument(
        "--dev", action="store_true", help="Development mode with auto-reload"
    )
    parser.add_argument(
        "--check-ollama",
        action="store_true",
        help="Check Ollama connectivity, model availability, and tool calling support",
    )
    parser.add_argument(
        "--check-openai-compatible",
        action="store_true",
        help="Check OpenAI-compatible endpoint connectivity and tool calling support",
    )
    parser.add_argument(
        "--doctor",
        action="store_true",
        help="Run categorized health diagnostics (Config, Storage, Connectivity, Updates)",
    )
    parser.add_argument(
        "--version", "-v", action="version", version=f"%(prog)s {_runtime_version()}"
    )
    parser.add_argument(
        "--stop",
        action="store_true",
        help="Force-stop running Mudabbir instance(s) (useful when CTRL+C is ignored)",
    )

    args = parser.parse_args()

    if args.stop:
        count = _stop_running_instances()
        if count > 0:
            print(f"Stopped {count} Mudabbir process(es).")
        else:
            print("No running Mudabbir process found.")
        raise SystemExit(0)

    # Fail fast if optional deps are missing for the chosen mode
    _check_extras_installed(args)

    settings = get_settings()

    if args.doctor:
        exit_code = asyncio.run(run_doctor(settings))
        raise SystemExit(exit_code)

    # Strict fail-fast for invalid Mem0 settings (do not crash later during chat).
    from Mudabbir.memory.validation import validate_mem0_settings

    mem0_errors = validate_mem0_settings(settings)
    if mem0_errors:
        _print_mem0_config_error(mem0_errors)
        raise SystemExit(2)

    # Run startup health checks (non-blocking, informational only)
    if settings.health_check_on_startup:
        try:
            from Mudabbir.health import get_health_engine

            engine = get_health_engine()
            results = engine.run_startup_checks()
            issues = [r for r in results if r.status != "ok"]
            if issues:
                print()
                for r in results:
                    if r.status == "ok":
                        print(f"  \033[32m[OK]\033[0m   {r.name}: {r.message}")
                    elif r.status == "warning":
                        print(f"  \033[33m[WARN]\033[0m {r.name}: {r.message}")
                        if r.fix_hint:
                            print(f"         {r.fix_hint}")
                    else:
                        print(f"  \033[31m[FAIL]\033[0m {r.name}: {r.message}")
                        if r.fix_hint:
                            print(f"         {r.fix_hint}")
                status = engine.overall_status
                color = {"healthy": "32", "degraded": "33", "unhealthy": "31"}.get(status, "0")
                print(f"\n  System: \033[{color}m{status.upper()}\033[0m\n")
        except Exception:
            pass  # Health engine failure never blocks startup

    # Check for updates (cached daily, silent on error)
    from Mudabbir.config import get_config_dir
    from Mudabbir.update_check import check_for_updates, print_update_notice

    update_info = check_for_updates(_runtime_version(), get_config_dir())
    if update_info and update_info.get("update_available"):
        print_update_notice(update_info)

    # Resolve host: explicit flag > config > auto-detect
    if args.host is not None:
        host = args.host
    elif settings.web_host != "127.0.0.1":
        host = settings.web_host
    elif _is_headless():
        host = "0.0.0.0"
        logger.info("Headless server detected — binding to 0.0.0.0")
    else:
        host = "127.0.0.1"

    has_channel_flag = (
        args.discord
        or args.slack
        or args.whatsapp
        or args.signal
        or args.matrix
        or args.teams
        or args.gchat
    )

    try:
        _write_runtime_pid()
        if args.check_ollama:
            exit_code = asyncio.run(check_ollama(settings))
            raise SystemExit(exit_code)
        elif args.check_openai_compatible:
            exit_code = asyncio.run(check_openai_compatible(settings))
            raise SystemExit(exit_code)
        elif args.security_audit:
            from Mudabbir.security.audit_cli import run_security_audit

            exit_code = asyncio.run(run_security_audit(fix=args.fix))
            raise SystemExit(exit_code)
        elif args.telegram:
            asyncio.run(run_telegram_mode(settings))
        elif has_channel_flag:
            asyncio.run(run_multi_channel_mode(settings, args))
        else:
            # Default: web dashboard (also handles --web flag)
            run_dashboard_mode(settings, host, args.port, dev=args.dev)
    except KeyboardInterrupt:
        logger.info("👋 Mudabbir stopped.")
    finally:
        # Coordinated singleton shutdown
        from Mudabbir.lifecycle import shutdown_all

        try:
            asyncio.run(shutdown_all())
        except RuntimeError:
            # Event loop already closed — best-effort sync cleanup
            pass
        _clear_runtime_pid()


if __name__ == "__main__":
    main()
