# Mudabbir Master Documentation (EN)

Mudabbir is a renamed and customized fork of the original `mudabbir` project.

## 1) Header

- Project name: `Mudabbir`
- Package/runtime identity: `mudabbir` (internal package naming may still appear in architecture/development references)
- Python requirement: 3.11+

## 2) Project Overview

Mudabbir is a self-hosted AI agent that runs locally and can be controlled through Telegram, Discord, Slack, WhatsApp, or a web dashboard.

### What changed

- Rebrand from `mudabbir` to `Mudabbir`.
- Runtime compatibility layer so `Mudabbir` imports work in current Windows install.
- Improved startup reliability:
  - safer secret-store handling and migration behavior,
  - better startup defaults for WhatsApp personal mode,
  - launcher stability improvements.
- Enhanced desktop/direct-intent behavior in Open Interpreter flow.

## 3) Quick Start

### Main project path

- `C:\Users\Admin\Documents\Mudabbir`

### Start commands

- `Mudabbir --port 8888`
- `python -m Mudabbir --port 8888`

### Web dashboard

- Open: `http://127.0.0.1:8888`

### Local-only files (not for upload)

- `C:\Users\Admin\Documents\for Mudabbir\local-only`
- This location contains local launcher, ops scripts, tests, and backup/release helpers.

## 4) Project Status

### Canonical source repo

- `C:\Users\Admin\Documents\Mudabbir`

### Mirrored working copies

- `C:\Python314\Lib\site-packages\Mudabbir\Mudabbir-main`

### Run commands

- `Mudabbir`
- `Mudabbir --port 8888`
- `python -m Mudabbir --port 8888`

### Branding state

- Legacy Pocket/Paw branding strings removed from text files.
- Legacy paw/pocket asset filenames renamed to Mudabbir names.

### Git state

- Branch: `main`
- Latest commits include complete rebrand and cleanup.
- Bundle export availability is environment-dependent and tied to local-only tooling.

## 5) Development & Commands

### Core commands

```bash
# Install dev dependencies
uv sync --dev

# Run the app (web dashboard is the default — auto-starts all configured adapters)
uv run mudabbir

# Run Telegram-only mode (legacy pairing flow)
uv run mudabbir --telegram

# Run headless Discord bot
uv run mudabbir --discord

# Run headless Slack bot (Socket Mode, no public URL needed)
uv run mudabbir --slack

# Run headless WhatsApp webhook server
uv run mudabbir --whatsapp

# Run multiple headless channels simultaneously
uv run mudabbir --discord --slack

# Run in development mode (auto-reload on file changes)
uv run mudabbir --dev

# Run all tests
uv run pytest

# Run a single test file
uv run pytest tests/test_bus.py

# Run a specific test
uv run pytest tests/test_bus.py::test_publish_subscribe -v

# Run tests (skip e2e, they need Playwright browsers)
uv run pytest --ignore=tests/e2e

# Run a specific test file (alternative)
uv run pytest tests/test_bus.py -v

# Lint
uv run ruff check .

# Format
uv run ruff format .

# Type check
uv run mypy .

# Build package
python -m build
```

## 6) Architecture

### Message Bus Pattern

The core architecture is an event-driven message bus (`src/mudabbir/bus/`). All communication flows through three event types defined in `bus/events.py`:

- **InboundMessage** — user input from any channel (Telegram, WebSocket, CLI)
- **OutboundMessage** — agent responses back to channels (supports streaming via `is_stream_chunk`/`is_stream_end`)
- **SystemEvent** — internal events (tool_start, tool_result, thinking, error) consumed by the web dashboard Activity panel

### AgentLoop -> AgentRouter -> Backend

The processing pipeline lives in `agents/loop.py` and `agents/router.py`:

1. **AgentLoop** consumes from the message bus, manages memory context, and streams responses back.
2. **AgentRouter** selects and delegates to one of three backends based on `settings.agent_backend`:
   - `claude_agent_sdk` (default/recommended) — Official Claude Agent SDK with built-in tools (Bash, Read, Write, etc.). Uses `PreToolUse` hooks for dangerous command blocking. Lives in `agents/claude_sdk.py`. SDK-specific settings: `claude_sdk_model` (empty = let Claude Code auto-select), `claude_sdk_max_turns` (default 25). Smart model routing is disabled by default to avoid conflicting with Claude Code's own routing.
   - `mudabbir_native` — Custom orchestrator: Anthropic SDK for reasoning + Open Interpreter for execution. Lives in `agents/mudabbir_native.py`.
   - `open_interpreter` — Standalone Open Interpreter supporting Ollama/OpenAI/Anthropic. Lives in `agents/open_interpreter.py`.
3. All backends yield standardized dicts with `type` (`message/tool_use/tool_result/error/done`), `content`, and `metadata`.

### Channel Adapters

`bus/adapters/` contains protocol translators that bridge external channels to the message bus:

- `TelegramAdapter` — python-telegram-bot
- `WebSocketAdapter` — FastAPI WebSockets
- `DiscordAdapter` — discord.py (optional dep `mudabbir[discord]`). Slash command `/mudabbir` + DM/mention support. Stream buffering with edit-in-place (1.5s rate limit).
- `SlackAdapter` — slack-bolt Socket Mode (optional dep `mudabbir[slack]`). Handles `app_mention` + DM events. No public URL needed. Thread support via `thread_ts` metadata.
- `WhatsAppAdapter` — WhatsApp Business Cloud API via `httpx` (core dep). No streaming; accumulates chunks and sends on `stream_end`. Dashboard exposes `/webhook/whatsapp` routes; standalone mode runs its own FastAPI server.

Dashboard channel management:

- The web dashboard (default mode) auto-starts all configured adapters on startup.
- Channels can be configured, started, and stopped dynamically from the Channels modal in the sidebar.
- REST API:
  - `GET /api/channels/status`
  - `POST /api/channels/save`
  - `POST /api/channels/toggle`

### Key Subsystems

- **Memory** (`memory/`) — Session history + long-term facts, file-based storage in `~/.mudabbir/memory/`. Protocol-based (`MemoryStoreProtocol`) for future backend swaps.
- **Browser** (`browser/`) — Playwright-based automation using accessibility tree snapshots (not screenshots). `BrowserDriver` returns `NavigationResult` with a `refmap` mapping ref numbers to CSS selectors.
- **Security** (`security/`) — Guardian AI (secondary LLM safety check) + append-only audit log (`~/.mudabbir/audit.jsonl`).
- **Tools** (`tools/`) — `ToolProtocol` with `ToolDefinition` supporting both Anthropic and OpenAI schema export. Built-in tools in `tools/builtin/`.
- **Bootstrap** (`bootstrap/`) — `AgentContextBuilder` assembles the system prompt from identity, memory, and current state.
- **Config** (`config.py`) — Pydantic Settings with `MUDABBIR_` env prefix, JSON config at `~/.mudabbir/config.json`. Channel-specific config includes:
  - `discord_bot_token`
  - `discord_allowed_guild_ids`
  - `discord_allowed_user_ids`
  - `slack_bot_token`
  - `slack_app_token`
  - `slack_allowed_channel_ids`
  - `whatsapp_access_token`
  - `whatsapp_phone_number_id`
  - `whatsapp_verify_token`
  - `whatsapp_allowed_phone_numbers`

### Frontend

The web dashboard (`frontend/`) is vanilla JS/CSS/HTML served via FastAPI+Jinja2. No build step. Communicates with the backend over WebSocket for real-time streaming.

### Project structure

```text
src/mudabbir/
  agents/            # Agent backends (Claude SDK, Native, Open Interpreter) + router
  bus/               # Message bus + event types
    adapters/        # Channel adapters (Telegram, Discord, Slack, WhatsApp, etc.)
  tools/
    builtin/         # 60+ built-in tools (Gmail, Spotify, web search, filesystem, etc.)
    protocol.py      # ToolProtocol interface (implement this for new tools)
    registry.py      # Central tool registry with policy filtering
    policy.py        # Tool access control (profiles, allow/deny lists)
  memory/            # Memory stores (file-based, mem0)
  security/          # Guardian AI, injection scanner, audit log
  mcp/               # MCP server configuration and management
  deep_work/         # Multi-step task decomposition and execution
  mission_control/   # Multi-agent orchestration
  daemon/            # Background tasks, triggers, proactive behaviors
  config.py          # Pydantic Settings with MUDABBIR_ env prefix
  credentials.py     # Fernet-encrypted credential store
  dashboard.py       # FastAPI server, WebSocket handler, REST APIs
  scheduler.py       # APScheduler-based reminders and cron jobs
frontend/            # Vanilla JS/CSS/HTML dashboard (no build step)
tests/               # pytest suite (130+ tests)
```

### Key conventions

- **Async everywhere**: All agent, bus, memory, and tool interfaces are async. Tests use `pytest-asyncio` with `asyncio_mode = "auto"`.
- **Protocol-oriented**: Core interfaces (`AgentProtocol`, `ToolProtocol`, `MemoryStoreProtocol`, `BaseChannelAdapter`) are Python `Protocol` classes for swappable implementations.
- **Env vars**: All settings use `MUDABBIR_` prefix (e.g., `MUDABBIR_ANTHROPIC_API_KEY`).
- **Ruff config**: line-length 100, target Python 3.11, lint rules E/F/I/UP.
- **Entry point**: `mudabbir.__main__:main`.
- **Lazy imports**: Agent backends are imported inside `AgentRouter._initialize_agent()` to avoid loading unused dependencies.

## 7) Contributing

Mudabbir is an open-source AI agent. We welcome all contribution types: bug fixes, new tools, channel adapters, docs, and tests.

### Branch strategy

> **All pull requests must target the `dev` branch.**
>
> PRs opened against `main` will be closed. The `main` branch is updated only via merge from `dev` when a release is ready.

### Before you start

- Search existing issues: <https://github.com/Ahmed5754/Mudabbir/issues>
- Check open pull requests: <https://github.com/Ahmed5754/Mudabbir/pulls>
- If an issue exists, comment that you are taking it.
- If no issue exists, open one first and discuss approach.
- Good starting label: <https://github.com/Ahmed5754/Mudabbir/labels/good%20first%20issue>

### Setup

1. Fork and clone your fork.
2. Create feature branch from `dev`:
   ```bash
   git checkout dev
   git pull origin dev
   git checkout -b feat/your-feature
   ```
3. Install dependencies:
   ```bash
   uv sync --dev
   ```
4. Verify setup:
   ```bash
   uv run mudabbir
   ```
   Dashboard should open at `http://localhost:8888`.

### Writing code

#### Conventions

- Async everywhere.
- Protocol-oriented design.
- Ruff config: line-length 100, target Python 3.11, rules E/F/I/UP.
- Lazy imports for optional/heavy dependencies.

#### Adding a new tool

1. Create file in `src/mudabbir/tools/builtin/`.
2. Subclass `BaseTool` from `tools/protocol.py`.
3. Implement `name`, `description`, `parameters` (JSON Schema), and `execute(**params) -> str`.
4. Add class to `tools/builtin/__init__.py` lazy imports.
5. Add tool to policy group in `tools/policy.py`.
6. Write tests.

#### Adding a new channel adapter

1. Create file in `src/mudabbir/bus/adapters/`.
2. Extend `BaseChannelAdapter`.
3. Implement `_on_start()`, `_on_stop()`, `send(message)`.
4. Use `self._publish_inbound()` for incoming messages.
5. Add optional dependency extras in `pyproject.toml`.

### Security checklist

- Never expose/log credentials.
- Add new secret config fields to `SECRET_FIELDS` in `credentials.py`.
- Shell-executing tools must respect Guardian AI safety checks.
- New API endpoints require auth middleware.
- Add injection tests for user-input features.

### Commit messages

Use Conventional Commits:

```text
feat: add Spotify playback tool
fix: handle empty WebSocket message
docs: update channel adapter guide
refactor: simplify model router thresholds
test: add coverage for injection scanner
```

- Keep subject under 72 chars.
- Add body when needed.

### Pull request checklist

- [ ] Branch based on `dev`
- [ ] PR targets `dev`
- [ ] Tests pass (`uv run pytest --ignore=tests/e2e`)
- [ ] Lint passes (`uv run ruff check .`)
- [ ] No secrets in diff
- [ ] New config fields added to `Settings.save()` dict
- [ ] New secret fields added to `SECRET_FIELDS`
- [ ] New tools registered in proper policy group
- [ ] New optional deps added to `pyproject.toml` extras

### Code review

- Maintainers review PRs; usual response is within days.
- Small, focused PRs are reviewed faster.
- If no response for a week, ping in related issue.

### Reporting bugs

Include:

- expected behavior
- actual behavior
- reproduction steps
- OS, Python version, Mudabbir version (`mudabbir --version`)

### Questions

- Open a discussion: <https://github.com/Ahmed5754/Mudabbir/discussions>
- or comment on a relevant issue.

## 8) License (MIT)

MIT License

Copyright (c) 2026 Mudabbir Team

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

