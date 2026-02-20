"""Smoke tests for package import and basic metadata."""

from __future__ import annotations

import Mudabbir
import Mudabbir.agents.codex_cli
import Mudabbir.agents.copilot_sdk
import Mudabbir.agents.google_adk
import Mudabbir.agents.openai_agents
import Mudabbir.agents.opencode
import Mudabbir.agents.tool_bridge


def test_package_imports() -> None:
    """Package import should work in CI."""
    assert Mudabbir is not None


def test_package_version_present() -> None:
    """Package should expose a non-empty version string."""
    assert isinstance(Mudabbir.__version__, str)
    assert Mudabbir.__version__.strip() != ""
