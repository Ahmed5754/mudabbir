# Health engine package — self-healing diagnostics for Mudabbir.
# Created: 2026-02-17
# Exports get_health_engine() singleton factory.

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from Mudabbir.health.engine import HealthEngine

_instance: HealthEngine | None = None


def get_health_engine() -> HealthEngine:
    """Get or create the singleton HealthEngine."""
    global _instance  # noqa: PLW0603
    if _instance is None:
        from Mudabbir.health.engine import HealthEngine

        _instance = HealthEngine()
    return _instance
