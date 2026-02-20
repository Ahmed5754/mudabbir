"""Capability schema for desktop/system command planning."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


RiskLevel = str  # safe | elevated | destructive


@dataclass(frozen=True)
class CapabilitySpec:
    """Describes one executable capability exposed to the planner."""

    id: str
    aliases_ar: tuple[str, ...] = ()
    aliases_en: tuple[str, ...] = ()
    risk_level: RiskLevel = "safe"
    required_privileges: tuple[str, ...] = ()
    executor: str = "desktop"
    result_schema: dict[str, Any] = field(default_factory=dict)
    stage: str = "A"

