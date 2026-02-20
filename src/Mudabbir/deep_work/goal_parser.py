"""Deep Work goal parser.

Converts free-form project descriptions into a structured analysis object
used by the two-step Start Project flow.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

VALID_DOMAINS = frozenset({"code", "business", "creative", "education", "events", "home", "hybrid"})
VALID_COMPLEXITIES = frozenset({"S", "M", "L", "XL"})
VALID_RESEARCH_DEPTHS = frozenset({"none", "quick", "standard", "deep"})


@dataclass
class GoalAnalysis:
    """Structured result for goal parsing."""

    goal: str = ""
    domain: str = "hybrid"
    sub_domains: list[str] = field(default_factory=list)
    complexity: str = "M"
    estimated_phases: int = 3
    ai_capabilities: list[str] = field(default_factory=list)
    human_requirements: list[str] = field(default_factory=list)
    constraints_detected: list[str] = field(default_factory=list)
    clarifications_needed: list[str] = field(default_factory=list)
    suggested_research_depth: str = "standard"
    confidence: float = 0.7

    def to_dict(self) -> dict[str, Any]:
        return {
            "goal": self.goal,
            "domain": self.domain,
            "sub_domains": self.sub_domains,
            "complexity": self.complexity,
            "estimated_phases": self.estimated_phases,
            "ai_capabilities": self.ai_capabilities,
            "human_requirements": self.human_requirements,
            "constraints_detected": self.constraints_detected,
            "clarifications_needed": self.clarifications_needed,
            "suggested_research_depth": self.suggested_research_depth,
            "confidence": self.confidence,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "GoalAnalysis":
        domain = str(data.get("domain", "hybrid")).lower().strip()
        if domain not in VALID_DOMAINS:
            domain = "hybrid"

        complexity = str(data.get("complexity", "M")).upper().strip()
        if complexity not in VALID_COMPLEXITIES:
            complexity = "M"

        research = str(data.get("suggested_research_depth", "standard")).lower().strip()
        if research not in VALID_RESEARCH_DEPTHS:
            research = "standard"

        phases_raw = data.get("estimated_phases", 3)
        try:
            phases = int(phases_raw)
        except (TypeError, ValueError):
            phases = 3
        phases = max(1, min(phases, 10))

        confidence_raw = data.get("confidence", 0.7)
        try:
            confidence = float(confidence_raw)
        except (TypeError, ValueError):
            confidence = 0.7
        confidence = max(0.0, min(confidence, 1.0))

        def _to_list(value: Any) -> list[str]:
            if not isinstance(value, list):
                return []
            return [str(item).strip() for item in value if str(item).strip()]

        return cls(
            goal=str(data.get("goal", "")).strip(),
            domain=domain,
            sub_domains=_to_list(data.get("sub_domains"))[:8],
            complexity=complexity,
            estimated_phases=phases,
            ai_capabilities=_to_list(data.get("ai_capabilities"))[:8],
            human_requirements=_to_list(data.get("human_requirements"))[:8],
            constraints_detected=_to_list(data.get("constraints_detected"))[:8],
            clarifications_needed=_to_list(data.get("clarifications_needed"))[:4],
            suggested_research_depth=research,
            confidence=confidence,
        )


class GoalParser:
    """Heuristic goal parser for Deep Work onboarding."""

    _DOMAIN_KEYWORDS: dict[str, tuple[str, ...]] = {
        "code": (
            "api",
            "backend",
            "frontend",
            "react",
            "fastapi",
            "python",
            "database",
            "deploy",
            "app",
            "mobile",
            "web",
            "sdk",
        ),
        "business": (
            "marketing",
            "sales",
            "strategy",
            "kpi",
            "revenue",
            "pricing",
            "funnel",
            "brand",
            "customer",
            "operations",
        ),
        "creative": (
            "design",
            "video",
            "content",
            "story",
            "creative",
            "script",
            "copy",
            "campaign",
            "visual",
        ),
        "education": (
            "course",
            "learn",
            "training",
            "curriculum",
            "lesson",
            "teacher",
            "student",
            "exam",
        ),
        "events": (
            "event",
            "conference",
            "meetup",
            "wedding",
            "schedule",
            "venue",
            "attendees",
            "registration",
        ),
        "home": (
            "home",
            "house",
            "apartment",
            "garden",
            "kitchen",
            "renovation",
            "furniture",
            "family",
        ),
    }

    _SUBDOMAIN_KEYWORDS: dict[str, tuple[str, ...]] = {
        "react": ("react", "next.js", "nextjs"),
        "fastapi": ("fastapi", "api"),
        "ai-agent": ("agent", "llm", "ai", "model", "rag"),
        "mobile-app": ("ios", "android", "mobile", "flutter", "react native"),
        "ecommerce": ("ecommerce", "shop", "store", "checkout"),
        "content-marketing": ("content", "blog", "seo"),
        "automation": ("automate", "automation", "workflow"),
        "analytics": ("analytics", "dashboard", "reporting", "kpi"),
    }

    _TIMELINE_RE = re.compile(r"\b(\d+)\s*(day|days|week|weeks|month|months|hour|hours)\b")
    _BUDGET_RE = re.compile(r"(\$|usd|eur|aed|sar)\s*\d+|\bbudget\b")

    async def parse(self, user_input: str) -> GoalAnalysis:
        """Parse a user goal string into a structured analysis."""
        text = (user_input or "").strip()
        lowered = text.lower()

        domain, confidence = self._detect_domain(lowered)
        sub_domains = self._detect_sub_domains(lowered)
        complexity = self._estimate_complexity(lowered)
        estimated_phases = {"S": 2, "M": 3, "L": 5, "XL": 7}.get(complexity, 3)
        constraints = self._detect_constraints(lowered)
        clarifications = self._build_clarifications(lowered, domain, complexity)

        research = self._suggest_research_depth(lowered, domain, complexity)
        ai_caps, human_reqs = self._split_ai_human(domain)

        return GoalAnalysis(
            goal=text[:240] if text else "",
            domain=domain,
            sub_domains=sub_domains,
            complexity=complexity,
            estimated_phases=estimated_phases,
            ai_capabilities=ai_caps,
            human_requirements=human_reqs,
            constraints_detected=constraints,
            clarifications_needed=clarifications,
            suggested_research_depth=research,
            confidence=confidence,
        )

    def _detect_domain(self, lowered: str) -> tuple[str, float]:
        scores: dict[str, int] = {}
        for domain, keywords in self._DOMAIN_KEYWORDS.items():
            scores[domain] = sum(1 for kw in keywords if kw in lowered)

        ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
        best_domain, best_score = ranked[0]
        second_score = ranked[1][1] if len(ranked) > 1 else 0

        if best_score == 0:
            return "hybrid", 0.5
        if best_score - second_score <= 1 and second_score > 0:
            return "hybrid", 0.65

        confidence = 0.7 + min(0.25, 0.03 * best_score)
        return best_domain, min(confidence, 0.95)

    def _detect_sub_domains(self, lowered: str) -> list[str]:
        matches: list[str] = []
        for name, keywords in self._SUBDOMAIN_KEYWORDS.items():
            if any(keyword in lowered for keyword in keywords):
                matches.append(name)
        return matches[:6]

    def _estimate_complexity(self, lowered: str) -> str:
        size_score = len(lowered)
        feature_score = sum(
            1
            for token in (
                " and ",
                " integrate ",
                " integration ",
                "multi",
                "dashboard",
                "authentication",
                "payment",
                "pipeline",
                "automation",
                "deploy",
            )
            if token in lowered
        )

        if any(word in lowered for word in ("enterprise", "multi-tenant", "distributed", "platform")):
            feature_score += 3

        total = size_score / 120 + feature_score
        if total >= 9:
            return "XL"
        if total >= 6:
            return "L"
        if total >= 3:
            return "M"
        return "S"

    def _detect_constraints(self, lowered: str) -> list[str]:
        constraints: list[str] = []
        if self._TIMELINE_RE.search(lowered):
            constraints.append("Timeline mentioned")
        if self._BUDGET_RE.search(lowered):
            constraints.append("Budget mentioned")
        if any(x in lowered for x in ("must", "required", "compliance", "gdpr", "hipaa")):
            constraints.append("Hard requirements mentioned")
        return constraints

    def _build_clarifications(self, lowered: str, domain: str, complexity: str) -> list[str]:
        questions: list[str] = []
        if " for " not in lowered and not any(k in lowered for k in ("user", "customer", "team", "client")):
            questions.append("Who is the primary user or audience for this project?")
        if not self._TIMELINE_RE.search(lowered):
            questions.append("What is your target timeline for the first usable version?")
        if domain == "code" and not any(
            token in lowered for token in ("web", "mobile", "desktop", "api", "bot")
        ):
            questions.append("Which platform should be prioritized first (web, mobile, API, or desktop)?")
        if complexity in {"L", "XL"} and "team" not in lowered:
            questions.append("Do you want this planned for solo execution or with collaborators?")
        return questions[:4]

    def _suggest_research_depth(self, lowered: str, domain: str, complexity: str) -> str:
        if "no research" in lowered or "skip research" in lowered:
            return "none"
        if complexity == "S":
            return "quick"
        if complexity == "XL":
            return "deep"
        if domain in {"business", "events"} and complexity in {"L", "XL"}:
            return "deep"
        return "standard"

    def _split_ai_human(self, domain: str) -> tuple[list[str], list[str]]:
        ai_common = [
            "Generate structured plans and task breakdowns",
            "Draft PRD/requirements and implementation steps",
            "Propose tools, architecture, and execution order",
        ]
        human_common = [
            "Approve tradeoffs and final scope decisions",
            "Provide credentials, access, and external approvals",
            "Validate outcomes against business or personal goals",
        ]

        if domain == "creative":
            ai_common[1] = "Generate content drafts and creative direction options"
            human_common[2] = "Select final creative direction and brand voice"
        elif domain == "events":
            ai_common[1] = "Draft runbooks, schedules, and logistics checklists"
            human_common[1] = "Confirm vendors, venue, and real-world constraints"
        elif domain == "home":
            ai_common[0] = "Generate renovation/organization plans and phased tasks"
            human_common[1] = "Approve physical changes, purchases, and safety constraints"

        return ai_common, human_common

