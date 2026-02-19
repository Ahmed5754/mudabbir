"""
Mudabbir Skills Module

Integrates with the AgentSkills ecosystem (skills.sh).
Loads skills from ~/.agents/skills/ and ~/.Mudabbir/skills/
"""

from .loader import SkillLoader, get_skill_loader, load_all_skills
from .executor import SkillExecutor

__all__ = [
    "SkillLoader",
    "get_skill_loader",
    "load_all_skills",
    "SkillExecutor",
]
