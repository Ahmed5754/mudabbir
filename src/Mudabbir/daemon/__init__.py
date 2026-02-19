"""
Mudabbir Proactive Daemon Module

Transforms Mudabbir from a reactive chatbot into a proactive AI agent
that initiates actions based on user-defined "intentions" and various triggers.
"""

from .intentions import IntentionStore, get_intention_store
from .triggers import TriggerEngine
from .context import ContextHub
from .executor import IntentionExecutor
from .proactive import ProactiveDaemon, get_daemon

__all__ = [
    "IntentionStore",
    "get_intention_store",
    "TriggerEngine",
    "ContextHub",
    "IntentionExecutor",
    "ProactiveDaemon",
    "get_daemon",
]
