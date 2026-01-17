"""
AI Agents for email management.

Built using the Claude Agent SDK pattern with specialized agents
for different tasks.
"""

from .base import BaseAgent
from .coordinator import CoordinatorAgent
from .drafting import DraftingAgent
from .teams_comms import TeamsCommsAgent
from .spam_filter import SpamFilterAgent

__all__ = [
    "BaseAgent",
    "CoordinatorAgent",
    "DraftingAgent",
    "TeamsCommsAgent",
    "SpamFilterAgent",
]
