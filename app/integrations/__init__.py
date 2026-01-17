"""
Integration modules for external services.
"""

from .mcp_client import MCPClient
from .mcp_email import EmailClient
from .mcp_teams import TeamsClient

__all__ = ["MCPClient", "EmailClient", "TeamsClient"]
