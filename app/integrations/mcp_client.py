"""
MCP (Model Context Protocol) client for communicating with MS365 MCP server.

This wraps the MCP server running locally that provides access to
Microsoft 365 APIs (Email, Calendar, Teams).
"""

import httpx
import logging
from typing import Any, Dict, Optional, List

from ..config import settings

logger = logging.getLogger(__name__)


class MCPClientError(Exception):
    """Error communicating with MCP server."""
    pass


class MCPClient:
    """
    Client for the MS365 MCP server.

    The MCP server exposes tools like:
    - list-mail-messages
    - get-mail-message
    - send-mail
    - list-calendar-events
    - send-channel-message
    - send-chat-message
    etc.
    """

    def __init__(self, base_url: Optional[str] = None):
        self.base_url = base_url or settings.ms365_mcp_url
        self.client = httpx.Client(timeout=30.0)

    def call_tool(self, tool_name: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Call an MCP tool.

        Args:
            tool_name: Name of the MCP tool (e.g., "list-mail-messages")
            params: Parameters for the tool

        Returns:
            Tool response as dictionary
        """
        try:
            # MCP servers typically expose tools via JSON-RPC or REST
            # Adjust this based on your specific MCP server implementation
            response = self.client.post(
                f"{self.base_url}/tools/{tool_name}",
                json=params,
                headers={"Content-Type": "application/json"}
            )
            response.raise_for_status()
            return response.json()
        except httpx.HTTPStatusError as e:
            logger.error(f"MCP tool call failed: {tool_name} - {e}")
            raise MCPClientError(f"Tool {tool_name} failed: {e}")
        except httpx.RequestError as e:
            logger.error(f"MCP connection error: {e}")
            raise MCPClientError(f"Connection error: {e}")

    # Email operations
    def list_mail_messages(
        self,
        mailbox: Optional[str] = None,
        folder: str = "inbox",
        top: int = 10,
        filter_query: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """List email messages from a mailbox."""
        params = {
            "folder": folder,
            "top": top,
        }
        if mailbox:
            params["sender_email"] = mailbox
        if filter_query:
            params["filter"] = filter_query

        result = self.call_tool("list-mail-messages", params)
        return result.get("messages", result.get("value", []))

    def get_mail_message(
        self,
        message_id: str,
        mailbox: Optional[str] = None
    ) -> Dict[str, Any]:
        """Get full details of an email message."""
        params = {"message_id": message_id}
        if mailbox:
            params["sender_email"] = mailbox
        return self.call_tool("get-mail-message", params)

    def send_mail(
        self,
        to: List[Dict[str, Any]],
        subject: str,
        body: str,
        body_type: str = "HTML",
        cc: Optional[List[Dict[str, Any]]] = None,
        sender_email: Optional[str] = None,
        importance: str = "normal"
    ) -> Dict[str, Any]:
        """Send an email."""
        params = {
            "to": to,
            "subject": subject,
            "body": {
                "content": body,
                "contentType": body_type
            },
            "importance": importance
        }
        if cc:
            params["cc"] = cc
        if sender_email:
            params["sender_email"] = sender_email

        return self.call_tool("send-mail", params)

    def reply_to_mail(
        self,
        message_id: str,
        body: str,
        body_type: str = "HTML",
        reply_all: bool = False,
        sender_email: Optional[str] = None
    ) -> Dict[str, Any]:
        """Reply to an email message."""
        # This would use the MS Graph reply endpoint
        # For now, we'll use send-mail with the original recipients
        params = {
            "message_id": message_id,
            "body": {
                "content": body,
                "contentType": body_type
            },
            "reply_all": reply_all
        }
        if sender_email:
            params["sender_email"] = sender_email

        # Note: Adjust based on MCP server capabilities
        return self.call_tool("reply-mail", params)

    # Teams operations
    def list_joined_teams(self) -> List[Dict[str, Any]]:
        """List Teams the user has joined."""
        result = self.call_tool("list-joined-teams", {})
        return result.get("teams", result.get("value", []))

    def list_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """List channels in a Team."""
        result = self.call_tool("list-team-channels", {"team_id": team_id})
        return result.get("channels", result.get("value", []))

    def send_channel_message(
        self,
        team_id: str,
        channel_id: str,
        content: str,
        content_type: str = "html"
    ) -> Dict[str, Any]:
        """Send a message to a Teams channel."""
        return self.call_tool("send-channel-message", {
            "team_id": team_id,
            "channel_id": channel_id,
            "content": content,
            "contentType": content_type
        })

    def list_channel_messages(
        self,
        team_id: str,
        channel_id: str,
        top: int = 20
    ) -> List[Dict[str, Any]]:
        """List messages from a Teams channel."""
        result = self.call_tool("list-channel-messages", {
            "team_id": team_id,
            "channel_id": channel_id,
            "top": top
        })
        return result.get("messages", result.get("value", []))

    def list_chats(self, top: int = 20) -> List[Dict[str, Any]]:
        """List Teams chats."""
        result = self.call_tool("list-chats", {"top": top})
        return result.get("chats", result.get("value", []))

    def send_chat_message(
        self,
        chat_id: str,
        content: str,
        content_type: str = "html"
    ) -> Dict[str, Any]:
        """Send a message to a Teams chat."""
        return self.call_tool("send-chat-message", {
            "chat_id": chat_id,
            "content": content,
            "contentType": content_type
        })

    def list_chat_messages(
        self,
        chat_id: str,
        top: int = 20
    ) -> List[Dict[str, Any]]:
        """List messages from a Teams chat."""
        result = self.call_tool("list-chat-messages", {
            "chat_id": chat_id,
            "top": top
        })
        return result.get("messages", result.get("value", []))

    # Calendar operations (for future use)
    def list_calendar_events(
        self,
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
        top: int = 10
    ) -> List[Dict[str, Any]]:
        """List calendar events."""
        params = {"top": top}
        if start_datetime:
            params["startDateTime"] = start_datetime
        if end_datetime:
            params["endDateTime"] = end_datetime

        result = self.call_tool("list-calendar-events", params)
        return result.get("events", result.get("value", []))

    def close(self):
        """Close the HTTP client."""
        self.client.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
