"""
MCP (Model Context Protocol) client for communicating with MS365 MCP server.

This wraps the MCP server running locally that provides access to
Microsoft 365 APIs (Email, Calendar, Teams).

Uses JSON-RPC 2.0 over HTTP with Bearer token authentication.
"""

import json
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

    Uses JSON-RPC 2.0 protocol with Bearer token authentication.
    """

    def __init__(self, base_url: Optional[str] = None, bearer_token: Optional[str] = None):
        self.base_url = base_url or settings.ms365_mcp_url
        self.bearer_token = bearer_token or getattr(settings, 'ms365_mcp_bearer_token', None)
        self._request_id = 0
        self.client = httpx.Client(timeout=60.0)

    def _get_headers(self) -> Dict[str, str]:
        """Get request headers with authentication."""
        headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream"
        }
        if self.bearer_token:
            headers["Authorization"] = f"Bearer {self.bearer_token}"
        return headers

    def _next_request_id(self) -> int:
        """Get next JSON-RPC request ID."""
        self._request_id += 1
        return self._request_id

    def call_tool(self, tool_name: str, params: Dict[str, Any]) -> Dict[str, Any]:
        """
        Call an MCP tool using JSON-RPC 2.0.

        Args:
            tool_name: Name of the MCP tool (e.g., "list-mail-messages")
            params: Parameters for the tool

        Returns:
            Tool response as dictionary
        """
        try:
            # Build JSON-RPC 2.0 request
            request_body = {
                "jsonrpc": "2.0",
                "method": "tools/call",
                "params": {
                    "name": tool_name,
                    "arguments": params
                },
                "id": self._next_request_id()
            }

            logger.debug(f"MCP call: {tool_name} with params: {params}")

            response = self.client.post(
                f"{self.base_url}/mcp",
                json=request_body,
                headers=self._get_headers()
            )
            response.raise_for_status()

            # Parse SSE response format
            result = self._parse_sse_response(response.text)

            if "error" in result:
                error_msg = result.get("error", {})
                if isinstance(error_msg, dict):
                    error_msg = error_msg.get("message", str(error_msg))
                logger.error(f"MCP tool error: {tool_name} - {error_msg}")
                raise MCPClientError(f"Tool {tool_name} failed: {error_msg}")

            # Extract the actual content from the response
            result_obj = result.get("result", {})
            if result_obj is None:
                return {}
            content = result_obj.get("content", [])
            if content and len(content) > 0:
                first_content = content[0]
                if first_content is None:
                    return {}
                text_content = first_content.get("text", "{}")
                try:
                    parsed = json.loads(text_content)
                    return parsed if parsed is not None else {}
                except json.JSONDecodeError:
                    # Some MCP responses have text prefix before JSON
                    # Try to extract JSON from the text
                    if "{" in text_content:
                        json_start = text_content.find("{")
                        try:
                            parsed = json.loads(text_content[json_start:])
                            return parsed if parsed is not None else {}
                        except json.JSONDecodeError:
                            pass
                    return {"text": text_content}

            return result_obj if result_obj is not None else {}

        except httpx.HTTPStatusError as e:
            logger.error(f"MCP tool call failed: {tool_name} - {e}")
            raise MCPClientError(f"Tool {tool_name} failed: {e}")
        except httpx.RequestError as e:
            logger.error(f"MCP connection error: {e}")
            raise MCPClientError(f"Connection error: {e}")

    def _parse_sse_response(self, response_text: str) -> Dict[str, Any]:
        """Parse Server-Sent Events response format."""
        # The MCP server returns SSE format: "event: message\ndata: {...}\n\n"
        for line in response_text.split('\n'):
            if line.startswith('data: '):
                try:
                    return json.loads(line[6:])
                except json.JSONDecodeError:
                    pass

        # If not SSE format, try direct JSON
        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            return {"error": f"Could not parse response: {response_text[:200]}"}

    # Email operations
    def list_mail_messages(
        self,
        mailbox: Optional[str] = None,
        folder: str = "inbox",
        top: int = 10,
        filter_query: Optional[str] = None,
        orderby: Optional[str] = None
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
        if orderby:
            params["orderby"] = orderby

        result = self.call_tool("list-mail-messages", params)

        # Handle different response formats
        if isinstance(result, list):
            return result
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

        return self.call_tool("reply-mail", params)

    def move_mail_message(
        self,
        message_id: str,
        destination_folder_id: str,
        sender_email: Optional[str] = None
    ) -> Dict[str, Any]:
        """Move an email to a different folder."""
        params = {
            "message_id": message_id,
            "destination_folder_id": destination_folder_id
        }
        if sender_email:
            params["sender_email"] = sender_email

        return self.call_tool("move-mail-message", params)

    def list_mail_folders(
        self,
        mailbox: Optional[str] = None,
        top: int = 100
    ) -> List[Dict[str, Any]]:
        """
        List mail folders in a mailbox.

        Args:
            mailbox: Email address of the mailbox
            top: Maximum number of folders to return (default 100)

        Returns:
            List of folder dictionaries with id, displayName, childFolderCount, etc.
        """
        params = {"top": top}
        if mailbox:
            params["sender_email"] = mailbox
        result = self.call_tool("list-mail-folders", params)

        if isinstance(result, list):
            return result
        return result.get("folders", result.get("value", []))

    def list_child_mail_folders(
        self,
        folder_id: str,
        mailbox: Optional[str] = None,
        top: int = 100
    ) -> List[Dict[str, Any]]:
        """
        List child folders (subfolders) of a specific mail folder.

        Args:
            folder_id: ID of the parent folder
            mailbox: Email address of the mailbox
            top: Maximum number of folders to return (default 100)

        Returns:
            List of child folder dictionaries
        """
        params = {"folder_id": folder_id, "top": top}
        if mailbox:
            params["sender_email"] = mailbox
        result = self.call_tool("list-child-mail-folders", params)
        if isinstance(result, list):
            return result
        return result.get("folders", result.get("value", []))

    def list_all_mail_folders_recursive(
        self,
        mailbox: Optional[str] = None,
        max_depth: int = 3
    ) -> List[Dict[str, Any]]:
        """
        List all mail folders recursively, including nested subfolders.

        Args:
            mailbox: Email address of the mailbox
            max_depth: Maximum depth to recurse (default 3)

        Returns:
            List of folders with nested 'children' property for subfolders
        """
        def fetch_folders(parent_id: Optional[str], depth: int) -> List[Dict[str, Any]]:
            if depth > max_depth:
                return []

            try:
                if parent_id:
                    folders = self.list_child_mail_folders(folder_id=parent_id, mailbox=mailbox)
                else:
                    folders = self.list_mail_folders(mailbox=mailbox)
            except Exception as e:
                logger.warning(f"Could not fetch folders (parent={parent_id}): {e}")
                return []

            result = []
            for folder in folders:
                folder_data = {
                    "id": folder.get("id"),
                    "name": folder.get("displayName"),
                    "total_count": folder.get("totalItemCount", 0),
                    "unread_count": folder.get("unreadItemCount", 0),
                    "child_folder_count": folder.get("childFolderCount", 0),
                    "children": []
                }

                # Recursively fetch children if this folder has any
                if folder_data["child_folder_count"] > 0 and depth < max_depth:
                    folder_data["children"] = fetch_folders(folder_data["id"], depth + 1)

                result.append(folder_data)

            return result

        return fetch_folders(None, 1)

    # Teams operations
    def list_joined_teams(self) -> List[Dict[str, Any]]:
        """List Teams the user has joined."""
        result = self.call_tool("list-joined-teams", {})
        if isinstance(result, list):
            return result
        return result.get("teams", result.get("value", []))

    def list_team_channels(self, team_id: str) -> List[Dict[str, Any]]:
        """List channels in a Team."""
        result = self.call_tool("list-team-channels", {"team_id": team_id})
        if isinstance(result, list):
            return result
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
        if result is None:
            return []
        if isinstance(result, list):
            return [m for m in result if m is not None]
        messages = result.get("messages", result.get("value", []))
        if messages is None:
            return []
        return [m for m in messages if m is not None]

    def list_chats(self, top: int = 20) -> List[Dict[str, Any]]:
        """List Teams chats."""
        result = self.call_tool("list-chats", {"top": top})
        if isinstance(result, list):
            return result
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
        if isinstance(result, list):
            return result
        return result.get("messages", result.get("value", []))

    def list_channel_message_replies(
        self,
        team_id: str,
        channel_id: str,
        message_id: str,
        top: int = 20
    ) -> List[Dict[str, Any]]:
        """List replies to a Teams channel message."""
        result = self.call_tool("list-channel-message-replies", {
            "team_id": team_id,
            "channel_id": channel_id,
            "message_id": message_id,
            "top": top
        })
        if result is None:
            return []
        if isinstance(result, list):
            return [m for m in result if m is not None]
        replies = result.get("replies", result.get("value", []))
        if replies is None:
            return []
        return [m for m in replies if m is not None]

    def list_chat_message_replies(
        self,
        chat_id: str,
        message_id: str,
        top: int = 20
    ) -> List[Dict[str, Any]]:
        """List replies to a Teams chat message (quote replies)."""
        result = self.call_tool("list-chat-message-replies", {
            "chat_id": chat_id,
            "message_id": message_id,
            "top": top
        })
        if result is None:
            return []
        if isinstance(result, list):
            return [m for m in result if m is not None]
        replies = result.get("replies", result.get("value", []))
        if replies is None:
            return []
        return [m for m in replies if m is not None]

    def get_conversation_messages(
        self,
        mailbox: str,
        conversation_id: str,
        top: int = 10
    ) -> List[Dict[str, Any]]:
        """
        Get all messages in an email conversation/thread.

        Args:
            mailbox: Email address of mailbox
            conversation_id: MS365 conversation ID
            top: Maximum messages to return

        Returns:
            List of email messages in the conversation
        """
        try:
            # Filter messages by conversation ID
            filter_query = f"conversationId eq '{conversation_id}'"
            params = {
                "folder": "inbox",
                "top": top,
                "filter": filter_query,
                "orderby": "receivedDateTime desc"
            }
            if mailbox:
                params["sender_email"] = mailbox

            result = self.call_tool("list-mail-messages", params)

            if isinstance(result, list):
                return result
            return result.get("messages", result.get("value", []))
        except MCPClientError as e:
            logger.warning(f"Could not fetch conversation history: {e}")
            return []

    # Calendar operations
    def list_calendar_events(
        self,
        start_datetime: Optional[str] = None,
        end_datetime: Optional[str] = None,
        top: int = 10,
        organizer_email: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """List calendar events."""
        params = {"top": top}
        if start_datetime:
            params["startDateTime"] = start_datetime
        if end_datetime:
            params["endDateTime"] = end_datetime
        if organizer_email:
            params["organizer_email"] = organizer_email

        result = self.call_tool("list-calendar-events", params)
        if isinstance(result, list):
            return result
        return result.get("events", result.get("value", []))

    def get_calendar_view(
        self,
        start_datetime: str,
        end_datetime: str,
        user_email: Optional[str] = None,
        timezone: str = "Australia/Perth"
    ) -> List[Dict[str, Any]]:
        """Get calendar view for a time range."""
        params = {
            "start_date_time": start_datetime,
            "end_date_time": end_datetime,
            "timezone": timezone
        }
        if user_email:
            params["user_email"] = user_email

        result = self.call_tool("get-calendar-view", params)
        if isinstance(result, list):
            return result
        return result.get("events", result.get("value", []))

    def accept_event_invite(
        self,
        event_id: str,
        comment: Optional[str] = None,
        user_email: Optional[str] = None
    ) -> Dict[str, Any]:
        """Accept a meeting invitation."""
        params = {"event_id": event_id}
        if comment:
            params["comment"] = comment
        if user_email:
            params["user_email"] = user_email
        return self.call_tool("accept-event-invite", params)

    def decline_event_invite(
        self,
        event_id: str,
        comment: Optional[str] = None,
        user_email: Optional[str] = None
    ) -> Dict[str, Any]:
        """Decline a meeting invitation."""
        params = {"event_id": event_id}
        if comment:
            params["comment"] = comment
        if user_email:
            params["user_email"] = user_email
        return self.call_tool("decline-event-invite", params)

    def tentatively_accept_event_invite(
        self,
        event_id: str,
        comment: Optional[str] = None,
        user_email: Optional[str] = None
    ) -> Dict[str, Any]:
        """Tentatively accept a meeting invitation."""
        params = {"event_id": event_id}
        if comment:
            params["comment"] = comment
        if user_email:
            params["user_email"] = user_email
        return self.call_tool("tentatively-accept-event-invite", params)

    def update_chat_message(
        self,
        chat_id: str,
        message_id: str,
        content: str,
        content_type: str = "html"
    ) -> Dict[str, Any]:
        """Update an existing chat message."""
        return self.call_tool("update-chat-message", {
            "chat_id": chat_id,
            "message_id": message_id,
            "body": {
                "content": content,
                "contentType": content_type
            }
        })

    def update_channel_message(
        self,
        team_id: str,
        channel_id: str,
        message_id: str,
        content: str,
        content_type: str = "html"
    ) -> Dict[str, Any]:
        """Update an existing channel message."""
        return self.call_tool("update-channel-message", {
            "team_id": team_id,
            "channel_id": channel_id,
            "message_id": message_id,
            "body": {
                "content": content,
                "contentType": content_type
            }
        })

    def close(self):
        """Close the HTTP client."""
        self.client.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
