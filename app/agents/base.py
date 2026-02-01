"""
Base agent class for all AI agents.

Uses the Claude Agent SDK pattern with tool definitions and
structured outputs.
"""

import json
import logging
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any, Dict, List, Optional

import anthropic

from ..config import settings
from ..db import Database
from ..models import AuditLogEntry

logger = logging.getLogger(__name__)


class BaseAgent(ABC):
    """
    Base class for all AI agents.

    Each agent:
    - Has access to specific tools
    - Logs all actions to the audit log
    - Uses Claude for decision making
    """

    def __init__(
        self,
        db: Database,
        name: str,
        system_prompt: str,
        tools: Optional[List[Dict[str, Any]]] = None
    ):
        self.db = db
        self.name = name
        self.system_prompt = system_prompt
        self.tools = tools or []
        self.client = anthropic.Anthropic(api_key=settings.anthropic_api_key)

    @abstractmethod
    async def process(self, *args, **kwargs) -> Any:
        """Process the agent's main task."""
        pass

    def call_claude(
        self,
        messages: List[Dict[str, Any]],
        max_tokens: int = 4096,
        temperature: float = 0.7
    ) -> str:
        """
        Call Claude with the agent's system prompt and tools.

        Args:
            messages: Conversation messages
            max_tokens: Maximum tokens in response
            temperature: Sampling temperature

        Returns:
            Claude's response text
        """
        try:
            response = self.client.messages.create(
                model=settings.agent_model,
                max_tokens=max_tokens,
                system=self.system_prompt,
                messages=messages,
                tools=self.tools if self.tools else anthropic.NOT_GIVEN,
                temperature=temperature
            )

            # Extract text response
            text_content = ""
            for block in response.content:
                if hasattr(block, "text"):
                    text_content += block.text

            return text_content

        except anthropic.APIError as e:
            logger.error(f"Claude API error in {self.name}: {e}")
            raise

    def call_claude_structured(
        self,
        messages: List[Dict[str, Any]],
        response_schema: Dict[str, Any],
        max_tokens: int = 4096
    ) -> Dict[str, Any]:
        """
        Call Claude expecting a structured JSON response.

        Args:
            messages: Conversation messages
            response_schema: JSON schema for expected response
            max_tokens: Maximum tokens

        Returns:
            Parsed JSON response
        """
        # Add schema instruction to messages
        schema_instruction = f"""
        Respond with valid JSON matching this schema:
        {json.dumps(response_schema, indent=2)}

        Only output the JSON, no other text.
        """

        augmented_messages = messages.copy()
        if augmented_messages:
            last_msg = augmented_messages[-1]
            if isinstance(last_msg.get("content"), str):
                last_msg["content"] += f"\n\n{schema_instruction}"

        response_text = self.call_claude(augmented_messages, max_tokens, temperature=0.3)

        # Parse JSON from response
        try:
            # Try to extract JSON if wrapped in markdown
            if "```json" in response_text:
                json_start = response_text.find("```json") + 7
                json_end = response_text.find("```", json_start)
                response_text = response_text[json_start:json_end].strip()
            elif "```" in response_text:
                json_start = response_text.find("```") + 3
                json_end = response_text.find("```", json_start)
                response_text = response_text[json_start:json_end].strip()

            return json.loads(response_text)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse Claude response as JSON: {e}")
            logger.error(f"Response was: {response_text[:500]}")
            raise ValueError(f"Invalid JSON response from Claude: {e}")

    def log_action(
        self,
        action: str,
        email_id: Optional[str] = None,
        details: Optional[Dict[str, Any]] = None,
        user_command: Optional[str] = None,
        success: bool = True,
        error: Optional[str] = None
    ) -> None:
        """
        Log an action to the audit log.

        Args:
            action: Description of the action taken
            email_id: Related email ID if applicable
            details: Additional context
            user_command: User command that triggered this (if any)
            success: Whether the action succeeded
            error: Error message if failed
        """
        entry = AuditLogEntry(
            email_id=email_id,
            timestamp=datetime.utcnow(),
            agent=self.name,
            action=action,
            details=details or {},
            user_command=user_command,
            success=success,
            error=error
        )
        self.db.log_audit(entry)

        log_msg = f"[{self.name}] {action}"
        if email_id:
            log_msg += f" (email: {email_id[:8]}...)"
        if error:
            logger.error(f"{log_msg} - ERROR: {error}")
        else:
            logger.info(log_msg)


class AgentTool:
    """Decorator for defining agent tools."""

    def __init__(self, name: str, description: str, parameters: Dict[str, Any]):
        self.name = name
        self.description = description
        self.parameters = parameters

    def to_dict(self) -> Dict[str, Any]:
        """Convert to Claude tool format."""
        return {
            "name": self.name,
            "description": self.description,
            "input_schema": {
                "type": "object",
                "properties": self.parameters,
                "required": list(self.parameters.keys())
            }
        }
