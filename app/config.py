"""
Configuration management using Pydantic settings.
"""

import os
from typing import Optional, List
from pydantic_settings import BaseSettings
from pydantic import Field


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    # API Keys
    anthropic_api_key: str = Field(..., description="Anthropic API key for Claude")

    # MCP Server Settings
    ms365_mcp_url: str = Field(
        default="http://localhost:3000",
        description="URL of the MS365 MCP server"
    )

    # Email Settings
    poll_interval_seconds: int = Field(
        default=60,
        description="How often to poll for new emails (seconds)"
    )
    mailbox_email: str = Field(
        ...,
        description="Primary mailbox email to monitor"
    )
    shared_mailbox_emails: List[str] = Field(
        default_factory=list,
        description="Additional shared mailboxes to monitor (Phase 4)"
    )

    # Teams Settings
    teams_channel_id: Optional[str] = Field(
        default=None,
        description="Teams channel ID for notifications"
    )
    teams_chat_id: Optional[str] = Field(
        default=None,
        description="Teams chat ID for 1:1 notifications"
    )
    teams_notify_urgent: bool = Field(
        default=True,
        description="Send immediate notification for urgent emails"
    )
    teams_daily_digest_time: str = Field(
        default="08:00",
        description="Time to send daily digest (HH:MM)"
    )

    # Safety Settings
    auto_send_enabled: bool = Field(
        default=False,
        description="Allow auto-sending of low-risk emails (Phase 4)"
    )
    max_emails_per_hour: int = Field(
        default=20,
        description="Maximum outbound emails per hour"
    )
    external_domain_require_approval: bool = Field(
        default=True,
        description="Always require approval for external domains"
    )

    # Database Settings
    db_path: str = Field(
        default="/data/email_manager.db",
        description="Path to SQLite database"
    )
    encrypt_email_bodies: bool = Field(
        default=True,
        description="Encrypt email bodies at rest"
    )
    encryption_key: Optional[str] = Field(
        default=None,
        description="Encryption key for email bodies (generated if not provided)"
    )

    # Dashboard Settings (Phase 3)
    dashboard_enabled: bool = Field(
        default=False,
        description="Enable web dashboard"
    )
    dashboard_port: int = Field(
        default=8080,
        description="Dashboard HTTP port"
    )
    dashboard_api_key: Optional[str] = Field(
        default=None,
        description="API key for dashboard authentication"
    )

    # Logging
    log_level: str = Field(
        default="INFO",
        description="Logging level"
    )
    log_file: Optional[str] = Field(
        default=None,
        description="Log file path (stdout if not set)"
    )

    # Agent Settings
    agent_model: str = Field(
        default="claude-sonnet-4-20250514",
        description="Claude model to use for agents"
    )
    max_tokens: int = Field(
        default=4096,
        description="Max tokens for agent responses"
    )

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"
        extra = "ignore"

    @property
    def all_mailboxes(self) -> List[str]:
        """Return all mailboxes to monitor."""
        mailboxes = [self.mailbox_email]
        mailboxes.extend(self.shared_mailbox_emails)
        return mailboxes


# Global settings instance
settings = Settings()
