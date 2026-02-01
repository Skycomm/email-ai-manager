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
        default="http://localhost:10001",
        description="URL of the MS365 MCP server"
    )
    ms365_mcp_bearer_token: Optional[str] = Field(
        default=None,
        description="Bearer token for MCP server authentication"
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
    teams_team_id: Optional[str] = Field(
        default=None,
        description="Teams team ID for notifications"
    )
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
    teams_morning_summary_hour: int = Field(
        default=7,
        description="Hour (0-23) to send morning summary of newsletters/FYI (default 7am)"
    )
    timezone: str = Field(
        default="Australia/Perth",
        description="Timezone for displaying times (default Australia/Perth, UTC+8)"
    )
    fyi_auto_archive_hours: int = Field(
        default=48,
        description="Auto-archive FYI/newsletter emails older than this many hours (default 48h)"
    )

    # Safety Settings
    auto_send_enabled: bool = Field(
        default=False,
        description="Allow auto-sending of low-risk emails (Phase 4)"
    )
    auto_send_max_priority: int = Field(
        default=4,
        description="Only auto-send emails with priority >= this (4-5 = low priority)"
    )
    auto_send_internal_only: bool = Field(
        default=True,
        description="Only auto-send to internal domains"
    )
    auto_send_categories: List[str] = Field(
        default_factory=lambda: ["fyi"],
        description="Categories eligible for auto-send (fyi, meeting)"
    )
    max_emails_per_hour: int = Field(
        default=20,
        description="Maximum outbound emails per hour"
    )
    external_domain_require_approval: bool = Field(
        default=True,
        description="Always require approval for external domains"
    )

    # VIP Settings
    vip_senders: List[str] = Field(
        default_factory=list,
        description="Email addresses that are always high priority"
    )
    vip_domains: List[str] = Field(
        default_factory=list,
        description="Domains that are always high priority"
    )
    internal_domains: List[str] = Field(
        default_factory=list,
        description="Internal company domains (for auto-send eligibility)"
    )

    # Calendar Integration
    calendar_integration_enabled: bool = Field(
        default=True,
        description="Enable calendar integration for meeting emails"
    )
    calendar_auto_accept_internal: bool = Field(
        default=False,
        description="Auto-accept meeting invites from internal senders"
    )
    calendar_check_conflicts: bool = Field(
        default=True,
        description="Check for calendar conflicts when suggesting meeting responses"
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

    # Spam Detection Settings
    spam_batch_size: int = Field(
        default=5,
        description="Number of spam emails to batch before notification"
    )
    spam_notification_interval_minutes: int = Field(
        default=5,
        description="Interval between spam batch notifications"
    )
    spam_auto_archive_threshold: int = Field(
        default=95,
        description="Spam score threshold for auto-archive (0-100)"
    )
    spam_ask_threshold: int = Field(
        default=70,
        description="Spam score threshold to ask user (0-100)"
    )
    spam_sender_domains: List[str] = Field(
        default_factory=list,
        description="Sender domains to always treat as spam"
    )
    spam_subject_patterns: List[str] = Field(
        default_factory=list,
        description="Subject patterns to always treat as spam"
    )

    # Notification Settings
    email_body_preview_length: int = Field(
        default=500,
        description="Length of email body preview"
    )
    teams_message_max_length: int = Field(
        default=3000,
        description="Max length for Teams messages"
    )

    # Alert senders - FYI emails from these senders/domains get immediate notification
    # (with deduplication), while other FYI emails wait for morning summary
    alert_sender_domains: List[str] = Field(
        default_factory=lambda: [
            "meraki.com",
            "uptimerobot.com",
            "pagerduty.com",
            "opsgenie.com",
            "datadog.com",
            "pingdom.com",
            "statuspage.io",
            "betterstack.com",
        ],
        description="Sender domains that trigger immediate FYI notifications (monitoring/alerts)"
    )
    alert_subject_patterns: List[str] = Field(
        default_factory=lambda: [
            "alert",
            "down",
            "offline",
            "failed",
            "critical",
            "warning",
            "error",
            "outage",
            "incident",
            "vpn",
            "connection lost",
            "unreachable",
        ],
        description="Subject patterns that trigger immediate FYI notifications"
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
