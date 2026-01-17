# Email AI Manager

An AI-powered email management system using Claude Agent SDK, Microsoft 365 MCP integration, and Microsoft Teams for workflow automation.

## Overview

Email AI Manager is a locally-running system that helps you manage email more efficiently by:

- **Triaging emails** - Automatically categorize and prioritize incoming emails
- **Summarizing content** - Get concise summaries of email contents
- **Drafting replies** - AI generates contextual reply drafts
- **Human-in-the-loop** - All outbound emails require explicit approval via Teams
- **Learning from behavior** - System learns your spam/routing preferences over time

```
┌─────────────────────────────────────────────────────────────────────┐
│                         EMAIL AI MANAGER                             │
├─────────────────────────────────────────────────────────────────────┤
│                                                                      │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐          │
│  │   INGESTION  │───▶│  COORDINATOR │───▶│    OUTPUT    │          │
│  │    AGENT     │    │    AGENT     │    │    AGENTS    │          │
│  └──────────────┘    └──────────────┘    └──────────────┘          │
│         │                   │                   │                   │
│         ▼                   ▼                   ▼                   │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐          │
│  │ Spam Filter  │    │   Drafting   │    │ Teams Comms  │          │
│  │    Agent     │    │    Agent     │    │    Agent     │          │
│  └──────────────┘    └──────────────┘    └──────────────┘          │
│                                                                      │
├─────────────────────────────────────────────────────────────────────┤
│                         SHARED SERVICES                              │
│  ┌────────────┐  ┌────────────┐  ┌────────────┐  ┌────────────┐    │
│  │  SQLite DB │  │ MCP Client │  │ Audit Log  │  │  Scheduler │    │
│  └────────────┘  └────────────┘  └────────────┘  └────────────┘    │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ▼
              ┌───────────────────────────────┐
              │      MS365 MCP SERVER         │
              │  (Email, Calendar, Teams)     │
              └───────────────────────────────┘
```

## Features

### Phase 1 (Current)
- ✅ Poll emails from Microsoft 365 mailbox
- ✅ AI-powered email categorization (Urgent, Action Required, FYI, etc.)
- ✅ Automatic email summarization
- ✅ Context-aware reply draft generation
- ✅ Teams notifications with approval workflow
- ✅ Command parsing from Teams replies
- ✅ Full audit logging
- ✅ Docker containerization

### Phase 2 (Planned)
- 🔲 Spam filter agent with learning
- 🔲 Email routing/forwarding suggestions
- 🔲 Daily digest notifications
- 🔲 VIP sender rules

### Phase 3 (Planned)
- 🔲 Web dashboard for bulk actions
- 🔲 Analytics and reporting
- 🔲 Spam rule management UI

### Phase 4 (Planned)
- 🔲 Multi-mailbox support (shared mailboxes)
- 🔲 Auto-send for low-risk emails (configurable)
- 🔲 Calendar integration
- 🔲 Thread-aware responses

## Prerequisites

1. **Anthropic API Key** - For Claude AI
2. **MS365 MCP Server** - Running locally with access to your Microsoft 365 account
3. **Docker** - For containerized deployment
4. **Microsoft Teams** - For notifications and approvals

## Quick Start

### 1. Clone the repository

```bash
git clone https://github.com/Skycomm/email-ai-manager.git
cd email-ai-manager
```

### 2. Configure environment

```bash
cp .env.example .env
# Edit .env with your settings
```

Required settings:
- `ANTHROPIC_API_KEY` - Your Anthropic API key
- `MAILBOX_EMAIL` - Email address to monitor
- `MS365_MCP_URL` - URL of your MS365 MCP server
- `TEAMS_CHANNEL_ID` or `TEAMS_CHAT_ID` - Where to send notifications

### 3. Start with Docker Compose

```bash
cd docker
docker-compose up -d
```

### 4. View logs

```bash
docker-compose logs -f
```

## Configuration

See `.env.example` for all available options. Key settings:

| Variable | Description | Default |
|----------|-------------|---------|
| `POLL_INTERVAL_SECONDS` | How often to check for emails | 60 |
| `AUTO_SEND_ENABLED` | Allow auto-sending (Phase 4) | false |
| `MAX_EMAILS_PER_HOUR` | Rate limit for outbound | 20 |
| `AGENT_MODEL` | Claude model to use | claude-sonnet-4-20250514 |

## Teams Integration

### Notification Format

When a new email arrives that needs attention, you'll receive a Teams message like:

```
📧 New Email Requiring Action
━━━━━━━━━━━━━━━━━━━━━━━━━━━

From: John Smith <john@vendor.com>
Subject: Q1 Invoice Payment Query
Priority: ⚡ High
Category: 💼 Action Required

📝 Summary:
John is asking about the status of invoice #4521
from December. He mentions the payment was due
Jan 15th and asks for an update.

✉️ Draft Reply:
"Hi John, Thanks for following up. I'll check
with our accounts team and get back to you by
end of day tomorrow with an update on invoice
#4521. Best regards, David"

━━━━━━━━━━━━━━━━━━━━━━━━━━━
Token: [a1b2c3]

Reply with:
• "approve" or "a1b2c3" - Send this reply
• "edit: [your changes]" - Modify the draft
• "rewrite" - Generate a new draft
• "ignore" - Skip, no reply needed
• "more" - Show full email
• "spam" - Mark as spam
```

### Available Commands

| Command | Description |
|---------|-------------|
| `approve` / `send` / `yes` | Send the draft reply |
| `[token]` | Approve specific email by token |
| `edit: [changes]` | Modify the draft with your instructions |
| `rewrite` | Generate a completely new draft |
| `ignore` / `skip` | Mark as handled, don't reply |
| `more` | Show the full email content |
| `spam` | Mark as spam and learn pattern |
| `forward [email]` | Forward to another person |

## Architecture

### Agents

1. **Coordinator Agent** - Orchestrates the workflow, routes emails to specialists
2. **Drafting Agent** - Generates summaries and reply drafts
3. **Teams Comms Agent** - Handles all Teams interactions
4. **Spam Filter Agent** (Phase 2) - Identifies and learns spam patterns
5. **Routing Agent** (Phase 2) - Suggests forwarding to colleagues

### State Machine

```
NEW
 │
 ├──▶ SPAM_DETECTED ──▶ ARCHIVED / DELETED
 │
 ├──▶ FYI_NOTIFIED ──▶ ACKNOWLEDGED
 │
 └──▶ ACTION_REQUIRED
      │
      ├──▶ DRAFT_GENERATED
      │    │
      │    └──▶ AWAITING_APPROVAL
      │         │
      │         ├──▶ APPROVED ──▶ SENT
      │         ├──▶ EDITED ──▶ AWAITING_APPROVAL
      │         └──▶ IGNORED
      │
      └──▶ FORWARD_SUGGESTED ──▶ FORWARDED
```

### Database Schema

SQLite database with tables for:
- `emails` - All tracked emails with state and drafts
- `audit_log` - Complete action history
- `spam_rules` - Learned spam patterns
- `processed_messages` - Deduplication tracking

## Development

### Local Development

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run locally
python -m app.main
```

### Project Structure

```
email-ai-manager/
├── app/
│   ├── __init__.py
│   ├── main.py              # Entry point
│   ├── config.py            # Configuration
│   ├── models.py            # Data models
│   ├── db.py                # Database operations
│   │
│   ├── agents/
│   │   ├── base.py          # Base agent class
│   │   ├── coordinator.py   # Main orchestrator
│   │   ├── drafting.py      # Email drafting
│   │   └── teams_comms.py   # Teams integration
│   │
│   └── integrations/
│       ├── mcp_client.py    # MCP server client
│       ├── mcp_email.py     # Email operations
│       └── mcp_teams.py     # Teams operations
│
├── migrations/
│   └── 001_initial.sql      # Database schema
│
├── docker/
│   ├── Dockerfile
│   └── docker-compose.yml
│
├── .env.example
├── requirements.txt
└── README.md
```

## Security

- **No auto-sending** - All outbound emails require explicit approval
- **Encrypted storage** - Email bodies encrypted at rest (optional)
- **Rate limiting** - Maximum emails per hour enforced
- **Audit logging** - Complete history of all actions
- **Non-root Docker** - Container runs as unprivileged user
- **Local execution** - All data stays on your infrastructure

## Troubleshooting

### MCP Connection Failed

Ensure your MS365 MCP server is running and accessible:

```bash
curl http://localhost:3000/health
```

### No Teams Notifications

1. Verify `TEAMS_CHANNEL_ID` or `TEAMS_CHAT_ID` is set correctly
2. Check MCP server has Teams permissions
3. Review logs for errors: `docker-compose logs -f`

### Database Locked

If you see SQLite locking errors:

```bash
docker-compose restart email-ai-manager
```

## License

MIT License - see LICENSE file for details.

## Contributing

Contributions welcome! Please read CONTRIBUTING.md for guidelines.

## Roadmap

See the [project plan](/Users/david/.claude/plans/) for detailed implementation phases.

---

Built with ❤️ using Claude Agent SDK and Microsoft 365 MCP
