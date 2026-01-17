"""
Email AI Manager - Main entry point.

A locally-running AI email management system that uses Claude Agent SDK
and MS365 MCP to triage, summarize, draft replies, and manage workflows
through Microsoft Teams.
"""

import asyncio
import time
import logging
import signal
import sys
from datetime import datetime

from .config import settings
from .db import Database
from .agents import CoordinatorAgent

# Configure logging
logging.basicConfig(
    level=getattr(logging, settings.log_level),
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
    ]
)

if settings.log_file:
    file_handler = logging.FileHandler(settings.log_file)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    ))
    logging.getLogger().addHandler(file_handler)

logger = logging.getLogger(__name__)


class EmailManager:
    """Main application class."""

    def __init__(self):
        self.db = Database(settings.db_path)
        self.coordinator = CoordinatorAgent(self.db)
        self.running = False
        self._shutdown_event = asyncio.Event()

    async def start(self):
        """Start the main polling loop."""
        self.running = True
        logger.info("=" * 60)
        logger.info("Email AI Manager starting...")
        logger.info(f"Version: 0.1.0")
        logger.info(f"Poll interval: {settings.poll_interval_seconds} seconds")
        logger.info(f"Database: {settings.db_path}")
        logger.info(f"Mailbox: {settings.mailbox_email}")
        logger.info(f"Auto-send enabled: {settings.auto_send_enabled}")
        logger.info("=" * 60)

        # Rehydrate pending items
        self._rehydrate()

        while self.running:
            try:
                await self._poll_cycle()
            except Exception as e:
                logger.error(f"Error in poll cycle: {e}", exc_info=True)

            # Wait for next poll or shutdown
            try:
                await asyncio.wait_for(
                    self._shutdown_event.wait(),
                    timeout=settings.poll_interval_seconds
                )
                # If we get here, shutdown was requested
                break
            except asyncio.TimeoutError:
                # Normal timeout, continue polling
                pass

        logger.info("Email AI Manager stopped.")

    def stop(self):
        """Stop the polling loop gracefully."""
        logger.info("Shutdown requested...")
        self.running = False
        self._shutdown_event.set()
        self.coordinator.close()

    def _rehydrate(self):
        """Rehydrate pending items from database after restart."""
        pending = self.db.get_pending_emails()
        if pending:
            logger.info(f"Rehydrated {len(pending)} pending email(s) from previous session:")
            for email in pending[:5]:
                logger.info(f"  - [{email.state.value}] {email.subject[:50]}")
            if len(pending) > 5:
                logger.info(f"  ... and {len(pending) - 5} more")

    async def _poll_cycle(self):
        """Execute a single poll cycle."""
        cycle_start = datetime.utcnow()
        logger.info(f"Starting poll cycle at {cycle_start.isoformat()}")

        try:
            summary = await self.coordinator.process()

            logger.info(
                f"Poll cycle complete: "
                f"{summary['new_emails']} new, "
                f"{summary['processed']} processed, "
                f"{summary['drafts_generated']} drafts, "
                f"{summary['notifications_sent']} notifications, "
                f"{summary['errors']} errors"
            )

        except Exception as e:
            logger.error(f"Poll cycle failed: {e}", exc_info=True)

        cycle_duration = (datetime.utcnow() - cycle_start).total_seconds()
        logger.debug(f"Cycle completed in {cycle_duration:.2f}s")


def main():
    """Main entry point."""
    manager = EmailManager()

    # Handle graceful shutdown
    def signal_handler(signum, frame):
        logger.info(f"Received signal {signum}")
        manager.stop()

    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)

    # Run the async main loop
    try:
        asyncio.run(manager.start())
    except KeyboardInterrupt:
        logger.info("Interrupted by user")
        manager.stop()


if __name__ == "__main__":
    main()
