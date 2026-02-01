"""
Rules Agent - Evaluates LLM-based email routing rules.

This agent:
- Evaluates emails against natural language rules
- Determines which rules match an email
- Executes rule actions (move to folder, forward, etc.)
"""

import logging
from typing import Dict, Any, List, Optional, Tuple

from .base import BaseAgent
from ..db import Database
from ..models import EmailRecord, EmailRule, RuleAction

logger = logging.getLogger(__name__)


RULES_SYSTEM_PROMPT = """You are the Email Rules Agent, responsible for evaluating emails against routing rules.

Your job is to determine if an email matches a given rule condition.

Rules use natural language descriptions like:
- "Invoices from subscription services like iTunes, Netflix, or Spotify"
- "Meeting confirmations from booking systems"
- "Shipping notifications and order confirmations"
- "Marketing emails from software vendors"

You will be given an email's details and a rule condition.
Respond with a JSON object containing:
- "matches": true/false - whether the email matches the rule
- "confidence": 0-100 - how confident you are in the match
- "reason": brief explanation of why it matches or doesn't

Be precise but reasonable:
- Match based on the sender, subject, and content
- Consider the spirit of the rule, not just exact keyword matches
- High confidence (80+) for clear matches
- Medium confidence (50-79) for probable matches
- Low confidence (<50) for uncertain matches
"""


class RulesAgent(BaseAgent):
    """
    Agent for evaluating LLM-based email routing rules.
    """

    def __init__(self, db: Database):
        super().__init__(
            db=db,
            name="rules",
            system_prompt=RULES_SYSTEM_PROMPT
        )

    async def process(self, email: EmailRecord) -> Dict[str, Any]:
        """
        Evaluate email against all active rules and apply matches.

        Args:
            email: The email to evaluate

        Returns:
            Dict with matched rules and actions taken
        """
        rules = self.db.get_active_email_rules()

        if not rules:
            return {
                "matched_rules": [],
                "actions_taken": [],
                "message": "No active rules to evaluate"
            }

        matched_rules = []
        actions_taken = []

        for rule in rules:
            match_result = await self.evaluate_rule(email, rule)

            if match_result["matches"] and match_result["confidence"] >= 50:
                matched_rules.append({
                    "rule_id": rule.id,
                    "rule_name": rule.name,
                    "confidence": match_result["confidence"],
                    "reason": match_result["reason"]
                })

                # Record the hit
                self.db.increment_email_rule_hit(rule.id)

                # Log the match
                self.log_action(
                    "rule_matched",
                    email_id=email.id,
                    details={
                        "rule_id": rule.id,
                        "rule_name": rule.name,
                        "confidence": match_result["confidence"],
                        "reason": match_result["reason"]
                    }
                )

                # If this rule stops further processing, break
                if rule.stop_processing:
                    break

        return {
            "matched_rules": matched_rules,
            "actions_taken": actions_taken,
            "message": f"Evaluated {len(rules)} rules, {len(matched_rules)} matched"
        }

    async def evaluate_rule(
        self,
        email: EmailRecord,
        rule: EmailRule
    ) -> Dict[str, Any]:
        """
        Evaluate a single rule against an email.

        Args:
            email: The email to evaluate
            rule: The rule to check

        Returns:
            Dict with matches (bool), confidence (int), reason (str)
        """
        messages = [{
            "role": "user",
            "content": f"""Evaluate if this email matches the following rule condition:

RULE CONDITION:
"{rule.match_prompt}"

EMAIL DETAILS:
From: {email.sender_name or ""} <{email.sender_email}>
Subject: {email.subject}
Preview: {email.body_preview[:500]}

---

Respond with valid JSON:
{{
    "matches": true or false,
    "confidence": 0-100,
    "reason": "brief explanation"
}}

Only output the JSON, nothing else."""
        }]

        try:
            result = self.call_claude_structured(
                messages,
                response_schema={
                    "type": "object",
                    "properties": {
                        "matches": {"type": "boolean"},
                        "confidence": {"type": "integer", "minimum": 0, "maximum": 100},
                        "reason": {"type": "string"}
                    },
                    "required": ["matches", "confidence", "reason"]
                },
                max_tokens=200
            )

            return result

        except Exception as e:
            logger.error(f"Rule evaluation failed for rule {rule.id}: {e}")
            return {
                "matches": False,
                "confidence": 0,
                "reason": f"Evaluation error: {str(e)}"
            }

    async def evaluate_single_email(
        self,
        rule: EmailRule,
        email: EmailRecord,
        min_confidence: int = 50
    ) -> Dict[str, Any]:
        """
        Evaluate a single email against a rule - used for streaming progress.

        Args:
            rule: The rule to check
            email: The email to evaluate
            min_confidence: Minimum confidence to consider a match

        Returns:
            Dict with matches (bool), confidence, reason
        """
        result = await self.evaluate_rule(email, rule)

        if result["matches"] and result["confidence"] >= min_confidence:
            self.db.increment_email_rule_hit(rule.id)
            return {
                "matches": True,
                "confidence": result["confidence"],
                "reason": result["reason"]
            }

        return {
            "matches": False,
            "confidence": result["confidence"],
            "reason": result["reason"]
        }

    async def evaluate_all_rules(
        self,
        email: EmailRecord,
        min_confidence: int = 50
    ) -> List[Tuple[EmailRule, Dict[str, Any]]]:
        """
        Evaluate email against all active rules and return matches.

        Args:
            email: The email to evaluate
            min_confidence: Minimum confidence threshold for a match

        Returns:
            List of tuples (rule, match_result) for matching rules
        """
        rules = self.db.get_active_email_rules()
        matches = []

        for rule in rules:
            result = await self.evaluate_rule(email, rule)

            if result["matches"] and result["confidence"] >= min_confidence:
                matches.append((rule, result))
                self.db.increment_email_rule_hit(rule.id)

                # Log the match
                self.log_action(
                    "rule_matched",
                    email_id=email.id,
                    details={
                        "rule_id": rule.id,
                        "rule_name": rule.name,
                        "confidence": result["confidence"],
                        "action": rule.action.value,
                        "action_value": rule.action_value
                    }
                )

                if rule.stop_processing:
                    break

        return matches

    async def get_matching_folder(self, email: EmailRecord) -> Optional[str]:
        """
        Get the folder to move an email to based on matching rules.

        Args:
            email: The email to check

        Returns:
            Folder name if a matching MOVE_TO_FOLDER rule exists, None otherwise
        """
        rules = self.db.get_active_email_rules()

        # Filter to only folder move rules
        folder_rules = [r for r in rules if r.action == RuleAction.MOVE_TO_FOLDER]

        for rule in folder_rules:
            result = await self.evaluate_rule(email, rule)

            if result["matches"] and result["confidence"] >= 60:
                self.db.increment_email_rule_hit(rule.id)

                self.log_action(
                    "folder_rule_matched",
                    email_id=email.id,
                    details={
                        "rule_id": rule.id,
                        "rule_name": rule.name,
                        "folder": rule.action_value,
                        "confidence": result["confidence"]
                    }
                )

                return rule.action_value

        return None

    async def test_rule(
        self,
        rule: EmailRule,
        limit: int = 10
    ) -> Dict[str, Any]:
        """
        Test a rule against recent emails to see what it would match.

        Args:
            rule: The rule to test
            limit: Maximum number of emails to test

        Returns:
            Dict with test results
        """
        recent_emails = self.db.get_recent_emails(hours=168, limit=limit)

        matches = []
        non_matches = []

        for email in recent_emails:
            result = await self.evaluate_rule(email, rule)

            email_summary = {
                "id": email.id,
                "subject": email.subject[:50],
                "sender": email.sender_email,
                "confidence": result["confidence"],
                "reason": result["reason"]
            }

            if result["matches"] and result["confidence"] >= 50:
                matches.append(email_summary)
            else:
                non_matches.append(email_summary)

        return {
            "rule_name": rule.name,
            "match_prompt": rule.match_prompt,
            "total_tested": len(recent_emails),
            "matches": matches,
            "non_matches": non_matches[:5],  # Limit non-matches for brevity
            "match_rate": f"{len(matches)}/{len(recent_emails)}"
        }

    async def suggest_rules_for_email(self, email: EmailRecord) -> List[Dict[str, Any]]:
        """
        Suggest potential rules based on an email's characteristics.

        Args:
            email: The email to analyze

        Returns:
            List of suggested rules
        """
        messages = [{
            "role": "user",
            "content": f"""Analyze this email and suggest routing rules that could apply to similar emails:

From: {email.sender_name or ""} <{email.sender_email}>
Subject: {email.subject}
Preview: {email.body_preview[:500]}

Suggest 1-3 rules in this format:
1. Rule Name: <name>
   Match Condition: <natural language condition>
   Suggested Action: <move_to_folder/archive/add_label>
   Folder/Label: <destination if applicable>

Focus on patterns that would apply to multiple similar emails, not just this one.
Only output the suggestions, no other text."""
        }]

        try:
            response = self.call_claude(messages, max_tokens=500, temperature=0.5)

            self.log_action(
                "rules_suggested",
                email_id=email.id,
                details={"suggestions": response[:200]}
            )

            return [{"suggestions": response}]

        except Exception as e:
            logger.error(f"Rule suggestion failed: {e}")
            return []

    async def run_rule_on_emails(
        self,
        rule: EmailRule,
        emails: List[EmailRecord],
        dry_run: bool = False
    ) -> Dict[str, Any]:
        """
        Execute a rule against a list of emails, optionally applying the action.

        Args:
            rule: The rule to run
            emails: List of emails to check
            dry_run: If True, only evaluate without applying actions

        Returns:
            Dict with results of the run
        """
        from ..integrations.mcp_email import EmailClient

        matched = []
        processed = []
        errors = []

        for email in emails:
            try:
                result = await self.evaluate_rule(email, rule)

                if result["matches"] and result["confidence"] >= 50:
                    matched.append({
                        "email_id": email.id,
                        "subject": email.subject[:60],
                        "sender": email.sender_email,
                        "confidence": result["confidence"],
                        "reason": result["reason"]
                    })

                    if not dry_run:
                        # Apply the action
                        action_result = await self._execute_rule_action(email, rule)
                        if action_result["success"]:
                            processed.append({
                                "email_id": email.id,
                                "action": rule.action.value,
                                "action_value": rule.action_value
                            })
                            # Record the hit
                            self.db.increment_email_rule_hit(rule.id)
                        else:
                            errors.append({
                                "email_id": email.id,
                                "error": action_result.get("error", "Unknown error")
                            })

            except Exception as e:
                logger.error(f"Error processing email {email.id} with rule {rule.id}: {e}")
                errors.append({
                    "email_id": email.id,
                    "error": str(e)
                })

        return {
            "rule_id": rule.id,
            "rule_name": rule.name,
            "total_evaluated": len(emails),
            "matched": len(matched),
            "processed": len(processed),
            "errors": len(errors),
            "dry_run": dry_run,
            "matches": matched,
            "processed_emails": processed,
            "error_details": errors
        }

    async def _execute_rule_action(
        self,
        email: EmailRecord,
        rule: EmailRule
    ) -> Dict[str, Any]:
        """
        Execute the action for a rule on an email.

        Args:
            email: The email to act on
            rule: The rule whose action to execute

        Returns:
            Dict with success status and details
        """
        from ..integrations.mcp_email import EmailClient

        try:
            email_client = EmailClient()

            if rule.action == RuleAction.MOVE_TO_FOLDER:
                if not rule.action_value:
                    return {"success": False, "error": "No folder specified"}

                # Move the email to the specified folder
                success = email_client.move_to_folder(
                    email.message_id,
                    rule.action_value,
                    email.mailbox
                )

                if success:
                    # Update email state in database so it doesn't show in main list
                    email.state = "archived"
                    self.db.save_email(email)

                    self.log_action(
                        "rule_action_executed",
                        email_id=email.id,
                        details={
                            "rule_id": rule.id,
                            "action": "move_to_folder",
                            "folder": rule.action_value
                        }
                    )
                    return {"success": True, "action": "moved", "folder": rule.action_value}
                else:
                    return {"success": False, "error": "Failed to move email"}

            elif rule.action == RuleAction.ARCHIVE:
                # Move to Archive folder
                success = email_client.move_to_folder(
                    email.message_id,
                    "Archive",
                    email.mailbox
                )

                if success:
                    # Update email state in database
                    email.state = "archived"
                    self.db.save_email(email)

                    self.log_action(
                        "rule_action_executed",
                        email_id=email.id,
                        details={
                            "rule_id": rule.id,
                            "action": "archive"
                        }
                    )
                    return {"success": True, "action": "archived"}
                else:
                    return {"success": False, "error": "Failed to archive email"}

            elif rule.action == RuleAction.SET_PRIORITY:
                # Update email priority in database
                try:
                    priority = int(rule.action_value) if rule.action_value else 3
                    email.priority = max(1, min(5, priority))
                    self.db.save_email(email)
                    return {"success": True, "action": "priority_set", "priority": email.priority}
                except ValueError:
                    return {"success": False, "error": "Invalid priority value"}

            else:
                return {"success": False, "error": f"Action {rule.action.value} not implemented yet"}

        except Exception as e:
            logger.error(f"Error executing rule action: {e}")
            return {"success": False, "error": str(e)}
