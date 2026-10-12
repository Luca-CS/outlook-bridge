"""Exclusions and redaction (spec §6.3).

Excluded items are never returned, even when asked for by ID. Redaction replaces every match
of `redact_patterns` in subjects, snippets, bodies and locations.
"""

from __future__ import annotations

import re
from typing import TypeVar

from outlook_bridge.config import MAX_BODY_CHARS_LIMIT, Config
from outlook_bridge.models import Event, MailDetail, MailSummary

REDACTED = "[masqué]"
SNIPPET_MAX = 200
# Extra body characters fetched from Outlook so that a pattern cut by truncation is still
# redacted: bodies are redacted first, then clipped (`clip_body`).
REDACT_MARGIN = 200


def body_limit(requested: int | None, default: int) -> int:
    """`max_chars` asked by the caller, defaulting to the config, capped at the hard limit."""
    return max(1, min(requested or default, MAX_BODY_CHARS_LIMIT))


def clip_body(body: str, truncated: bool, max_chars: int) -> dict[str, str | bool]:
    """`model_copy` update truncating an already-redacted body to `max_chars`."""
    return {"body_text": body[:max_chars], "body_truncated": truncated or len(body) > max_chars}


M = TypeVar("M", MailSummary, MailDetail)


def _fold(values: list[str]) -> list[str]:
    return [v.strip().casefold() for v in values if v.strip()]


class Policy:
    def __init__(self, config: Config) -> None:
        self._accounts = set(_fold(config.accounts_allowed))
        self._folders = set(_fold(config.folders_excluded))
        self._senders = [s.lstrip("@") for s in _fold(config.senders_excluded)]
        self._keywords = _fold(config.subject_keywords_excluded)
        self._patterns = [re.compile(p) for p in config.redact_patterns]

    # ------------------------------------------------------------------ exclusions

    def account_allowed(self, *names: str) -> bool:
        """`accounts_allowed` empty = all. Matches a display name or an SMTP address."""
        if not self._accounts:
            return True
        return any(name and name.casefold() in self._accounts for name in names)

    def folder_excluded(self, path: str) -> bool:
        """Excluded if any segment of the path matches (subfolders of Junk are excluded too)."""
        return any(part.casefold() in self._folders for part in path.split("/") if part)

    def sender_excluded(self, email: str) -> bool:
        email = email.strip().casefold()
        if not email:
            return False
        domain = email.rsplit("@", 1)[-1]
        for rule in self._senders:
            if "@" in rule:
                if email == rule:
                    return True
            elif domain == rule or domain.endswith("." + rule):
                return True
        return False

    def subject_excluded(self, subject: str) -> bool:
        folded = subject.casefold()
        return any(keyword in folded for keyword in self._keywords)

    def mail_visible(self, mail: MailSummary) -> bool:
        return (
            self.account_allowed(mail.account)
            and not self.folder_excluded(mail.folder)
            and not self.sender_excluded(mail.sender_email)
            and not self.subject_excluded(mail.subject)
        )

    def event_visible(self, event: Event) -> bool:
        return (
            self.account_allowed(event.account)
            and not self.folder_excluded(event.calendar)
            and not self.subject_excluded(event.subject)
        )

    # ------------------------------------------------------------------ redaction

    def redact(self, text: str) -> str:
        for pattern in self._patterns:
            text = pattern.sub(REDACTED, text)
        return text

    def redact_mail(self, mail: M) -> M:
        update = {
            "subject": self.redact(mail.subject),
            "snippet": self.redact(mail.snippet)[:SNIPPET_MAX],
        }
        if isinstance(mail, MailDetail):
            update["body_text"] = self.redact(mail.body_text)
        return mail.model_copy(update=update)

    def redact_event(self, event: Event) -> Event:
        update: dict[str, str | None] = {
            "subject": self.redact(event.subject),
            "location": self.redact(event.location),
        }
        if event.body_text is not None:
            update["body_text"] = self.redact(event.body_text)
        return event.model_copy(update=update)
