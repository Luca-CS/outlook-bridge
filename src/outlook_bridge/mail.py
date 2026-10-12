"""Mail: search, detail, thread. Validation, exclusions and redaction live here."""

from __future__ import annotations

from datetime import datetime, tzinfo

from outlook_bridge.config import MAX_BODY_CHARS_LIMIT, Config
from outlook_bridge.errors import InvalidRequest, NotFound
from outlook_bridge.gateway import MailQuery, OutlookGateway
from outlook_bridge.models import MailDetail, MailSummary
from outlook_bridge.redact import Policy
from outlook_bridge.timeutil import dasl_datetime, ensure_aware

MAX_LIMIT = 100
MAX_THREAD_LIMIT = 100
# Extra body characters fetched so that a pattern cut by truncation is still redacted.
REDACT_MARGIN = 200

DASL_RECEIVED = "urn:schemas:httpmail:datereceived"
DASL_READ = "urn:schemas:httpmail:read"
DASL_SUBJECT = "urn:schemas:httpmail:subject"
DASL_BODY = "urn:schemas:httpmail:textdescription"
DASL_FROM_NAME = "urn:schemas:httpmail:fromname"
DASL_FROM_EMAIL = "urn:schemas:httpmail:fromemail"
DASL_SENDER_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"


def _quote(value: str) -> str:
    return value.replace("'", "''")


def build_dasl_filter(query: MailQuery, tz: tzinfo) -> str:
    """`@SQL=` filter for Folder.GetTable. Dates in UTC (Phase 0); empty string = no filter."""
    parts: list[str] = []
    if query.since is not None:
        parts.append(f"\"{DASL_RECEIVED}\" >= '{dasl_datetime(query.since, tz)}'")
    if query.until is not None:
        parts.append(f"\"{DASL_RECEIVED}\" < '{dasl_datetime(query.until, tz)}'")
    if query.unread_only:
        parts.append(f'"{DASL_READ}" = 0')
    if query.sender:
        s = _quote(query.sender.strip())
        fields = (DASL_SENDER_SMTP, DASL_FROM_NAME, DASL_FROM_EMAIL)
        parts.append("(" + " OR ".join(f"\"{f}\" LIKE '%{s}%'" for f in fields) + ")")
    if query.query:
        s = _quote(query.query.strip())
        fields = (DASL_SUBJECT, DASL_BODY)
        parts.append("(" + " OR ".join(f"\"{f}\" LIKE '%{s}%'" for f in fields) + ")")
    return "@SQL=" + " AND ".join(parts) if parts else ""


class MailService:
    def __init__(self, gateway: OutlookGateway, config: Config, policy: Policy) -> None:
        self._gw = gateway
        self._config = config
        self._policy = policy

    def search_mail(
        self,
        query: str | None = None,
        folder: str = "Inbox",
        account: str | None = None,
        since: datetime | None = None,
        until: datetime | None = None,
        sender: str | None = None,
        unread_only: bool = False,
        limit: int = 25,
    ) -> list[MailSummary]:
        if limit < 1:
            raise InvalidRequest("limit doit être ≥ 1.")
        limit = min(limit, MAX_LIMIT)
        tz = self._config.tz
        since = ensure_aware(since, tz) if since else None
        until = ensure_aware(until, tz) if until else None
        if since and until and since >= until:
            raise InvalidRequest("since doit précéder until.")
        if self._policy.folder_excluded(folder):
            raise NotFound("Dossier")
        mail_query = MailQuery(
            folder=folder,
            account=account,
            query=query or None,
            since=since,
            until=until,
            sender=sender or None,
            unread_only=unread_only,
            limit=limit,
        )
        found = self._gw.search_mail(mail_query, self._policy.mail_visible)
        return [self._policy.redact_mail(m) for m in found[:limit]]

    def get_mail(self, id: str, store_id: str, max_chars: int | None = None) -> MailDetail:
        max_chars = max(1, min(max_chars or self._config.max_body_chars, MAX_BODY_CHARS_LIMIT))
        mail = self._gw.get_mail(id, store_id, max_chars + REDACT_MARGIN)
        if mail is None or not self._policy.mail_visible(mail):
            raise NotFound("Mail")
        mail = self._policy.redact_mail(mail)
        body = mail.body_text
        truncated = mail.body_truncated or len(body) > max_chars
        return mail.model_copy(update={"body_text": body[:max_chars], "body_truncated": truncated})

    def get_thread(self, conversation_id: str, limit: int = 20) -> list[MailSummary]:
        if limit < 1:
            raise InvalidRequest("limit doit être ≥ 1.")
        limit = min(limit, MAX_THREAD_LIMIT)
        found = self._gw.get_thread(conversation_id, limit, self._policy.mail_visible)
        return [self._policy.redact_mail(m) for m in found[-limit:]]
