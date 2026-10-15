"""OutlookGateway: the boundary between the core and Outlook.

Only `com_gateway.py` implements it against COM; tests use `tests/fakes.py::FakeGateway`.
Gateways return pydantic models with tz-aware dates; exclusion and redaction stay in the core.

The `accept` callbacks are pure-Python predicates (exclusion policy). The gateway applies them
while scanning so that `limit` counts visible items only. They may run on the COM worker
thread, so they must not touch COM or shared mutable state.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime
from typing import Protocol

from outlook_bridge.models import Account, Event, Folder, MailDetail, MailSummary, Status

AcceptMail = Callable[[MailSummary], bool]
AcceptEvent = Callable[[Event], bool]


@dataclass(frozen=True)
class MailQuery:
    folder: str = "Inbox"
    account: str | None = None
    query: str | None = None  # substring of subject or body
    since: datetime | None = None  # aware
    until: datetime | None = None  # aware, exclusive
    sender: str | None = None  # substring of sender address or name
    unread_only: bool = False
    limit: int = 25


class OutlookGateway(Protocol):
    def status(self) -> Status: ...

    def list_accounts(self) -> list[Account]: ...

    def list_folders(self, account: str | None = None) -> list[Folder]: ...

    def search_mail(self, query: MailQuery, accept: AcceptMail) -> list[MailSummary]:
        """Newest first, at most `query.limit` accepted items."""
        ...

    def get_mail(self, entry_id: str, store_id: str, max_chars: int) -> MailDetail | None:
        """None if the item does not exist or is not a mail."""
        ...

    def get_thread(self, conversation_id: str, limit: int, accept: AcceptMail) -> list[MailSummary]:
        """Chronological order, the `limit` most recent accepted items."""
        ...

    def list_events(
        self, start: datetime, end: datetime, calendar: str | None, accept: AcceptEvent
    ) -> list[Event]:
        """Events overlapping [start, end), recurrences expanded, sorted by start."""
        ...

    def get_event(
        self,
        entry_id: str,
        store_id: str,
        max_chars: int,
        occurrence_start: datetime | None = None,
    ) -> Event | None:
        """`occurrence_start` selects one occurrence of a recurring series (same EntryID)."""
        ...

    def close(self) -> None:
        """Releases the connection (COM worker thread). Idempotent; safe if never used."""
        ...
