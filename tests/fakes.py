"""FakeGateway and synthetic fixtures. 100 % invented data: never derive from a real mailbox."""

from __future__ import annotations

import json
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from outlook_bridge.errors import NewOutlookDetected
from outlook_bridge.gateway import AcceptEvent, AcceptMail, MailQuery
from outlook_bridge.models import Account, Event, Folder, MailDetail, MailSummary, Status

PARIS = ZoneInfo("Europe/Paris")
STORE = "STORE-1"
ACCOUNT = "alice@example.com"
NOW = datetime(2026, 10, 12, 12, 0, tzinfo=PARIS)  # a Monday


def make_mail(n: int, **overrides) -> MailDetail:
    data = {
        "id": f"MAIL-{n}",
        "store_id": STORE,
        "conversation_id": f"CONV-{n}",
        "subject": f"Réunion test {n}",
        "sender_name": "Bob Exemple",
        "sender_email": "bob@example.com",
        "to": ["Alice Exemple"],
        "received": NOW - timedelta(hours=n),
        "unread": n % 2 == 0,
        "folder": "Inbox",
        "account": ACCOUNT,
        "snippet": f"Corps synthétique {n}",
        "body_text": f"Corps synthétique {n}.",
    }
    data.update(overrides)
    return MailDetail(**data)


def make_event(n: int, start: datetime, minutes: int = 60, **overrides) -> Event:
    data = {
        "id": f"EVT-{n}",
        "store_id": STORE,
        "calendar": "Calendar",
        "account": ACCOUNT,
        "subject": f"Point projet {n}",
        "start": start,
        "end": start + timedelta(minutes=minutes),
        "body_text": f"Ordre du jour synthétique {n}.",
    }
    data.update(overrides)
    return Event(**data)


def _summary(mail: MailDetail) -> MailSummary:
    return MailSummary(**mail.model_dump(include=set(MailSummary.model_fields)))


class MemorySink:
    """AuditSink keeping the records in memory, JSON round-tripped as JsonlSink writes them."""

    def __init__(self) -> None:
        self.records: list[dict] = []
        self.whens: list[datetime] = []

    def write(self, record: dict[str, object], when: datetime) -> None:
        self.records.append(json.loads(json.dumps(record, ensure_ascii=False)))
        self.whens.append(when)


class StepTimer:
    """Deterministic perf_counter: every call advances by `step_s` seconds."""

    def __init__(self, step_s: float = 0.25) -> None:
        self.step_s = step_s
        self.value = 0.0

    def __call__(self) -> float:
        current = self.value
        self.value += self.step_s
        return current


class FakeGateway:
    """In-memory OutlookGateway with the same contract as ComGateway.

    `errors` maps a method name to the exception that method raises. `status` never raises:
    like ComGateway, it reports the failure as an unreachable Status.
    """

    def __init__(
        self,
        mails: list[MailDetail] | None = None,
        events: list[Event] | None = None,
        folders: list[Folder] | None = None,
        errors: dict[str, BaseException] | None = None,
    ) -> None:
        self.mails = mails or []
        self.events = events or []
        self.folders = folders or []
        self.errors = errors or {}
        self.calls: list[tuple] = []
        self.closed = False

    def _fail(self, method: str) -> None:
        if (exc := self.errors.get(method)) is not None:
            raise exc

    def close(self) -> None:
        self.closed = True

    def status(self) -> Status:
        if (exc := self.errors.get("status")) is not None:
            return Status(
                reachable=False, classic=not isinstance(exc, NewOutlookDetected), message=str(exc)
            )
        return Status(reachable=True, outlook_version="16.0", accounts=self._accounts())

    def list_accounts(self) -> list[Account]:
        self._fail("list_accounts")
        return self._accounts()

    def _accounts(self) -> list[Account]:
        return [
            Account(display_name=ACCOUNT, smtp_address=ACCOUNT, store_id=STORE),
            Account(display_name="perso@example.org", smtp_address="", store_id="STORE-2"),
        ]

    def list_folders(self, account: str | None = None) -> list[Folder]:
        self._fail("list_folders")
        return [f for f in self.folders if account is None or f.account == account]

    def search_mail(self, query: MailQuery, accept: AcceptMail) -> list[MailSummary]:
        self._fail("search_mail")
        self.calls.append(("search_mail", query))
        folder = query.folder.casefold()
        found: list[MailSummary] = []
        for mail in sorted(self.mails, key=lambda m: m.received, reverse=True):
            if len(found) >= query.limit:
                break
            if mail.folder.casefold() != folder:
                continue
            if query.account and mail.account != query.account:
                continue
            if query.since and mail.received < query.since:
                continue
            if query.until and mail.received >= query.until:
                continue
            if query.unread_only and not mail.unread:
                continue
            if query.sender and not any(
                query.sender.casefold() in s.casefold()
                for s in (mail.sender_email, mail.sender_name)
            ):
                continue
            if query.query and not any(
                query.query.casefold() in s.casefold() for s in (mail.subject, mail.body_text)
            ):
                continue
            summary = _summary(mail)
            if accept(summary):
                found.append(summary)
        return found

    def get_mail(self, entry_id: str, store_id: str, max_chars: int) -> MailDetail | None:
        self._fail("get_mail")
        self.calls.append(("get_mail", entry_id, store_id, max_chars))
        for mail in self.mails:
            if mail.id == entry_id and mail.store_id == store_id:
                body = mail.body_text
                return mail.model_copy(
                    update={"body_text": body[:max_chars], "body_truncated": len(body) > max_chars}
                )
        return None

    def get_thread(self, conversation_id: str, limit: int, accept: AcceptMail) -> list[MailSummary]:
        self._fail("get_thread")
        found = [
            _summary(m)
            for m in sorted(self.mails, key=lambda m: m.received)
            if m.conversation_id == conversation_id
        ]
        return [m for m in found if accept(m)][-limit:]

    def list_events(
        self, start: datetime, end: datetime, calendar: str | None, accept: AcceptEvent
    ) -> list[Event]:
        self._fail("list_events")
        self.calls.append(("list_events", start, end, calendar))
        return sorted(
            (
                e.model_copy(update={"body_text": None})
                for e in self.events
                if e.start < end
                and e.end > start
                and (calendar is None or e.calendar.casefold() == calendar.casefold())
                and accept(e)
            ),
            key=lambda e: e.start,
        )

    def get_event(
        self,
        entry_id: str,
        store_id: str,
        max_chars: int,
        occurrence_start: datetime | None = None,
    ) -> Event | None:
        self._fail("get_event")
        for event in self.events:
            if event.id != entry_id or event.store_id != store_id:
                continue
            if occurrence_start is not None and event.start != occurrence_start:
                continue
            body = event.body_text or ""
            return event.model_copy(
                update={"body_text": body[:max_chars], "body_truncated": len(body) > max_chars}
            )
        return None
