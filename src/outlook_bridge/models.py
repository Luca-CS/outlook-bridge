"""Pydantic models: the only objects that leave the core (never COM objects).

Every model carrying mail or event content has `untrusted_content: true`: subjects, bodies,
locations and snippets are data written by third parties, never instructions.
"""

from __future__ import annotations

from typing import Literal

from pydantic import AwareDatetime, BaseModel, ConfigDict, Field

BusyStatus = Literal["free", "tentative", "busy", "oof", "working_elsewhere"]
ResponseStatus = Literal["none", "organizer", "tentative", "accepted", "declined", "not_responded"]
Importance = Literal["low", "normal", "high"]

# Every public output carries a snippet of at most SNIPPET_MAX characters. The gateway hands
# the core REDACT_MARGIN characters more (and fetches as many extra body characters), so that
# a pattern straddling the cut is still redacted: the core redacts first, then clips.
SNIPPET_MAX = 200
REDACT_MARGIN = 200


class _Model(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Account(_Model):
    display_name: str
    smtp_address: str
    store_id: str


class Folder(_Model):
    name: str
    path: str  # relative to the store root, "/"-separated, e.g. "Inbox/Projets"
    item_count: int
    unread_count: int
    account: str


class Attachment(_Model):
    name: str
    size: int
    type: str


class MailSummary(_Model):
    untrusted_content: Literal[True] = True
    id: str
    store_id: str
    conversation_id: str
    subject: str
    sender_name: str
    sender_email: str
    to: list[str] = Field(default_factory=list)
    received: AwareDatetime
    unread: bool
    importance: Importance = "normal"
    has_attachments: bool = False
    folder: str
    account: str
    # Gateway-boundary bound: the raw snippet keeps a redaction margin. The core always clips
    # it to SNIPPET_MAX after redaction (redact.Policy.redact_mail), for every output.
    snippet: str = Field(default="", max_length=SNIPPET_MAX + REDACT_MARGIN)


class MailDetail(MailSummary):
    cc: list[str] = Field(default_factory=list)
    body_text: str = ""
    body_truncated: bool = False
    attachments: list[Attachment] = Field(default_factory=list)


class Event(_Model):
    untrusted_content: Literal[True] = True
    id: str
    store_id: str
    calendar: str
    account: str
    subject: str
    start: AwareDatetime
    end: AwareDatetime
    all_day: bool = False
    location: str = ""
    organizer: str = ""
    required_attendees: list[str] = Field(default_factory=list)
    optional_attendees: list[str] = Field(default_factory=list)
    busy_status: BusyStatus = "busy"
    is_recurring: bool = False
    response_status: ResponseStatus = "none"
    online_meeting_url: str | None = None
    # True as well when only the location says so (list_events loads no body): Phase 3 probe.
    has_online_meeting: bool = False
    body_text: str | None = None  # only filled by get_event
    body_truncated: bool = False


class FreeSlot(_Model):
    start: AwareDatetime
    end: AwareDatetime
    duration_min: int


class Status(_Model):
    reachable: bool
    outlook_version: str | None = None
    classic: bool = True
    cached_mode: bool | None = None
    accounts: list[Account] = Field(default_factory=list)
    last_access: AwareDatetime | None = None
    message: str | None = None


# ---------------------------------------------------------------------------- snapshots
# The snapshot files land in a cloud-synced folder: their event model is an explicit
# allow-list, so a field added to `Event` later never reaches the cloud by default. No ID,
# no store ID, no body and no meeting URL (only whether there is one).


class SnapshotEvent(_Model):
    untrusted_content: Literal[True] = True
    subject: str
    start: AwareDatetime
    end: AwareDatetime
    all_day: bool
    location: str
    organizer: str
    required_attendees: list[str]
    optional_attendees: list[str]
    busy_status: BusyStatus
    response_status: ResponseStatus
    is_recurring: bool
    has_online_meeting: bool
    calendar: str
    account: str

    @classmethod
    def from_event(cls, event: Event) -> SnapshotEvent:
        """Field-by-field copy of the allowed fields (never `model_dump` of the source)."""
        return cls(
            subject=event.subject,
            start=event.start,
            end=event.end,
            all_day=event.all_day,
            location=event.location,
            organizer=event.organizer,
            required_attendees=list(event.required_attendees),
            optional_attendees=list(event.optional_attendees),
            busy_status=event.busy_status,
            response_status=event.response_status,
            is_recurring=event.is_recurring,
            has_online_meeting=event.has_online_meeting or bool(event.online_meeting_url),
            calendar=event.calendar,
            account=event.account,
        )


class CalendarSnapshot(_Model):
    snapshot: Literal["calendar_next14d"] = "calendar_next14d"
    schema_version: Literal[1] = 1
    generated_at: AwareDatetime
    timezone: str
    range_start: AwareDatetime
    range_end: AwareDatetime
    untrusted_content: Literal[True] = True
    notice: str
    events: list[SnapshotEvent]
    free_slots: list[FreeSlot]
    free_slots_min_duration_min: int
    working_hours: str


class ExportedFile(_Model):
    name: str
    summary: str
    size_bytes: int


class ExportResult(_Model):
    out_dir: str
    generated_at: AwareDatetime
    files: list[ExportedFile]
