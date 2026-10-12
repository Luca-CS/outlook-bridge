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
    snippet: str = Field(default="", max_length=200)


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
