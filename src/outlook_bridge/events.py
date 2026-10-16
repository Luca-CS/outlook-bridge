"""Calendar: events, Jet range filter, free slots."""

from __future__ import annotations

import re
from collections.abc import Iterable
from datetime import datetime, time, timedelta, tzinfo

from outlook_bridge.audit import Auditor, audited
from outlook_bridge.config import Config
from outlook_bridge.errors import InvalidRequest, NotFound
from outlook_bridge.gateway import OutlookGateway
from outlook_bridge.models import REDACT_MARGIN, Event, FreeSlot
from outlook_bridge.redact import Policy, body_limit, clip_body
from outlook_bridge.timeutil import ensure_aware, jet_datetime, parse_working_hours

MAX_RANGE_DAYS = 62

_MEETING_URL = re.compile(
    r"https://(?:"
    r"teams\.microsoft\.com/l/meetup-join/"
    r"|teams\.live\.com/meet/"
    r"|(?:[\w-]+\.)?zoom\.us/(?:j|my|w)/"
    r"|meet\.google\.com/"
    r"|[\w-]+\.webex\.com/"
    r")[^\s<>\"']+",
    re.IGNORECASE,
)


# Locations Outlook writes for a Teams meeting ("Réunion Microsoft Teams", "Microsoft Teams
# Meeting", "Salle A; Réunion Microsoft Teams"). The Phase 3 probe found them on exactly the
# events that carry a Teams link, so list_events can flag them without loading any body.
ONLINE_LOCATION_MARKERS = ("microsoft teams",)


def online_location(location: str | None) -> bool:
    """Whether the location names an online meeting service."""
    folded = (location or "").casefold()
    return any(marker in folded for marker in ONLINE_LOCATION_MARKERS)


def find_meeting_url(*texts: str | None) -> str | None:
    """First Teams / Zoom / Meet / Webex link found in the given texts."""
    for text in texts:
        if text and (match := _MEETING_URL.search(text)):
            return match.group(0).rstrip(".,;)>]")
    return None


def build_jet_range(start: datetime, end: datetime, tz: tzinfo) -> str:
    """Jet filter for events overlapping [start, end). `dd/MM/yyyy HH:mm` only (Phase 0)."""
    return f"[Start] < '{jet_datetime(end, tz)}' AND [End] > '{jet_datetime(start, tz)}'"


def _check_range(start: datetime, end: datetime) -> None:
    if end <= start:
        raise InvalidRequest("La fin de la plage doit suivre son début.")
    if end - start > timedelta(days=MAX_RANGE_DAYS):
        raise InvalidRequest(f"Plage trop large : {MAX_RANGE_DAYS} jours maximum.")


def _blocking(event: Event, treat_tentative_as_busy: bool) -> bool:
    if event.busy_status in ("busy", "oof"):
        return True
    return event.busy_status == "tentative" and treat_tentative_as_busy


def compute_free_slots(
    events: Iterable[Event],
    start: datetime,
    end: datetime,
    *,
    min_duration_min: int,
    working_hours: tuple[time, time],
    include_weekends: bool,
    treat_tentative_as_busy: bool,
    tz: tzinfo,
) -> list[FreeSlot]:
    """Free intervals inside working hours, outside blocking events. Pure function."""
    start, end = ensure_aware(start, tz), ensure_aware(end, tz)
    busy = sorted(
        (max(e.start, start), min(e.end, end))
        for e in events
        if _blocking(e, treat_tentative_as_busy) and e.end > start and e.start < end
    )
    merged: list[tuple[datetime, datetime]] = []
    for b_start, b_end in busy:
        if merged and b_start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], b_end))
        else:
            merged.append((b_start, b_end))

    min_duration = timedelta(minutes=min_duration_min)
    slots: list[FreeSlot] = []

    def emit(slot_start: datetime, slot_end: datetime) -> None:
        if slot_end - slot_start >= min_duration:
            slots.append(
                FreeSlot(
                    start=slot_start.astimezone(tz),
                    end=slot_end.astimezone(tz),
                    duration_min=int((slot_end - slot_start).total_seconds() // 60),
                )
            )

    day, last_day = start.astimezone(tz).date(), end.astimezone(tz).date()
    while day <= last_day:
        if include_weekends or day.weekday() < 5:
            low = max(datetime.combine(day, working_hours[0], tzinfo=tz), start)
            high = min(datetime.combine(day, working_hours[1], tzinfo=tz), end)
            cursor = low
            for b_start, b_end in merged:
                if b_end <= cursor:
                    continue
                if b_start >= high:
                    break
                if b_start > cursor:
                    emit(cursor, b_start)
                cursor = max(cursor, b_end)
            if cursor < high:
                emit(cursor, high)
        day += timedelta(days=1)
    return slots


class CalendarService:
    def __init__(
        self,
        gateway: OutlookGateway,
        config: Config,
        policy: Policy,
        audit: Auditor | None = None,
    ) -> None:
        self._audit = audit if audit is not None else Auditor.disabled()
        self._gw = gateway
        self._config = config
        self._policy = policy

    @audited
    def list_events(
        self, start: datetime, end: datetime, calendar: str | None = None
    ) -> list[Event]:
        tz = self._config.tz
        start, end = ensure_aware(start, tz), ensure_aware(end, tz)
        _check_range(start, end)
        if calendar and self._policy.folder_excluded(calendar):
            # Same answer as an unknown calendar (the gateway returns no event): an error
            # here would reveal which names are in `folders_excluded`.
            return []
        events = self._gw.list_events(start, end, calendar, self._policy.event_visible)
        return [self._policy.redact_event(e) for e in sorted(events, key=lambda e: e.start)]

    @audited
    def get_event(
        self,
        id: str,
        store_id: str,
        max_chars: int | None = None,
        start: datetime | None = None,
    ) -> Event:
        tz = self._config.tz
        max_chars = body_limit(max_chars, self._config.max_body_chars)
        occurrence = ensure_aware(start, tz) if start else None
        event = self._gw.get_event(id, store_id, max_chars + REDACT_MARGIN, occurrence)
        if event is None or not self._policy.event_visible(event):
            raise NotFound("Événement")
        event = self._policy.redact_event(event)
        return event.model_copy(
            update=clip_body(event.body_text or "", event.body_truncated, max_chars)
        )

    @audited
    def find_free_slots(
        self,
        start: datetime,
        end: datetime,
        min_duration_min: int = 30,
        working_hours: str | None = None,
        include_weekends: bool = False,
        treat_tentative_as_busy: bool = True,
    ) -> list[FreeSlot]:
        tz = self._config.tz
        start, end = ensure_aware(start, tz), ensure_aware(end, tz)
        _check_range(start, end)
        if min_duration_min < 1:
            raise InvalidRequest("min_duration_min doit valoir au moins 1.")
        hours = parse_working_hours(working_hours or self._config.working_hours)
        # Busy time counts even for events hidden by keyword: only their time slot is used,
        # never their content. Excluded accounts and calendars are ignored entirely.
        events = self._gw.list_events(start, end, None, self._time_visible)
        return compute_free_slots(
            events,
            start,
            end,
            min_duration_min=min_duration_min,
            working_hours=hours,
            include_weekends=include_weekends,
            treat_tentative_as_busy=treat_tentative_as_busy,
            tz=tz,
        )

    def _time_visible(self, event: Event) -> bool:
        return self._policy.account_allowed(event.account) and not self._policy.folder_excluded(
            event.calendar
        )
