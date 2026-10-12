"""Integration tests against the real classic Outlook (`uv run pytest -m outlook`).

Assertions only compare dates, counts and booleans so that a failure never prints mail
content. Default Config(): the user's config.toml is not read here.
"""

from __future__ import annotations

import sys
from collections.abc import Iterator
from datetime import datetime, timedelta

import pytest

from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.timeutil import DEFAULT_TZ as PARIS

if sys.platform == "win32":
    import pythoncom
    import win32com.client

    from outlook_bridge.com_gateway import ComGateway

pytestmark = [
    pytest.mark.outlook,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows + Outlook only"),
]

ONE_SECOND = timedelta(seconds=1)


def naive(value) -> datetime:
    return datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)


@pytest.fixture(scope="module")
def bridge() -> Iterator[Bridge]:
    gateway = ComGateway()
    yield Bridge(gateway, Config())
    gateway.close()


@pytest.fixture(scope="module")
def raw_ns():
    """Independent COM connection on the test thread, to cross-check the gateway."""
    pythoncom.CoInitialize()
    yield win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    pythoncom.CoUninitialize()


@pytest.fixture(scope="module")
def recent(bridge):
    mails = bridge.mail.search_mail(limit=20)
    if not mails:
        pytest.skip("empty inbox")
    return mails


def window() -> tuple[datetime, datetime]:
    start = datetime.now(PARIS).replace(hour=0, minute=0, second=0, microsecond=0)
    return start, start + timedelta(days=62)


def test_status(bridge):
    status = bridge.status()
    assert status.reachable, status.message
    assert status.classic
    assert status.accounts
    assert status.outlook_version and status.outlook_version.startswith("16.")


def test_search_mail_shape(recent):
    assert all(m.untrusted_content for m in recent)
    assert all(m.received.tzinfo == PARIS for m in recent)
    assert all(len(m.snippet) <= 200 for m in recent)
    assert all(m.id and m.store_id and m.conversation_id for m in recent)
    received = [m.received for m in recent]
    assert received == sorted(received, reverse=True)


def test_received_matches_outlook_wall_clock(recent, raw_ns):
    for mail in recent[:5]:
        item = raw_ns.GetItemFromID(mail.id, mail.store_id)
        wall_clock = naive(item.ReceivedTime)  # local time mislabelled UTC (Phase 0)
        assert abs(naive(mail.received) - wall_clock) < ONE_SECOND


def test_since_filter_agrees_with_dates(bridge, recent):
    since = recent[len(recent) // 2].received
    filtered = bridge.mail.search_mail(since=since, limit=100)
    expected = {m.id for m in recent if m.received >= since}
    missing = len(expected - {m.id for m in filtered})
    assert missing == 0
    assert all(m.received >= since - ONE_SECOND for m in filtered)


def test_ids_are_long_term_entry_ids(recent, raw_ns):
    # Table rows carry session-only short-term ids in cached mode; ours must be stable.
    stable = [raw_ns.GetItemFromID(m.id, m.store_id).EntryID == m.id for m in recent[:5]]
    assert all(stable)


def test_unread_only(bridge):
    assert all(m.unread for m in bridge.mail.search_mail(unread_only=True, limit=20))


def test_get_mail_roundtrip(bridge, recent):
    summary = recent[0]
    detail = bridge.mail.get_mail(summary.id, summary.store_id, max_chars=50)
    same_ids = (detail.id == summary.id, detail.conversation_id == summary.conversation_id)
    assert same_ids == (True, True)
    assert len(detail.body_text) <= 50
    assert abs(detail.received - summary.received) < ONE_SECOND
    assert detail.untrusted_content is True


def test_get_mail_never_marks_as_read(bridge, raw_ns):
    unread = bridge.mail.search_mail(unread_only=True, limit=1)
    if not unread:
        pytest.skip("no unread mail")
    mail = unread[0]
    bridge.mail.get_mail(mail.id, mail.store_id)
    assert raw_ns.GetItemFromID(mail.id, mail.store_id).UnRead is True


def test_get_thread_contains_the_mail(bridge, recent):
    summary = recent[0]
    thread = bridge.mail.get_thread(summary.conversation_id)
    assert any(m.id == summary.id for m in thread)
    assert [m.received for m in thread] == sorted(m.received for m in thread)


def _calendar_folders(folder, depth=0):
    if depth > 5:
        return
    if folder.DefaultItemType == 1:
        yield folder
    for i in range(1, folder.Folders.Count + 1):
        yield from _calendar_folders(folder.Folders.Item(i), depth + 1)


def test_event_starts_match_outlook_wall_clock(bridge, raw_ns):
    """StartUTC-based dates must equal Outlook's local Start, on every calendar folder."""
    start, end = window()
    flt = f"[Start] < '{end:%d/%m/%Y %H:%M}' AND [End] > '{start:%d/%m/%Y %H:%M}'"
    checked = 0
    for folder in _calendar_folders(raw_ns.GetDefaultFolder(9)):
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        restricted = items.Restrict(flt)
        raw = []
        item = restricted.GetFirst()
        while item is not None and len(raw) < 500:
            raw.append(naive(item.Start))
            item = restricted.GetNext()
        path = folder.FolderPath.split("\\", 3)[-1].replace("\\", "/")
        events = bridge.calendar.list_events(start, end, calendar=path)
        assert sorted(naive(e.start) for e in events) == sorted(raw)
        checked += len(raw)
    if checked == 0:
        pytest.skip("no event in the next 62 days")


def test_recurrences_are_expanded(bridge):
    start, end = window()
    events = bridge.calendar.list_events(start, end)
    recurring = [e for e in events if e.is_recurring]
    if not recurring:
        pytest.skip("no recurring occurrence in the next 62 days")
    assert all(e.start.tzinfo == PARIS and e.end > e.start for e in events)
    # Occurrences share the series EntryID but each one has its own start.
    pairs = [(e.id, e.start) for e in recurring]
    assert len(pairs) == len(set(pairs))


def test_get_event_occurrence(bridge):
    start, end = window()
    recurring = [e for e in bridge.calendar.list_events(start, end) if e.is_recurring]
    if not recurring:
        pytest.skip("no recurring occurrence in the next 62 days")
    occurrence = recurring[-1]
    detail = bridge.calendar.get_event(
        occurrence.id, occurrence.store_id, max_chars=100, start=occurrence.start
    )
    assert detail.start == occurrence.start
    assert detail.body_text is not None and len(detail.body_text) <= 100


def test_find_free_slots_within_working_hours(bridge):
    start, _ = window()
    slots = bridge.calendar.find_free_slots(start, start + timedelta(days=14), min_duration_min=60)
    assert all(9 <= s.start.hour and (s.end.hour, s.end.minute) <= (19, 0) for s in slots)
    assert all(s.start.weekday() < 5 for s in slots)
    assert all(s.duration_min >= 60 for s in slots)
