"""PYTEST_DONT_REWRITE: integration tests against the real classic Outlook (`uv run pytest -m
outlook`).

A failure must never print mail content or an ID (rule 5). The marker above turns off
pytest's assertion rewriting for this module, and every assert tests a boolean or a count
computed beforehand, with a literal message: pytest has no operand to show. Default
Config(): the user's config.toml is not read here.
"""

from __future__ import annotations

import sys
from collections.abc import Iterator
from datetime import datetime, timedelta

import pytest

from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.models import SNIPPET_MAX
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
    reachable = status.reachable
    assert reachable, "Outlook injoignable : voir outlook-bridge status"
    classic = status.classic
    assert classic, "Outlook classique attendu"
    accounts = len(status.accounts)
    assert accounts > 0, "aucun compte"
    version_16 = (status.outlook_version or "").startswith("16.")
    assert version_16, "Outlook 16.x attendu"


def test_search_mail_shape(recent):
    flagged = all(m.untrusted_content for m in recent)
    assert flagged, "untrusted_content absent"
    in_paris = all(m.received.tzinfo == PARIS for m in recent)
    assert in_paris, "date de réception hors Europe/Paris"
    bounded = all(len(m.snippet) <= SNIPPET_MAX for m in recent)
    assert bounded, "extrait trop long"
    with_ids = all(m.id and m.store_id and m.conversation_id for m in recent)
    assert with_ids, "identifiant vide"
    received = [m.received for m in recent]
    newest_first = received == sorted(received, reverse=True)
    assert newest_first, "mails pas du plus récent au plus ancien"


def test_received_matches_outlook_wall_clock(recent, raw_ns):
    for mail in recent[:5]:
        item = raw_ns.GetItemFromID(mail.id, mail.store_id)
        wall_clock = naive(item.ReceivedTime)  # local time mislabelled UTC (Phase 0)
        close = abs(naive(mail.received) - wall_clock) < ONE_SECOND
        assert close, "date de réception différente d'Outlook"


def test_since_filter_agrees_with_dates(bridge, recent):
    since = recent[len(recent) // 2].received
    filtered = bridge.mail.search_mail(since=since, limit=100)
    expected = {m.id for m in recent if m.received >= since}
    missing = len(expected - {m.id for m in filtered})
    assert missing == 0, "mails manquants avec since"
    after = all(m.received >= since - ONE_SECOND for m in filtered)
    assert after, "mail antérieur à since"


def test_ids_are_long_term_entry_ids(recent, raw_ns):
    # Table rows carry session-only short-term ids in cached mode; ours must be stable.
    stable = all(raw_ns.GetItemFromID(m.id, m.store_id).EntryID == m.id for m in recent[:5])
    assert stable, "EntryID à court terme renvoyé"


def test_unread_only(bridge):
    only_unread = all(m.unread for m in bridge.mail.search_mail(unread_only=True, limit=20))
    assert only_unread, "mail lu avec unread_only"


def test_get_mail_roundtrip(bridge, recent):
    summary = recent[0]
    detail = bridge.mail.get_mail(summary.id, summary.store_id, max_chars=50)
    same_ids = detail.id == summary.id and detail.conversation_id == summary.conversation_id
    assert same_ids, "identifiants différents entre liste et détail"
    body_bounded = len(detail.body_text) <= 50
    assert body_bounded, "corps plus long que max_chars"
    close = abs(detail.received - summary.received) < ONE_SECOND
    assert close, "dates de réception différentes entre liste et détail"
    flagged = detail.untrusted_content is True
    assert flagged, "untrusted_content absent"


def test_get_mail_never_marks_as_read(bridge, raw_ns):
    unread = bridge.mail.search_mail(unread_only=True, limit=1)
    if not unread:
        pytest.skip("no unread mail")
    mail = unread[0]
    bridge.mail.get_mail(mail.id, mail.store_id)
    still_unread = raw_ns.GetItemFromID(mail.id, mail.store_id).UnRead is True
    assert still_unread, "get_mail a marqué le mail comme lu"


def test_get_thread_contains_the_mail(bridge, recent):
    summary = recent[0]
    thread = bridge.mail.get_thread(summary.conversation_id)
    found = any(m.id == summary.id for m in thread)
    assert found, "mail absent de sa conversation"
    received = [m.received for m in thread]
    chronological = received == sorted(received)
    assert chronological, "conversation pas dans l'ordre chronologique"


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
        same_starts = sorted(naive(e.start) for e in events) == sorted(raw)
        assert same_starts, "dates de début différentes d'Outlook"
        checked += len(raw)
    if checked == 0:
        pytest.skip("no event in the next 62 days")


def test_recurrences_are_expanded(bridge):
    start, end = window()
    events = bridge.calendar.list_events(start, end)
    recurring = [e for e in events if e.is_recurring]
    if not recurring:
        pytest.skip("no recurring occurrence in the next 62 days")
    well_formed = all(e.start.tzinfo == PARIS and e.end > e.start for e in events)
    assert well_formed, "événement hors Europe/Paris ou de durée nulle"
    # Occurrences share the series EntryID but each one has its own start.
    pairs = [(e.id, e.start) for e in recurring]
    distinct = len(pairs) == len(set(pairs))
    assert distinct, "occurrences en double"


def test_get_event_occurrence(bridge):
    start, end = window()
    recurring = [e for e in bridge.calendar.list_events(start, end) if e.is_recurring]
    if not recurring:
        pytest.skip("no recurring occurrence in the next 62 days")
    occurrence = recurring[-1]
    detail = bridge.calendar.get_event(
        occurrence.id, occurrence.store_id, max_chars=100, start=occurrence.start
    )
    same_start = detail.start == occurrence.start
    assert same_start, "occurrence différente de celle demandée"
    body_bounded = detail.body_text is not None and len(detail.body_text) <= 100
    assert body_bounded, "corps absent ou plus long que max_chars"


def test_find_free_slots_within_working_hours(bridge):
    start, _ = window()
    slots = bridge.calendar.find_free_slots(start, start + timedelta(days=14), min_duration_min=60)
    in_hours = all(9 <= s.start.hour and (s.end.hour, s.end.minute) <= (19, 0) for s in slots)
    assert in_hours, "créneau hors des heures de travail"
    weekdays = all(s.start.weekday() < 5 for s in slots)
    assert weekdays, "créneau un week-end"
    long_enough = all(s.duration_min >= 60 for s in slots)
    assert long_enough, "créneau de moins de 60 min"
