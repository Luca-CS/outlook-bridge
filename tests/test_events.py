from datetime import datetime, timedelta

import pytest

from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.errors import InvalidRequest, NotFound
from outlook_bridge.events import build_jet_range, find_meeting_url
from tests.fakes import NOW, PARIS, STORE, FakeGateway, make_event

DAY = datetime(2026, 10, 13, tzinfo=PARIS)  # Tuesday


def bridge_with(events, **config) -> tuple[Bridge, FakeGateway]:
    gw = FakeGateway(events=events)
    return Bridge(gw, Config(**config)), gw


def test_list_events_sorted_and_flagged():
    events = [make_event(2, DAY.replace(hour=14)), make_event(1, DAY.replace(hour=9))]
    bridge, _ = bridge_with(events)
    result = bridge.calendar.list_events(DAY, DAY + timedelta(days=1))
    assert [e.id for e in result] == ["EVT-1", "EVT-2"]
    assert all(e.untrusted_content for e in result)


def test_range_limits():
    bridge, _ = bridge_with([])
    with pytest.raises(InvalidRequest, match="62"):
        bridge.calendar.list_events(DAY, DAY + timedelta(days=63))
    with pytest.raises(InvalidRequest):
        bridge.calendar.list_events(DAY, DAY)
    bridge.calendar.list_events(DAY, DAY + timedelta(days=62))


def test_naive_bounds_are_paris_time():
    bridge, gw = bridge_with([])
    bridge.calendar.list_events(datetime(2026, 10, 13), datetime(2026, 10, 14))
    assert gw.calls[-1][1] == datetime(2026, 10, 13, tzinfo=PARIS)


def test_list_events_exclusions_and_redaction():
    events = [
        make_event(1, DAY.replace(hour=9), subject="Entretien confidentiel"),
        make_event(2, DAY.replace(hour=10), location="Salle FR1234567890123456789012345"),
    ]
    bridge, _ = bridge_with(events, subject_keywords_excluded=["confidentiel"])
    result = bridge.calendar.list_events(DAY, DAY + timedelta(days=1))
    assert [e.id for e in result] == ["EVT-2"]
    assert "FR12" not in result[0].location


def test_excluded_calendar_is_refused():
    bridge, _ = bridge_with([])
    with pytest.raises(NotFound):
        bridge.calendar.list_events(DAY, DAY + timedelta(days=1), calendar="Éléments supprimés")


def test_get_event_body_and_exclusion():
    events = [
        make_event(1, DAY.replace(hour=9), body_text="y" * 500),
        make_event(2, DAY.replace(hour=10), subject="Confidentiel"),
    ]
    bridge, _ = bridge_with(events, subject_keywords_excluded=["confidentiel"])
    event = bridge.calendar.get_event("EVT-1", STORE, max_chars=50)
    assert len(event.body_text) == 50
    assert event.body_truncated
    with pytest.raises(NotFound):
        bridge.calendar.get_event("EVT-2", STORE)


def test_get_event_occurrence_start_is_forwarded():
    start = DAY.replace(hour=9)
    bridge, _ = bridge_with([make_event(1, start)])
    assert bridge.calendar.get_event("EVT-1", STORE, start=start.replace(tzinfo=None)).id == "EVT-1"
    with pytest.raises(NotFound):
        bridge.calendar.get_event("EVT-1", STORE, start=start + timedelta(days=7))


def test_jet_range_uses_day_first_local_format():
    assert build_jet_range(NOW, NOW + timedelta(days=7), PARIS) == (
        "[Start] < '19/10/2026 12:00' AND [End] > '12/10/2026 12:00'"
    )


@pytest.mark.parametrize(
    ("text", "expected"),
    [
        (
            "Rejoindre : https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0?context=x.",
            "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0?context=x",
        ),
        (
            "Zoom https://example.zoom.us/j/123456789?pwd=abc",
            "https://example.zoom.us/j/123456789?pwd=abc",
        ),
        ("(https://meet.google.com/abc-defg-hij)", "https://meet.google.com/abc-defg-hij"),
        ("Salle B204", None),
        ("https://example.com/not-a-meeting", None),
    ],
)
def test_find_meeting_url(text, expected):
    assert find_meeting_url(None, text) == expected
