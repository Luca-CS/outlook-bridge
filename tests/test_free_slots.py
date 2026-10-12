from datetime import datetime, time, timedelta

import pytest

from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.errors import InvalidRequest
from outlook_bridge.events import compute_free_slots
from tests.fakes import PARIS, FakeGateway, make_event

TUE = datetime(2026, 10, 13, tzinfo=PARIS)
HOURS = (time(9), time(19))


def at(day: datetime, hour: int, minute: int = 0) -> datetime:
    return day.replace(hour=hour, minute=minute)


def slots(events, start, end, **kwargs):
    options = {
        "min_duration_min": 30,
        "working_hours": HOURS,
        "include_weekends": False,
        "treat_tentative_as_busy": True,
        "tz": PARIS,
    }
    options.update(kwargs)
    result = compute_free_slots(events, start, end, **options)
    return [(s.start.strftime("%a %H:%M"), s.end.strftime("%H:%M"), s.duration_min) for s in result]


def test_empty_day_is_one_working_hours_slot():
    assert slots([], TUE, TUE + timedelta(days=1)) == [("Tue 09:00", "19:00", 600)]


def test_events_split_the_day_and_overlaps_merge():
    events = [
        make_event(1, at(TUE, 10), 60),
        make_event(2, at(TUE, 10, 30), 60),  # overlaps the first one
        make_event(3, at(TUE, 14), 30),
        make_event(4, at(TUE, 8), 90),  # starts before working hours
    ]
    assert slots(events, TUE, TUE + timedelta(days=1)) == [
        ("Tue 09:30", "10:00", 30),
        ("Tue 11:30", "14:00", 150),
        ("Tue 14:30", "19:00", 270),
    ]


def test_min_duration_filters_short_gaps():
    events = [make_event(1, at(TUE, 9, 20), 40)]
    assert slots(events, TUE, TUE + timedelta(days=1), min_duration_min=60) == [
        ("Tue 10:00", "19:00", 540)
    ]


def test_busy_status_rules():
    events = [
        make_event(1, at(TUE, 9), 60, busy_status="free"),
        make_event(2, at(TUE, 10), 60, busy_status="tentative"),
        make_event(3, at(TUE, 11), 60, busy_status="working_elsewhere"),
        make_event(4, at(TUE, 12), 60, busy_status="oof"),
    ]
    end = at(TUE, 14)
    assert slots(events, TUE, end) == [
        ("Tue 09:00", "10:00", 60),
        ("Tue 11:00", "12:00", 60),
        ("Tue 13:00", "14:00", 60),
    ]
    assert slots(events, TUE, end, treat_tentative_as_busy=False) == [
        ("Tue 09:00", "12:00", 180),
        ("Tue 13:00", "14:00", 60),
    ]


def test_weekends_skipped_unless_requested():
    sat = datetime(2026, 10, 17, tzinfo=PARIS)
    assert slots([], sat, sat + timedelta(days=2)) == []
    assert len(slots([], sat, sat + timedelta(days=2), include_weekends=True)) == 2


def test_range_bounds_clip_working_hours():
    assert slots([], at(TUE, 15), at(TUE, 17)) == [("Tue 15:00", "17:00", 120)]


def test_all_day_busy_event_blocks_the_day():
    events = [make_event(1, TUE, 24 * 60, all_day=True, busy_status="oof")]
    assert slots(events, TUE, TUE + timedelta(days=2)) == [("Wed 09:00", "19:00", 600)]


def test_dst_change_keeps_local_working_hours():
    # Sunday 25/10/2026 is the switch to winter time; Monday 26 keeps 09:00-19:00 local.
    mon = datetime(2026, 10, 26, tzinfo=PARIS)
    (slot,) = compute_free_slots(
        [],
        datetime(2026, 10, 24, tzinfo=PARIS),
        mon + timedelta(days=1),
        min_duration_min=30,
        working_hours=HOURS,
        include_weekends=False,
        treat_tentative_as_busy=True,
        tz=PARIS,
    )
    assert slot.start == at(mon, 9)
    assert slot.start.utcoffset() == timedelta(hours=1)


def test_service_validates_and_ignores_keyword_exclusions_for_busy_time():
    events = [make_event(1, at(TUE, 9), 600, subject="Confidentiel")]
    bridge = Bridge(FakeGateway(events=events), Config(subject_keywords_excluded=["confidentiel"]))
    assert bridge.calendar.list_events(TUE, TUE + timedelta(days=1)) == []
    # Hidden by keyword, but the time is still busy: no free slot that day.
    assert bridge.calendar.find_free_slots(TUE, TUE + timedelta(days=1)) == []
    with pytest.raises(InvalidRequest):
        bridge.calendar.find_free_slots(TUE, TUE + timedelta(days=1), working_hours="bad")
    with pytest.raises(InvalidRequest):
        bridge.calendar.find_free_slots(TUE, TUE + timedelta(days=90))


def test_service_ignores_excluded_calendars():
    events = [make_event(1, at(TUE, 9), 600, calendar="Éléments supprimés")]
    bridge = Bridge(FakeGateway(events=events), Config())
    assert len(bridge.calendar.find_free_slots(TUE, TUE + timedelta(days=1))) == 1
