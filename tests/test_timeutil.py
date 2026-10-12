from datetime import UTC, datetime, time, timedelta, timezone
from zoneinfo import ZoneInfo

import pytest

from outlook_bridge.errors import InvalidRequest
from outlook_bridge.timeutil import (
    dasl_datetime,
    ensure_aware,
    jet_datetime,
    parse_working_hours,
    to_aware,
    to_com,
)

PARIS = ZoneInfo("Europe/Paris")


def test_local_value_mislabelled_utc_is_read_as_wall_clock():
    # What pywin32 returns for a 14:00 Paris meeting: 14:00 tagged UTC (Phase 0).
    com_value = datetime(2026, 10, 15, 14, 0, tzinfo=UTC)
    result = to_aware(com_value)
    assert result == datetime(2026, 10, 15, 14, 0, tzinfo=PARIS)
    assert result.utcoffset() == timedelta(hours=2)


def test_true_utc_value_is_converted_to_paris():
    com_value = datetime(2026, 10, 15, 12, 0, tzinfo=UTC)  # StartUTC of the same meeting
    assert to_aware(com_value, utc=True) == datetime(2026, 10, 15, 14, 0, tzinfo=PARIS)


def test_both_sources_agree_across_dst_change():
    # 11 November: winter time, offset +1.
    local = to_aware(datetime(2026, 11, 11, 9, 0, tzinfo=UTC))
    utc = to_aware(datetime(2026, 11, 11, 8, 0, tzinfo=UTC), utc=True)
    assert local == utc
    assert utc.utcoffset() == timedelta(hours=1)


def test_utc_source_disambiguates_the_repeated_hour():
    # 25/10/2026 02:30 happens twice in Paris; only the UTC source tells them apart.
    first = to_aware(datetime(2026, 10, 25, 0, 30), utc=True)
    second = to_aware(datetime(2026, 10, 25, 1, 30), utc=True)
    assert first.hour == second.hour == 2
    assert (first.utcoffset(), second.utcoffset()) == (timedelta(hours=2), timedelta(hours=1))
    # Same-tzinfo arithmetic ignores `fold`: compare in UTC.
    assert second.astimezone(UTC) - first.astimezone(UTC) == timedelta(hours=1)


def test_to_com_is_the_inverse_of_to_aware():
    aware = datetime(2026, 11, 11, 9, 0, tzinfo=PARIS)
    com = to_com(aware)
    assert com == datetime(2026, 11, 11, 9, 0, tzinfo=UTC)  # wall clock tagged UTC
    assert to_aware(com) == aware
    assert to_com(datetime(2026, 11, 11, 8, 0, tzinfo=UTC)) == com


def test_naive_value_and_none_and_no_date():
    assert to_aware(datetime(2026, 1, 5, 9, 30)) == datetime(2026, 1, 5, 9, 30, tzinfo=PARIS)
    assert to_aware(None) is None
    assert to_aware(datetime(4501, 1, 1)) is None


def test_ensure_aware():
    assert ensure_aware(datetime(2026, 1, 5, 9)) == datetime(2026, 1, 5, 9, tzinfo=PARIS)
    other = datetime(2026, 1, 5, 9, tzinfo=timezone(timedelta(hours=-5)))
    assert ensure_aware(other).tzinfo == PARIS
    assert ensure_aware(other) == other


def test_jet_format_is_day_first_local():
    assert jet_datetime(datetime(2026, 10, 9, 0, 0, tzinfo=PARIS)) == "09/10/2026 00:00"
    assert jet_datetime(datetime(2026, 10, 8, 22, 0, tzinfo=UTC)) == "09/10/2026 00:00"


def test_dasl_format_is_iso_utc():
    assert dasl_datetime(datetime(2026, 10, 9, 0, 0, tzinfo=PARIS)) == "2026-10-08 22:00"
    assert dasl_datetime(datetime(2026, 1, 9, 0, 0)) == "2026-01-08 23:00"


def test_working_hours():
    assert parse_working_hours("09:00-19:00") == (time(9), time(19))
    assert parse_working_hours(" 08:30 - 12:00 ") == (time(8, 30), time(12))
    for bad in ("9h-19h", "19:00-09:00", "09:00"):
        with pytest.raises(InvalidRequest):
            parse_working_hours(bad)
