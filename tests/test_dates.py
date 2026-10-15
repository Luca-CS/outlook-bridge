"""Date grammar of the CLI (design section 5) and the Phase 2 timeutil helpers."""

from datetime import UTC, date, datetime, timedelta, timezone
from zoneinfo import ZoneInfo

import pytest

from outlook_bridge.errors import InvalidRequest
from outlook_bridge.timeutil import (
    ceil_to,
    day_start,
    ensure_aware,
    normalize,
    parse_when,
)

PARIS = ZoneInfo("Europe/Paris")
NOW = datetime(2026, 10, 12, 12, 0, tzinfo=PARIS)  # lundi, same value as tests/fakes.NOW


def paris(*args: int, **kwargs: int) -> datetime:
    return datetime(*args, tzinfo=PARIS, **kwargs)


def check(text: str, role: str, expected: str, now: datetime = NOW) -> None:
    result = parse_when(text, role, now=now)
    assert result.tzinfo is PARIS
    assert result.isoformat(timespec="minutes") == expected


# --- Forms table (section 5), now = lun. 12/10/2026 12:00+02:00 ----------------------------

FORMS = [
    # instants: same value for every role
    ("now", "from", "2026-10-12T12:00+02:00"),
    ("now", "since", "2026-10-12T12:00+02:00"),
    ("now", "to", "2026-10-12T12:00+02:00"),
    ("maintenant", "from", "2026-10-12T12:00+02:00"),
    (" Maintenant ", "to", "2026-10-12T12:00+02:00"),
    ("NOW", "since", "2026-10-12T12:00+02:00"),
    # today
    ("today", "since", "2026-10-12T00:00+02:00"),
    ("today", "from", "2026-10-12T00:00+02:00"),
    ("today", "to", "2026-10-13T00:00+02:00"),
    ("auj", "from", "2026-10-12T00:00+02:00"),
    ("AUJ", "to", "2026-10-13T00:00+02:00"),
    ("aujourd'hui", "from", "2026-10-12T00:00+02:00"),
    ("Aujourd\u2019hui", "since", "2026-10-12T00:00+02:00"),
    ("  Today\t", "from", "2026-10-12T00:00+02:00"),
    # tomorrow / yesterday
    ("demain", "from", "2026-10-13T00:00+02:00"),
    ("Demain", "to", "2026-10-14T00:00+02:00"),
    ("tomorrow", "from", "2026-10-13T00:00+02:00"),
    ("hier", "since", "2026-10-11T00:00+02:00"),
    ("HIER", "from", "2026-10-11T00:00+02:00"),
    ("yesterday", "to", "2026-10-12T00:00+02:00"),
    # day offsets (d or j), unsigned follows the role
    ("2d", "since", "2026-10-10T00:00+02:00"),
    ("2j", "since", "2026-10-10T00:00+02:00"),
    ("-2d", "since", "2026-10-10T00:00+02:00"),
    ("2d", "from", "2026-10-14T00:00+02:00"),
    ("+2j", "from", "2026-10-14T00:00+02:00"),
    ("-2d", "from", "2026-10-10T00:00+02:00"),
    ("2d", "to", "2026-10-15T00:00+02:00"),
    ("-2j", "to", "2026-10-11T00:00+02:00"),
    ("0d", "since", "2026-10-12T00:00+02:00"),
    ("999d", "from", "2029-07-07T00:00+02:00"),
    ("2D", "from", "2026-10-14T00:00+02:00"),
    # hour offsets: instants, elapsed time
    ("3h", "since", "2026-10-12T09:00+02:00"),
    ("-3h", "from", "2026-10-12T09:00+02:00"),
    ("3h", "from", "2026-10-12T15:00+02:00"),
    ("+3h", "to", "2026-10-12T15:00+02:00"),
    ("14h", "to", "2026-10-13T02:00+02:00"),  # an offset, not a time of day (form order)
    ("+0h", "since", "2026-10-12T12:00+02:00"),
    # absolute days
    ("2026-10-16", "from", "2026-10-16T00:00+02:00"),
    ("2026-10-16", "to", "2026-10-17T00:00+02:00"),
    ("16/10/2026", "from", "2026-10-16T00:00+02:00"),
    ("16/10/2026", "to", "2026-10-17T00:00+02:00"),
    ("16/10", "from", "2026-10-16T00:00+02:00"),
    ("16/10", "to", "2026-10-17T00:00+02:00"),
    ("5/1", "from", "2027-01-05T00:00+01:00"),  # nearest occurrence of 5 January
    ("09/10", "since", "2026-10-09T00:00+02:00"),
    ("2026-01-15", "since", "2026-01-15T00:00+01:00"),
    # date + time: instants
    ("2026-10-16 14:30", "from", "2026-10-16T14:30+02:00"),
    ("2026-10-16T14:30", "to", "2026-10-16T14:30+02:00"),
    ("2026-10-16t14h", "from", "2026-10-16T14:00+02:00"),
    ("16/10/2026 14h30", "to", "2026-10-16T14:30+02:00"),
    ("16/10 14h", "to", "2026-10-16T14:00+02:00"),
    ("16/10 9:05", "from", "2026-10-16T09:05+02:00"),
    ("16/10\u00a0 14H30", "from", "2026-10-16T14:30+02:00"),  # NBSP + space collapsed
    ("10/10 08h", "since", "2026-10-10T08:00+02:00"),
    # time alone: today at that local time
    ("14:00", "from", "2026-10-12T14:00+02:00"),
    ("14h00", "to", "2026-10-12T14:00+02:00"),
    ("9h30", "from", "2026-10-12T09:30+02:00"),
    ("08:15", "since", "2026-10-12T08:15+02:00"),
    ("00:00", "from", "2026-10-12T00:00+02:00"),
    # ISO with an offset or Z: converted to Paris
    ("2026-10-15T14:00+00:00", "from", "2026-10-15T16:00+02:00"),
    ("2026-10-15T12:00Z", "to", "2026-10-15T14:00+02:00"),
    ("2026-10-15t12:00z", "from", "2026-10-15T14:00+02:00"),
    ("2026-11-15T12:00:30-05:00", "from", "2026-11-15T18:00+01:00"),
    ("2026-10-10T08:00+02:00", "since", "2026-10-10T08:00+02:00"),
]


@pytest.mark.parametrize(("text", "role", "expected"), FORMS)
def test_forms(text, role, expected):
    check(text, role, expected)


# --- Examples table (section 5) ------------------------------------------------------------

EXAMPLES = [
    (NOW, "2d", "since", "2026-10-10T00:00+02:00"),
    (NOW, "-2d", "since", "2026-10-10T00:00+02:00"),
    (NOW, "3h", "since", "2026-10-12T09:00+02:00"),
    (NOW, "hier", "since", "2026-10-11T00:00+02:00"),
    (NOW, "today", "from", "2026-10-12T00:00+02:00"),
    (NOW, "+7d", "to", "2026-10-20T00:00+02:00"),
    (NOW, "7d", "to", "2026-10-20T00:00+02:00"),
    (NOW, "today", "to", "2026-10-13T00:00+02:00"),
    (NOW, "16/10", "to", "2026-10-17T00:00+02:00"),
    (NOW, "16/10 14h", "to", "2026-10-16T14:00+02:00"),
    (NOW, "14:00", "from", "2026-10-12T14:00+02:00"),
    (NOW, "+14d", "to", "2026-10-27T00:00+01:00"),
    (NOW, "2026-10-15T14:00+00:00", "from", "2026-10-15T16:00+02:00"),
    (NOW, "2026-03-29 02:30", "from", "2026-03-29T03:30+02:00"),
    (paris(2026, 10, 24, 12), "+24h", "to", "2026-10-25T11:00+01:00"),
    (paris(2026, 10, 25, 12), "24h", "since", "2026-10-24T13:00+02:00"),
    (paris(2026, 3, 29, 12), "12h", "since", "2026-03-28T23:00+01:00"),
    (paris(2026, 12, 20, 12), "05/01", "to", "2027-01-06T00:00+01:00"),
    (paris(2027, 1, 10, 12), "20/12", "since", "2026-12-20T00:00+01:00"),
]


@pytest.mark.parametrize(("now", "text", "role", "expected"), EXAMPLES)
def test_examples(now, text, role, expected):
    check(text, role, expected, now=now)


def test_defaults_of_the_commands():
    # events: today -> +7d = 8 days; slots: +1d -> +7d = 7 days.
    start, end = parse_when("today", "from", now=NOW), parse_when("+7d", "to", now=NOW)
    assert (start, end) == (paris(2026, 10, 12), paris(2026, 10, 20))
    assert parse_when("+1d", "from", now=NOW) == paris(2026, 10, 13)


# --- DST ------------------------------------------------------------------------------------


def test_spring_gap_wall_time_is_normalised():
    result = parse_when("29/03/2026 02:30", "from", now=NOW)
    assert result.isoformat() == "2026-03-29T03:30:00+02:00"
    assert result.utcoffset() == timedelta(hours=2)


def test_ambiguous_wall_time_takes_first_occurrence():
    result = parse_when("2026-10-25 02:30", "from", now=NOW)
    assert result.isoformat() == "2026-10-25T02:30:00+02:00"
    assert result.astimezone(UTC) == datetime(2026, 10, 25, 0, 30, tzinfo=UTC)


def test_hour_offsets_are_elapsed_time_across_dst():
    autumn = parse_when("+24h", "to", now=paris(2026, 10, 24, 12))
    assert autumn.astimezone(UTC) - paris(2026, 10, 24, 12).astimezone(UTC) == timedelta(hours=24)
    spring = parse_when("-12h", "from", now=paris(2026, 3, 29, 12))
    assert paris(2026, 3, 29, 12).astimezone(UTC) - spring.astimezone(UTC) == timedelta(hours=12)


def test_day_offsets_land_on_local_midnight_across_dst():
    for days in range(0, 30):
        result = parse_when(f"{days}d", "from", now=NOW)
        assert (result.hour, result.minute) == (0, 0)
        assert result.date() == date(2026, 10, 12) + timedelta(days=days)


# --- DD/MM nearest occurrence ----------------------------------------------------------------


@pytest.mark.parametrize(
    ("now", "text", "expected"),
    [
        (paris(2026, 12, 20, 12), "05/01", date(2027, 1, 5)),
        (paris(2027, 1, 10, 12), "20/12", date(2026, 12, 20)),
        (NOW, "12/10", date(2026, 10, 12)),
        (NOW, "01/04", date(2027, 4, 1)),  # 171 days ahead vs 194 back
        (NOW, "01/05", date(2026, 5, 1)),  # 164 days back vs 201 ahead
        # tie (183 days each way, leap day in between): the future wins
        (paris(2027, 8, 31, 12), "01/03", date(2028, 3, 1)),
        (paris(2028, 1, 15, 12), "29/02", date(2028, 2, 29)),  # only valid in 2028
    ],
)
def test_day_month_is_nearest_occurrence(now, text, expected):
    assert parse_when(text, "from", now=now).date() == expected


def test_day_month_tie_is_really_a_tie():
    today = date(2027, 8, 31)
    assert (today - date(2027, 3, 1)).days == (date(2028, 3, 1) - today).days == 183


# --- DD/MM --to resolved around the start of the range (`after`) ----------------------------


def span(first: str, last: str, now: datetime = NOW) -> tuple[str, str]:
    """(--from, --to) as the CLI resolves them: --to around the resolved --from."""
    start = parse_when(first, "from", now=now)
    end = parse_when(last, "to", now=now, after=start)
    return start.isoformat(timespec="minutes"), end.isoformat(timespec="minutes")


@pytest.mark.parametrize(
    ("now", "first", "last", "expected"),
    [
        # 12/10/2026: 01/04 alone is next April (171 days ahead), 30/04 alone last April
        # (165 days back); as a range both are the same April.
        (NOW, "01/04", "30/04", ("2027-04-01T00:00+02:00", "2027-05-01T00:00+02:00")),
        (NOW, "10/04 14h", "14/04", ("2027-04-10T14:00+02:00", "2027-04-15T00:00+02:00")),
        (NOW, "01/04", "30/04 18:00", ("2027-04-01T00:00+02:00", "2027-04-30T18:00+02:00")),
        # across New Year, typed before it, on New Year's Eve and after it
        (NOW, "20/12", "05/01", ("2026-12-20T00:00+01:00", "2027-01-06T00:00+01:00")),
        (
            paris(2026, 12, 31, 18),
            "31/12 22h",
            "01/01 02h",
            ("2026-12-31T22:00+01:00", "2027-01-01T02:00+01:00"),
        ),
        (
            paris(2027, 1, 2, 9),
            "20/12",
            "05/01",
            ("2026-12-20T00:00+01:00", "2027-01-06T00:00+01:00"),
        ),
        (
            paris(2027, 1, 2, 9),
            "-20d",
            "31/12",
            ("2026-12-13T00:00+01:00", "2027-01-01T00:00+01:00"),
        ),
        # a full-year start keeps working with a yearless end
        (NOW, "28/12/2026", "03/01", ("2026-12-28T00:00+01:00", "2027-01-04T00:00+01:00")),
        (NOW, "today", "16/10", ("2026-10-12T00:00+02:00", "2026-10-17T00:00+02:00")),
    ],
)
def test_day_month_end_follows_the_start(now, first, last, expected):
    assert span(first, last, now) == expected


def test_day_month_end_alone_is_still_the_nearest_occurrence():
    assert parse_when("30/04", "to", now=NOW).date() == date(2026, 5, 1)  # 2026-04-30 included


def test_every_short_day_month_range_lands_after_its_start():
    """Any start over two years and any length the core accepts (62 days at most): the
    yearless end is the first occurrence after the start, whatever the year change."""
    for offset in range(0, 730, 3):
        start_day = date(2026, 1, 1) + timedelta(days=offset)
        now = paris(2026, 6, 1, 12) + timedelta(days=offset // 2)
        start = parse_when(start_day.isoformat(), "from", now=now)
        for length in range(0, 62):
            last_day = start_day + timedelta(days=length)
            end = parse_when(f"{last_day.day:02d}/{last_day.month:02d}", "to", now=now, after=start)
            assert end == day_start(last_day + timedelta(days=1), PARIS)


@pytest.mark.parametrize(
    ("first", "last", "end_day"),
    [
        ("16/10", "14/10", date(2026, 10, 15)),  # reversed: not 14/10/2027
        ("today", "10/10", date(2026, 10, 11)),
        ("16/10 14h", "16/10 10h", None),
    ],
)
def test_reversed_day_month_range_stays_reversed(first, last, end_day):
    """So the core reports "end before start", not a range one year long."""
    start = parse_when(first, "from", now=NOW)
    end = parse_when(last, "to", now=NOW, after=start)
    assert end.astimezone(UTC) < start.astimezone(UTC)
    if end_day is not None:
        assert end == day_start(end_day, PARIS)


@pytest.mark.parametrize("text", ["+7d", "demain", "16/10/2026", "2026-10-16", "14:00", "now"])
def test_after_only_changes_yearless_day_month(text):
    after = paris(2027, 4, 1)
    assert parse_when(text, "to", now=NOW, after=after) == parse_when(text, "to", now=NOW)


def test_after_follows_the_local_day_of_the_start():
    # 31/12 23:30 UTC is 01/01/2027 in Paris. 02/07 is then 182 days ahead (2027) against
    # 183 back; around the UTC day (31/12/2026) the 2026 occurrence would win.
    after = datetime(2026, 12, 31, 23, 30, tzinfo=UTC)
    assert parse_when("02/07", "to", now=NOW, after=after) == paris(2027, 7, 3)
    assert parse_when("02/07", "to", now=NOW, after=paris(2026, 12, 31)) == paris(2026, 7, 3)


def test_after_is_for_aware_to_bounds_only():
    with pytest.raises(ValueError):
        parse_when("16/10", "from", now=NOW, after=NOW)
    with pytest.raises(ValueError):
        parse_when("16/10", "since", now=NOW, after=NOW)
    with pytest.raises(ValueError):
        parse_when("16/10", "to", now=NOW, after=datetime(2026, 10, 12))


def test_after_keeps_invalid_dates_invalid():
    with pytest.raises(InvalidRequest, match="^Date invalide pour --to : « 31/02 »"):
        parse_when("31/02", "to", now=NOW, after=NOW)
    # 29/02 exists in 2028 only: valid around a 2028 start, invalid around a 2026 one.
    assert parse_when("29/02", "to", now=NOW, after=paris(2028, 2, 1)) == paris(2028, 3, 1)
    with pytest.raises(InvalidRequest):
        parse_when("29/02", "to", now=NOW, after=NOW)


# --- invalid input -------------------------------------------------------------------------

INVALID = [
    "",
    "   ",
    "2x",
    "+d",
    "7",
    "1000d",
    "1000h",
    "5m",
    "31/02",
    "29/02",  # no leap year among 2025-2027
    "25:00",
    "14h60",
    "24:00",
    "demain 14h",
    "today 14:00",
    "+7 d",
    "+-7d",
    "2026-13-01",
    "2026-1-15",
    "16/10/1899",
    "16/10/3000",
    "16/13",
    "1/1/26",
    "2026-10-15T14:00:30",  # ISO without offset and with seconds
    "1899-12-31T10:00Z",
    "16/10 14",
    "16/10  14h x",
    "x" * 41,
    "1" + " " * 50 + "d",
    "aujourdhui",
    "il y a 2 jours",
]


@pytest.mark.parametrize("role", ["since", "from", "to"])
@pytest.mark.parametrize("text", INVALID)
def test_invalid_input(text, role):
    with pytest.raises(InvalidRequest) as info:
        parse_when(text, role, now=NOW)
    message = str(info.value)
    assert message.startswith(f"Date invalide pour --{role} : « ")
    assert "Formats : now, today (auj), demain, hier, 2d, +7d, -3h" in message
    assert "heure optionnelle 14h30 ou 14:30." in message


def test_invalid_message_quotes_the_input_and_uses_the_label():
    with pytest.raises(InvalidRequest, match="Date invalide pour --to : « 2x ». Formats : "):
        parse_when("2x", "to", now=NOW)
    with pytest.raises(InvalidRequest, match="^Date invalide pour --depuis : "):
        parse_when("2x", "since", now=NOW, label="--depuis")


def test_invalid_message_echo_is_bounded_and_printable():
    with pytest.raises(InvalidRequest) as info:
        parse_when("a\x1b[2J" + "b" * 100, "from", now=NOW)
    message = str(info.value)
    assert "\x1b" not in message
    assert "b" * 41 not in message
    assert "..." in message


def test_length_limit_applies_after_whitespace_collapse():
    check("  16/10   14h  " + " " * 40, "to", "2026-10-16T14:00+02:00")


@pytest.mark.parametrize("text", ["+1d", "+2d", "demain", "+1h", "23:00", "16/10", "2026-10-13"])
def test_since_in_the_future_is_rejected(text):
    with pytest.raises(InvalidRequest) as info:
        parse_when(text, "since", now=NOW)
    message = str(info.value)
    assert message == (
        f"--since désigne une date future (« {text} ») : utilisez 2d pour « il y a 2 jours »."
    )


def test_since_future_check_compares_instants_in_the_repeated_hour():
    # now = second 02:15 (+01:00, 01:15 UTC); "02:30" is the first 02:30 (+02:00, 00:30 UTC).
    now = paris(2026, 10, 25, 2, 15, fold=1)
    result = parse_when("02:30", "since", now=now)
    assert result.astimezone(UTC) == datetime(2026, 10, 25, 0, 30, tzinfo=UTC)


@pytest.mark.parametrize("text", ["now", "today", "0d", "12:00", "-0h"])
def test_since_now_or_earlier_is_accepted(text):
    assert parse_when(text, "since", now=NOW) <= NOW


@pytest.mark.parametrize("text", ["+1d", "demain", "23:00"])
def test_future_is_fine_for_from_and_to(text):
    assert parse_when(text, "from", now=NOW) > NOW
    assert parse_when(text, "to", now=NOW) > NOW


def test_programming_errors():
    with pytest.raises(ValueError):
        parse_when("today", "from", now=datetime(2026, 10, 12, 12))  # naive now
    with pytest.raises(ValueError):
        parse_when("today", "until", now=NOW)  # type: ignore[arg-type]


def test_result_follows_the_timezone_of_now():
    utc_now = NOW.astimezone(UTC)
    result = parse_when("today", "from", now=utc_now)
    assert result.tzinfo is UTC
    assert result == datetime(2026, 10, 12, 0, 0, tzinfo=UTC)


# --- helpers: normalize, ensure_aware, day_start, ceil_to -----------------------------------


def test_ensure_aware_normalises_the_spring_gap():
    result = ensure_aware(datetime(2026, 3, 29, 2, 30))
    assert result.isoformat() == "2026-03-29T03:30:00+02:00"
    assert result.tzinfo is PARIS


def test_ensure_aware_takes_first_occurrence_of_ambiguous_time():
    assert ensure_aware(datetime(2026, 10, 25, 2, 30)).isoformat() == "2026-10-25T02:30:00+02:00"


def test_ensure_aware_keeps_existing_values():
    assert ensure_aware(datetime(2026, 1, 5, 9)).isoformat() == "2026-01-05T09:00:00+01:00"
    other = datetime(2026, 1, 5, 9, tzinfo=timezone(timedelta(hours=-5)))
    assert ensure_aware(other) == other
    assert ensure_aware(other).tzinfo is PARIS


def test_normalize():
    gap = paris(2026, 3, 29, 2, 30)  # nonexistent wall time, aware
    assert normalize(gap, PARIS).isoformat() == "2026-03-29T03:30:00+02:00"
    # Same-tzinfo comparison ignores offsets: compare instants in UTC.
    assert normalize(gap, PARIS).astimezone(UTC) == gap.astimezone(UTC)
    assert normalize(datetime(2026, 3, 29, 2, 30), PARIS).isoformat() == (
        "2026-03-29T03:30:00+02:00"
    )
    second = paris(2026, 10, 25, 2, 30, fold=1)
    assert normalize(second, PARIS).utcoffset() == timedelta(hours=1)
    assert normalize(datetime(2026, 10, 12, 10, tzinfo=UTC), PARIS) == paris(2026, 10, 12, 12)


@pytest.mark.parametrize(
    ("day", "expected"),
    [
        (date(2026, 10, 12), "2026-10-12T00:00:00+02:00"),
        (date(2026, 10, 25), "2026-10-25T00:00:00+02:00"),
        (date(2026, 10, 26), "2026-10-26T00:00:00+01:00"),
        (date(2026, 3, 29), "2026-03-29T00:00:00+01:00"),
        (date(2026, 3, 30), "2026-03-30T00:00:00+02:00"),
    ],
)
def test_day_start(day, expected):
    result = day_start(day, PARIS)
    assert result.isoformat() == expected
    assert result.tzinfo is PARIS


@pytest.mark.parametrize(
    ("value", "minutes", "expected"),
    [
        (paris(2026, 10, 12, 12, 7), 15, "2026-10-12T12:15:00+02:00"),
        (paris(2026, 10, 12, 12, 15), 15, "2026-10-12T12:15:00+02:00"),
        (paris(2026, 10, 12, 12, 15, 1), 15, "2026-10-12T12:30:00+02:00"),
        (paris(2026, 10, 12, 12, 15, 0, 1), 15, "2026-10-12T12:30:00+02:00"),
        (paris(2026, 10, 12, 23, 50), 15, "2026-10-13T00:00:00+02:00"),
        (paris(2026, 10, 12, 12, 7), 60, "2026-10-12T13:00:00+02:00"),
        # autumn: the first 02:50 (+02:00) rounds up to the second 02:00 (+01:00)
        (paris(2026, 10, 25, 2, 50), 15, "2026-10-25T02:00:00+01:00"),
        (paris(2026, 10, 25, 2, 50, fold=1), 15, "2026-10-25T03:00:00+01:00"),
        # spring: 01:50 (+01:00) rounds up to 03:00 (+02:00), 02:00 does not exist
        (paris(2026, 3, 29, 1, 50), 15, "2026-03-29T03:00:00+02:00"),
    ],
)
def test_ceil_to(value, minutes, expected):
    result = ceil_to(value, minutes)
    assert result.isoformat() == expected
    assert result.tzinfo is PARIS
    assert (
        timedelta(0) <= result.astimezone(UTC) - value.astimezone(UTC) < timedelta(minutes=minutes)
    )


def test_ceil_to_default_step_and_errors():
    assert ceil_to(paris(2026, 10, 12, 12, 7)) == paris(2026, 10, 12, 12, 15)
    with pytest.raises(ValueError):
        ceil_to(datetime(2026, 10, 12, 12, 7))
    with pytest.raises(ValueError):
        ceil_to(paris(2026, 10, 12, 12, 7), 0)
