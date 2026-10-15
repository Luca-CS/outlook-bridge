"""fmt: sanitisers, French labels, agenda helpers, path masking. Synthetic data only."""

from __future__ import annotations

import json
import locale
import re
import time
from datetime import UTC, date, datetime, timedelta
from html import unescape as html_unescape
from pathlib import Path, PureWindowsPath
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pytest

from outlook_bridge import fmt
from outlook_bridge.fmt import (
    ago,
    clean_block,
    clean_line,
    day_label,
    display_path,
    duration,
    event_flags,
    group_by_day,
    long_dt,
    mail_flags,
    mail_marks,
    mask_paths,
    md_inline,
    range_label,
    short_day,
    short_dt,
    size,
    time_span,
    to_json,
)
from outlook_bridge.models import Event, MailSummary

PARIS = ZoneInfo("Europe/Paris")
NOW = datetime(2026, 10, 12, 12, 0, tzinfo=PARIS)  # lundi
MONDAY = date(2026, 10, 12)


def at(day: int, hour: int = 0, minute: int = 0, month: int = 10) -> datetime:
    return datetime(2026, month, day, hour, minute, tzinfo=PARIS)


def mail(**overrides) -> MailSummary:
    data = {
        "id": "MAIL-1",
        "store_id": "STORE-1",
        "conversation_id": "CONV-1",
        "subject": "Réunion test",
        "sender_name": "Bob Exemple",
        "sender_email": "bob@example.com",
        "received": NOW - timedelta(hours=1),
        "unread": False,
        "folder": "Inbox",
        "account": "alice@example.com",
    }
    data.update(overrides)
    return MailSummary(**data)


def event(n: int, start: datetime, end: datetime, **overrides) -> Event:
    data = {
        "id": f"EVT-{n}",
        "store_id": "STORE-1",
        "calendar": "Calendar",
        "account": "alice@example.com",
        "subject": f"Point projet {n}",
        "start": start,
        "end": end,
    }
    data.update(overrides)
    return Event(**data)


# ---------------------------------------------------------------------------- clean_line


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("\x1b[31mRouge\x1b[0m test", "Rouge test"),  # CSI colour
        ("\x1b[2J\x1b[1;1HEffacé", "Effacé"),  # CSI erase + cursor move
        ("\x1b]8;;https://example.com\x1b\\lien\x1b]8;;\x1b\\", "lien"),  # OSC 8 hyperlink
        ("\x1b]0;Titre piégé\x07Sujet", "Sujet"),  # OSC window title (BEL)
        ("\x1b]2;Titre\x1b\\Sujet", "Sujet"),  # OSC title (ST)
        ("a\x07b\x00c\x7fd", "abcd"),  # BEL, NUL, DEL
        ("\x9b2JSujet", "Sujet"),  # C1 CSI
        ("a\x9cb\x85c", "ab c"),  # C1 ST removed, NEL is a line break
        ("abc\u202edef\u2066g\u2069h\u200ei\u061cj", "abcdefghij"),  # bidi controls
        ("\ufeffRéunion", "Réunion"),  # BOM
        ("a\r\nb\tc\u2028d\u2029e\x0bf\x0cg", "a b c d e f g"),
        ("a\xa0\xa0b\u202fc", "a b c"),  # NBSP runs
        ("  Ligne1\n---\ngenerated_at: faux\n# Titre  ", "Ligne1 --- generated_at: faux # Titre"),
        ("vis\U000e0041\U000e0042ible", "visible"),  # Unicode tags (invisible text)
        ("lone \ud800surrogate", "lone surrogate"),
        ("Réunion 🎉 « test » · ok", "Réunion 🎉 « test » · ok"),  # printable text untouched
    ],
)
def test_clean_line_removes_controls_and_sequences(raw: str, expected: str):
    out = clean_line(raw)
    assert out == expected
    assert out.isprintable()


@pytest.mark.parametrize("raw", [None, "", "   ", "\x1b[0m", "\n\t\r"])
def test_clean_line_empty(raw):
    assert clean_line(raw) == ""


def test_clean_line_cuts_with_ascii_ellipsis():
    assert clean_line("x" * 50, 10) == "xxxxxxx..."
    assert len(clean_line("Réunion test très longue", 12)) == 12
    assert clean_line("Réunion test", 12) == "Réunion test"  # exact length: no cut
    assert clean_line("ab cd ef", 6) == "ab..."  # no space before the ellipsis
    assert clean_line("abcdef", 2) == "ab"


# Introducers of the string sequences (OSC, DCS, SOS, PM, APC), C1 and ESC forms.
INTRODUCERS = ["\x9d", "\x90", "\x98", "\x9e", "\x9f", "\x1b]", "\x1bP", "\x1bX", "\x1b^", "\x1b_"]


def _timed(clean, text: str) -> tuple[str, float]:
    start = time.perf_counter()
    out = clean(text)
    return out, time.perf_counter() - start


@pytest.mark.parametrize("clean", [clean_line, clean_block], ids=["line", "block"])
@pytest.mark.parametrize("intro", INTRODUCERS, ids=lambda s: s.encode("unicode_escape").decode())
def test_cleaning_is_linear_on_hostile_input(clean, intro: str):
    """A body (up to max_body_chars) full of unterminated sequences: linear is ~10 ms here,
    a backtracking regex takes seconds. The bound leaves a wide margin for a slow machine."""
    out, elapsed = _timed(clean, intro * 20000 + "fin")
    assert out.endswith("fin")
    assert elapsed < 1.0
    out, elapsed = _timed(clean, "\x9d" * 20000 + "\x1b]" * 20000 + "fin")
    assert out.endswith("fin")
    assert elapsed < 1.0


# ---------------------------------------------------------------------------- clean_block


def test_clean_block_keeps_lines_and_caps_blank_runs():
    raw = (
        "\ufeffBonjour\r\n\r\n\r\n\r\n\r\nFin\rX\tY  \n\n"
        "\x1b[2J\x07Signé\u2028\u2028\u2028\u2028Z\n\n"
    )
    assert clean_block(raw) == "Bonjour\n\n\nFin\nX    Y\n\nSign\xe9\n\n\nZ"


def test_clean_block_details():
    assert clean_block(None) == ""
    assert clean_block("\n\n  \nCorps\n\n") == "Corps"
    assert clean_block("  indenté\n\tun\xa0deux") == "  indenté\n    un deux"
    assert clean_block("a\n\n\nb") == "a\n\n\nb"  # 2 blank lines kept
    assert clean_block("a\n\n\n\n\n\nb") == "a\n\n\nb"
    out = clean_block("\x1b]8;;https://example.com\x1b\\lien\x1b]8;;\x1b\\\n\u202esuite\x00")
    assert out == "lien\nsuite"
    assert all(line.isprintable() for line in out.split("\n"))


# ---------------------------------------------------------------------------- md_inline

HOSTILE_MD = "![x](https://example.com/p.png) [lien](https://example.com) <img src=x> | `c`"


def test_md_inline_escapes_markdown_specials():
    assert md_inline("a_b*c") == r"a\_b\*c"
    assert md_inline("C:\\dossier") == r"C:\\dossier"
    assert md_inline("Ligne1\n---\n# Titre") == "Ligne1 --- # Titre"
    assert md_inline(None) == ""


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("# Instructions", r"\# Instructions"),
        ("  ## Titre", r"\## Titre"),  # clean_line strips the indent first
        ("- x", r"\- x"),
        ("+ x", r"\+ x"),
        ("---", r"\---"),
        ("===", r"\==="),
        ("~~~ py", r"\~~~ py"),
        ("1. étape", r"1\. étape"),
        ("12) étape", r"12\) étape"),
        ("> cité", r"\> cité"),
        ("Point #3 - a + b = c", "Point #3 - a + b = c"),  # only a leading marker
        ("2026-10-12 1.5", "2026-10-12 1.5"),
        ("a & b", "a &amp; b"),
        ("&#x202E;fdp.exe", "&amp;#x202E;fdp.exe"),
    ],
)
def test_md_inline_neutralises_block_markers_and_entities(raw: str, expected: str):
    assert md_inline(raw) == expected


def test_md_inline_cuts_before_escaping():
    assert md_inline("*" * 10, 5) == r"\*\*..."
    assert md_inline("ab\\cdefgh", 6) == r"ab\\..."  # the backslash stays escaped


def _tags(html: str) -> set[str]:
    return set(re.findall(r"</?[a-zA-Z][^>]*>", html))


ALLOWED_TAGS = {"p", "ul", "li", "strong", "blockquote"}


BLOCK_STARTS = ["# x", "## x", "1. x", "1) x", "- x", "+ x", "---", "===", "~~~", "> x"]


@pytest.mark.parametrize(
    "raw",
    [
        HOSTILE_MD,
        "<script>alert(1)</script>",
        "**gras** _it_ x\ny",
        *BLOCK_STARTS,
        "&#x202E;fdp.exe",
        "Facture &#8238;fdp &lt;b&gt; &amp;",
    ],
)
def test_md_inline_renders_inert(raw: str):
    markdown_it = pytest.importorskip("markdown_it")
    md = markdown_it.MarkdownIt("commonmark", {"html": True}).enable("table")
    out = md_inline(raw)
    assert "\n" not in out
    contexts = (out, f"- 12:00 · **{out}** · ok", f"  > {out}", f"- ligne\n  > {out}", f"t\n{out}")
    for doc in contexts:
        html = md.render(doc)
        assert "<img" not in html and "<a" not in html and "<script" not in html
        assert "‮" not in html and "<b>" not in html  # no entity was decoded
        assert _tags(html) <= {f"<{t}>" for t in ALLOWED_TAGS} | {f"</{t}>" for t in ALLOWED_TAGS}
    cell = md.render(f"| a | b |\n|---|---|\n| {out} | x |")
    assert cell.count("<td>") == 2  # the pipe did not open a new column


def test_md_inline_round_trips_text():
    markdown_it = pytest.importorskip("markdown_it")
    md = markdown_it.MarkdownIt("commonmark")
    html = md.render(md_inline(HOSTILE_MD))
    assert "[lien](https://example.com)" in html
    assert "&lt;img src=x&gt;" in html
    for raw in [*BLOCK_STARTS, "&#x202E;fdp.exe", "a & b", "Point #3"]:
        rendered = md.render(f"  > {md_inline(raw)}")
        text = re.sub(r"<[^>]+>", "", rendered).strip()
        assert html_unescape(text) == raw  # the reader sees the original text


# ---------------------------------------------------------------------------- date labels


def test_day_label():
    assert day_label(MONDAY) == "lundi 12 octobre"
    assert day_label(MONDAY, year=True) == "lundi 12 octobre 2026"
    assert day_label(date(2026, 8, 1), year=True) == "samedi 1er août 2026"
    assert day_label(date(2026, 2, 1)) == "dimanche 1er février"
    assert day_label(date(2026, 12, 31)) == "jeudi 31 décembre"


def test_short_labels():
    assert short_day(MONDAY) == "lun. 12/10"
    assert short_day(MONDAY, year=True) == "lun. 12/10/2026"
    assert short_day(date(2026, 1, 5)) == "lun. 05/01"
    assert short_dt(at(12, 11), PARIS) == "lun. 12/10 11:00"
    assert short_dt(datetime(2026, 10, 12, 9, 5, tzinfo=UTC), PARIS) == "lun. 12/10 11:05"
    assert short_dt(datetime(2026, 10, 12, 11, 0), PARIS) == "lun. 12/10 11:00"  # naive = local
    assert long_dt(NOW, PARIS) == "lundi 12 octobre 2026 à 12:00"
    assert long_dt(datetime(2026, 10, 25, 23, 30, tzinfo=UTC), PARIS) == (
        "lundi 26 octobre 2026 à 00:30"
    )


@pytest.mark.parametrize("name", ["C", "", "German_Germany.1252", "de_DE.UTF-8", "en_US.UTF-8"])
def test_labels_ignore_the_process_locale(name: str):
    saved = locale.setlocale(locale.LC_TIME)
    try:
        try:
            locale.setlocale(locale.LC_TIME, name)
        except locale.Error:
            pytest.skip(f"locale {name!r} indisponible")
        assert day_label(date(2026, 8, 1), year=True) == "samedi 1er août 2026"
        assert long_dt(NOW, PARIS) == "lundi 12 octobre 2026 à 12:00"
        assert short_dt(NOW, PARIS) == "lun. 12/10 12:00"
    finally:
        locale.setlocale(locale.LC_TIME, saved)


@pytest.mark.parametrize(
    ("start", "end", "expected"),
    [
        (at(12), at(20), "du lun. 12/10 au lun. 19/10/2026 inclus"),
        (at(13), at(20), "du mar. 13/10 au lun. 19/10/2026 inclus"),
        (at(12), at(13), "du lun. 12/10 au lun. 12/10/2026 inclus"),
        (at(12, 14), at(13, 9), "du lun. 12/10 14:00 au mar. 13/10/2026 09:00"),
        (at(12, 14), at(20), "du lun. 12/10 14:00 au lun. 19/10/2026 inclus"),
        # 14 days across the 25/10 DST change; end given in UTC
        (
            at(12),
            datetime(2026, 10, 25, 23, 0, tzinfo=UTC),
            "du lun. 12/10 au dim. 25/10/2026 inclus",
        ),
        (
            datetime(2026, 12, 20, tzinfo=PARIS),
            datetime(2027, 1, 6, tzinfo=PARIS),
            "du dim. 20/12/2026 au mar. 05/01/2027 inclus",
        ),
    ],
)
def test_range_label(start: datetime, end: datetime, expected: str):
    assert range_label(start, end, PARIS) == expected


# ---------------------------------------------------------------------------- small formatters


@pytest.mark.parametrize(
    ("minutes", "expected"),
    [(0, "0 min"), (45, "45 min"), (60, "1 h"), (90, "1 h 30"), (65, "1 h 05"), (600, "10 h")]
    + [(2340, "39 h")],
)
def test_duration(minutes: int, expected: str):
    assert duration(minutes) == expected


@pytest.mark.parametrize(
    ("n", "expected"),
    [
        (0, "0 o"),
        (850, "850 o"),
        (1023, "1023 o"),
        (1024, "1 Ko"),
        (1536, "2 Ko"),
        (12 * 1024, "12 Ko"),
        (int(3.4 * 1024**2), "3,4 Mo"),
        (2 * 1024**2, "2 Mo"),
        (15 * 1024**2, "15 Mo"),
        (int(1.2 * 1024**3), "1,2 Go"),
        (-5, "0 o"),
    ],
)
def test_size(n: int, expected: str):
    assert size(n) == expected


@pytest.mark.parametrize(
    ("delta", "expected"),
    [
        (timedelta(seconds=-30), "à l'instant"),
        (timedelta(seconds=59), "à l'instant"),
        (timedelta(minutes=4, seconds=50), "il y a 4 min"),
        (timedelta(minutes=59), "il y a 59 min"),
        (timedelta(hours=2, minutes=5), "il y a 2 h"),
        (timedelta(days=3, hours=2), "il y a 3 j"),
    ],
)
def test_ago(delta: timedelta, expected: str):
    assert ago(delta) == expected


def test_mail_marks_and_flags():
    full = mail(unread=True, importance="high", has_attachments=True)
    assert mail_marks(full) == "*!+"
    assert mail_flags(full) == ["non lu", "importance haute", "pièce jointe"]
    plain = mail()
    assert mail_marks(plain) == ""
    assert mail_flags(plain) == []
    attached = mail(has_attachments=True, importance="low")
    assert mail_marks(attached) == "+"
    assert mail_flags(attached) == ["importance faible", "pièce jointe"]
    assert mail_marks(mail(unread=True)) == "*"


# ---------------------------------------------------------------------------- agenda helpers


def agenda_events() -> list[Event]:
    return [
        event(1, at(12, 14), at(12, 15), location="Salle A"),
        event(2, at(11, 22), at(12, 1)),  # started before the range
        event(3, at(14), at(17), all_day=True, busy_status="oof"),  # 3 days, all-day
        event(4, at(16, 22), at(17, 1)),  # crosses midnight
        event(5, at(25, 23, 30), at(26, 0, 30)),  # last day, after the DST change
        event(6, at(12, 9), at(12, 10)),
        event(7, at(1), at(2)),  # entirely before the range
        event(8, at(26, 9), at(26, 10)),  # after the range
    ]


def test_group_by_day_lists_every_day_with_the_events_covering_it():
    events = agenda_events()
    days = group_by_day(events, at(12), at(26), PARIS)
    assert [d for d, _ in days] == [MONDAY + timedelta(days=i) for i in range(14)]
    by_day = {d: [e.id for e in evs] for d, evs in days}
    assert by_day[MONDAY] == ["EVT-2", "EVT-6", "EVT-1"]  # sorted, started-before first
    assert by_day[date(2026, 10, 13)] == []
    # a 3-day all-day event is under each of its days, not only the first one
    assert by_day[date(2026, 10, 14)] == ["EVT-3"]
    assert by_day[date(2026, 10, 15)] == ["EVT-3"]
    assert by_day[date(2026, 10, 16)] == ["EVT-3", "EVT-4"]
    assert by_day[date(2026, 10, 17)] == ["EVT-4"]  # crossed midnight
    assert by_day[date(2026, 10, 18)] == []
    assert by_day[date(2026, 10, 25)] == ["EVT-5"]  # its last 30 min are after the range
    listed = {e.id for _, evs in days for e in evs}
    assert listed == {"EVT-1", "EVT-2", "EVT-3", "EVT-4", "EVT-5", "EVT-6"}


def test_group_by_day_never_shows_a_covered_day_as_empty():
    leave = event(1, at(19), at(24), all_day=True, busy_status="oof", subject="Congés test")
    on_call = event(2, at(19, 18), at(21, 8), subject="Astreinte test")  # timed, 3 days
    midnight = event(3, at(22, 23), at(23))  # ends at 00:00: the next day is not touched
    instant = event(4, at(20, 0), at(20, 0))  # zero length, at midnight
    days = group_by_day([leave, on_call, midnight, instant], at(19), at(26), PARIS)
    by_day = {d: [e.id for e in evs] for d, evs in days}
    assert by_day == {
        date(2026, 10, 19): ["EVT-1", "EVT-2"],
        date(2026, 10, 20): ["EVT-1", "EVT-2", "EVT-4"],  # by start: continuations first
        date(2026, 10, 21): ["EVT-1", "EVT-2"],
        date(2026, 10, 22): ["EVT-1", "EVT-3"],
        date(2026, 10, 23): ["EVT-1"],
        date(2026, 10, 24): [],
        date(2026, 10, 25): [],
    }
    # a range that starts or ends inside the event lists only the days in the range
    inside = group_by_day([leave], at(21), at(22, 12), PARIS)
    assert [(d, [e.id for e in evs]) for d, evs in inside] == [
        (date(2026, 10, 21), ["EVT-1"]),
        (date(2026, 10, 22), ["EVT-1"]),
    ]


def test_group_by_day_partial_range_and_empty():
    days = group_by_day(agenda_events(), at(12, 14), at(13, 9), PARIS)
    assert [d for d, _ in days] == [MONDAY, date(2026, 10, 13)]
    assert [e.id for e in days[0][1]] == ["EVT-1"]  # the morning events ended before 14:00
    assert days[1][1] == []
    assert group_by_day(agenda_events(), at(13), at(13), PARIS) == []
    assert group_by_day([], at(12), at(13), PARIS) == [(MONDAY, [])]


def test_time_span():
    one_day = event(1, at(14), at(15), all_day=True)
    assert time_span(one_day, date(2026, 10, 14), PARIS) == "journée"
    multi = event(2, at(14), at(17), all_day=True)
    assert time_span(multi, date(2026, 10, 14), PARIS) == "journées jusqu'au ven. 16/10"
    # all-day across the DST change: 25/10 is a 25-hour day
    assert time_span(event(3, at(25), at(26), all_day=True), date(2026, 10, 25), PARIS) == "journée"
    timed = event(4, at(12, 14), at(12, 15))
    assert time_span(timed, MONDAY, PARIS) == "14:00-15:00"
    crossing = event(5, at(16, 22), at(17, 1))
    assert time_span(crossing, date(2026, 10, 16), PARIS) == "22:00 au sam. 17/10 01:00"
    utc = event(
        8, datetime(2026, 10, 12, 12, 0, tzinfo=UTC), datetime(2026, 10, 12, 13, 0, tzinfo=UTC)
    )
    assert time_span(utc, MONDAY, PARIS) == "14:00-15:00"
    midnight = event(9, at(16, 23), at(17))
    assert time_span(midnight, date(2026, 10, 16), PARIS) == "23:00 au sam. 17/10 00:00"


def test_time_span_marks_the_following_days_as_suite():
    leave = event(1, at(19), at(24), all_day=True)  # lun. 19/10 to ven. 23/10 included
    assert time_span(leave, date(2026, 10, 19), PARIS) == "journées jusqu'au ven. 23/10"
    assert time_span(leave, date(2026, 10, 20), PARIS) == "journées jusqu'au ven. 23/10 (suite)"
    assert time_span(leave, date(2026, 10, 22), PARIS) == "journées jusqu'au ven. 23/10 (suite)"
    assert time_span(leave, date(2026, 10, 23), PARIS) == "journée (suite)"
    crossing = event(2, at(16, 22), at(17, 1))
    assert time_span(crossing, date(2026, 10, 17), PARIS) == "ven. 16/10 22:00 au 01:00 (suite)"
    before = event(3, at(11, 22), at(12, 1))  # started before the range
    assert time_span(before, MONDAY, PARIS) == "dim. 11/10 22:00 au 01:00 (suite)"
    long = event(4, at(11, 22), at(13, 1))
    assert time_span(long, MONDAY, PARIS) == "dim. 11/10 22:00 au mar. 13/10 01:00 (suite)"
    assert time_span(long, date(2026, 10, 13), PARIS) == "dim. 11/10 22:00 au 01:00 (suite)"
    # across the DST change, seen from the day after
    dst = event(5, at(24), at(26), all_day=True)
    assert time_span(dst, date(2026, 10, 25), PARIS) == "journée (suite)"


@pytest.mark.parametrize(
    ("overrides", "expected"),
    [
        ({}, []),
        ({"busy_status": "oof"}, ["absent"]),
        ({"busy_status": "free"}, ["libre"]),
        ({"busy_status": "working_elsewhere"}, ["ailleurs"]),
        ({"response_status": "declined"}, ["refusé"]),
        ({"response_status": "not_responded"}, ["sans réponse"]),
        ({"response_status": "accepted"}, []),
        (
            {
                "busy_status": "tentative",
                "is_recurring": True,
                "online_meeting_url": "https://meet.example.com/abc",
            },
            ["provisoire", "récurrent", "visio"],
        ),
    ],
)
def test_event_flags(overrides: dict, expected: list[str]):
    assert event_flags(event(1, at(12, 14), at(12, 15), **overrides)) == expected


def test_agenda_helpers_accept_snapshot_like_events():
    snap = SimpleNamespace(
        start=at(12, 14),
        end=at(12, 15),
        all_day=False,
        busy_status="busy",
        response_status="accepted",
        is_recurring=False,
        has_online_meeting=True,
    )
    assert event_flags(snap) == ["visio"]
    assert time_span(snap, MONDAY, PARIS) == "14:00-15:00"
    assert group_by_day([snap], at(12), at(13), PARIS) == [(MONDAY, [snap])]


# ---------------------------------------------------------------------------- paths

HOME = r"Z:\profil-test"


@pytest.fixture
def profile(monkeypatch) -> None:
    monkeypatch.setenv("USERPROFILE", HOME)
    monkeypatch.setenv("APPDATA", HOME + r"\AppData\Roaming")
    monkeypatch.setenv("LOCALAPPDATA", HOME + "\\AppData\\Local\\")  # trailing separator


@pytest.mark.usefixtures("profile")
@pytest.mark.parametrize(
    ("path", "expected"),
    [
        (
            PureWindowsPath(HOME, "AppData", "Roaming", "outlook-bridge", "config.toml"),
            r"%APPDATA%\outlook-bridge\config.toml",
        ),
        (HOME + r"\AppData\Local\outlook-bridge\logs", r"%LOCALAPPDATA%\outlook-bridge\logs"),
        (HOME + r"\AppData\LocalLow\x", r"%USERPROFILE%\AppData\LocalLow\x"),
        (HOME + r"\Documents", r"%USERPROFILE%\Documents"),
        (HOME, "%USERPROFILE%"),
        (r"z:\PROFIL-TEST\appdata\roaming\x", r"%APPDATA%\x"),  # case-insensitive
        ("Z:/profil-test/AppData/Roaming/x", "%APPDATA%/x"),
        (r"Z:\profil-test2\x", r"Z:\profil-test2\x"),  # another folder sharing the prefix
        (r"G:\Mon Drive\outlook-bridge", r"G:\Mon Drive\outlook-bridge"),
    ],
)
def test_display_path(path, expected: str):
    assert display_path(path) == expected


def test_display_path_accepts_path_objects(profile):
    assert display_path(Path("relatif") / "x") == str(Path("relatif") / "x")


def test_longest_prefix_wins(monkeypatch, profile):
    # APPDATA nested under LOCALAPPDATA (pytest tmp dir): the longer, more specific one wins
    monkeypatch.setenv("APPDATA", HOME + r"\AppData\Local\Temp\pytest-0\cfg")
    message = f"Configuration invalide ({HOME}\\AppData\\Local\\Temp\\pytest-0\\cfg\\c.toml) : x"
    assert mask_paths(message) == r"Configuration invalide (%APPDATA%\c.toml) : x"


def test_mask_paths_in_messages(profile):
    message = (
        f"Configuration invalide ({HOME}\\AppData\\Roaming\\outlook-bridge\\config.toml) : "
        f"voir aussi '{HOME}\\\\AppData\\\\Local\\\\x' et {HOME.lower()}\\Bureau."
    )
    out = mask_paths(message)
    assert "profil-test" not in out.casefold()
    assert r"(%APPDATA%\outlook-bridge\config.toml)" in out
    assert r"'%LOCALAPPDATA%\\x'" in out
    assert out.endswith(r"%USERPROFILE%\Bureau.")


def test_mask_paths_without_profile_variables(monkeypatch):
    for name in fmt.PATH_VARS:
        monkeypatch.delenv(name, raising=False)
    assert mask_paths(r"Z:\profil-test\x") == r"Z:\profil-test\x"
    assert display_path(r"Z:\profil-test\x") == r"Z:\profil-test\x"


# ---------------------------------------------------------------------------- JSON


def test_to_json_is_ascii_and_round_trips():
    items = [mail(subject="Réunion « test » 🎉", received=NOW)]
    text = to_json(items)
    assert text.isascii()
    data = json.loads(text)
    assert data[0]["subject"] == "Réunion « test » 🎉"
    assert data[0]["untrusted_content"] is True
    assert data[0]["received"] == "2026-10-12T12:00:00+02:00"
    assert MailSummary.model_validate(data[0]) == items[0]
    assert "Réunion" in to_json(items[0], ascii=False)
    assert json.loads(to_json([])) == []


# ---------------------------------------------------------------------------- code pages


def _formatter_outputs() -> list[str]:
    outputs: list[str] = []
    outputs += [day_label(MONDAY + timedelta(days=i), year=True) for i in range(7)]
    outputs += [day_label(date(2026, m, 1)) for m in range(1, 13)]
    outputs += [short_day(MONDAY), short_day(MONDAY, year=True), short_dt(NOW, PARIS)]
    outputs += [long_dt(NOW, PARIS), range_label(at(12), at(20), PARIS)]
    outputs += [range_label(at(12, 14), at(13, 9), PARIS)]
    outputs += [duration(m) for m in (5, 60, 90)] + [size(n) for n in (5, 5000, 5 * 1024**2)]
    outputs += [ago(timedelta(seconds=s)) for s in (0, 300, 7200, 300000)]
    outputs += mail_flags(mail(unread=True, importance="high", has_attachments=True))
    outputs += mail_flags(mail(importance="low")) + [mail_marks(mail(unread=True))]
    for status in fmt.BUSY_FR:
        for response in ("declined", "not_responded", "accepted"):
            ev = event(1, at(12, 14), at(12, 15), busy_status=status, response_status=response)
            outputs += event_flags(ev.model_copy(update={"is_recurring": True}))
    for ev in agenda_events():
        outputs.append(time_span(ev, max(ev.start.date(), MONDAY), PARIS))
    outputs += [clean_line("Point projet"), clean_block("a\n\nb"), md_inline("a_b [c]", 5)]
    outputs += list(fmt.JOURS + fmt.JOURS_COURTS + fmt.MOIS)
    outputs += [*fmt.BUSY_FR.values(), *fmt.RESPONSE_FR.values(), *fmt.IMPORTANCE_FR.values()]
    return outputs


@pytest.mark.parametrize("codec", ["cp850", "cp1252"])
def test_formatter_output_fits_the_console_code_pages(codec: str):
    for text in _formatter_outputs():
        text.encode(codec)  # raises UnicodeEncodeError otherwise


@pytest.mark.parametrize("codec", ["cp850", "cp1252"])
def test_module_source_fits_the_console_code_pages(codec: str):
    Path(fmt.__file__).read_text(encoding="utf-8").encode(codec)
