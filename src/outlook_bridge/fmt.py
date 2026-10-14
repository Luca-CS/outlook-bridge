"""French, locale-independent formatting and sanitising of untrusted text.

Shared by the CLI and the snapshot export. Pure module: no I/O, except reading the profile
environment variables used to mask paths. No rich, no `locale`, no `strftime` names.

Our own literals stay within cp1252 and cp850 (pipes and consoles of a French Windows): no
typographic dash, ellipsis, arrow or curly apostrophe; use "-", "...", "·", "« »" and "'".
"""

from __future__ import annotations

import json
import os
import re
from collections.abc import Sequence
from datetime import date, datetime, time, timedelta, tzinfo
from pathlib import Path
from typing import Protocol

from pydantic_core import to_jsonable_python

from outlook_bridge.models import MailSummary
from outlook_bridge.timeutil import ensure_aware

JOURS: tuple[str, ...] = (
    "lundi",
    "mardi",
    "mercredi",
    "jeudi",
    "vendredi",
    "samedi",
    "dimanche",
)
JOURS_COURTS: tuple[str, ...] = ("lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim.")
MOIS: tuple[str, ...] = (
    "janvier",
    "février",
    "mars",
    "avril",
    "mai",
    "juin",
    "juillet",
    "août",
    "septembre",
    "octobre",
    "novembre",
    "décembre",
)
BUSY_FR = {
    "free": "libre",
    "tentative": "provisoire",
    "busy": "occupé",
    "oof": "absent",
    "working_elsewhere": "ailleurs",
}
RESPONSE_FR = {"declined": "refusé", "not_responded": "sans réponse"}
IMPORTANCE_FR = {"low": "importance faible", "high": "importance haute"}

ELLIPSIS = "..."
PATH_VARS = ("LOCALAPPDATA", "APPDATA", "USERPROFILE")

# ---------------------------------------------------------------------------- sanitising

# Terminal escape sequences, removed whole so that their parameters (colours, OSC 8 URLs,
# window titles) do not leak into the text. String bodies exclude every introducer and
# terminator, so the scan stays linear even on hostile input.
_ESCAPES = re.compile(
    r"(?:\x1b\[|\x9b)[0-?]*[ -/]*[@-~]"  # CSI: colours, cursor moves, erase
    r"|(?:\x1b[\]PX^_]|[\x90\x98\x9d\x9e\x9f])"  # OSC, DCS, SOS, PM, APC ...
    r"[^\x07\x1b\x90\x98\x9c\x9d\x9e\x9f]*(?:\x07|\x1b\\|\x9c)"  # ... up to BEL or ST
    r"|\x1b[ -/]*[0-~]"  # other ESC sequences (charset, reset...)
)
# Controls and invisible characters: C0 (except the whitespace ones), DEL, C1 (except NEL),
# bidi controls, BOM, lone surrogates and Unicode tags (invisible "ASCII smuggling").
_INVISIBLE = re.compile(
    r"[\x00-\x08\x0e-\x1f\x7f-\x84\x86-\x9f"
    r"\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff\ud800-\udfff"
    r"\U000e0000-\U000e007f]"
)
_WHITESPACE = re.compile(r"\s+")
_LINE_BREAKS = re.compile(r"\r\n?|[\x0b\x0c\x85\u2028\u2029]")
_INLINE_SPACE = re.compile(r"[^\S\n]")
_BLANK_RUNS = re.compile(r"\n{4,}")
_MD_SPECIAL = re.compile(r"([\\`*_\[\]<>|])")


def _strip_controls(text: str) -> str:
    return _INVISIBLE.sub("", _ESCAPES.sub("", text))


def _cut(text: str, max_len: int | None) -> str:
    if max_len is None or len(text) <= max_len:
        return text
    if max_len <= len(ELLIPSIS):
        return text[:max_len]
    return text[: max_len - len(ELLIPSIS)].rstrip() + ELLIPSIS


def clean_line(text: str | None, max_len: int | None = None) -> str:
    """Untrusted text on one line: no control, escape sequence or bidi override left; any
    whitespace run (line breaks included) becomes one space; cut with "..." to `max_len`."""
    if not text:
        return ""
    return _cut(_WHITESPACE.sub(" ", _strip_controls(text)).strip(), max_len)


def clean_block(text: str | None) -> str:
    """Untrusted multi-line text (mail body): same cleaning, but line breaks are kept
    (normalised to LF), tabs become 4 spaces, trailing spaces are dropped and blank runs are
    capped at 2 lines."""
    if not text:
        return ""
    text = _strip_controls(_LINE_BREAKS.sub("\n", text))
    text = _INLINE_SPACE.sub(" ", text.replace("\t", "    "))
    text = "\n".join(line.rstrip() for line in text.split("\n"))
    return _BLANK_RUNS.sub("\n\n\n", text).strip("\n")


def md_inline(text: str | None, max_len: int | None = None) -> str:
    """Untrusted text for an inline Markdown context: `clean_line`, then the characters that
    could build a link, an image, raw HTML, code or a table cell are backslash-escaped. The cut
    happens before escaping, so an escape is never split."""
    return _MD_SPECIAL.sub(r"\\\1", clean_line(text, max_len))


# ---------------------------------------------------------------------------- dates


def _hm(value: datetime) -> str:
    return f"{value.hour:02d}:{value.minute:02d}"


def _is_midnight(value: datetime) -> bool:
    return value.time() == time(0)


def _last_day(end: datetime) -> date:
    """Last local day touched by a range ending at `end` (exclusive bound)."""
    return end.date() - timedelta(days=1) if _is_midnight(end) else end.date()


def day_label(d: date, *, year: bool = False) -> str:
    """`lundi 12 octobre`, `samedi 1er août 2026`."""
    label = f"{JOURS[d.weekday()]} {'1er' if d.day == 1 else d.day} {MOIS[d.month - 1]}"
    return f"{label} {d.year}" if year else label


def short_day(d: date, *, year: bool = False) -> str:
    """`lun. 12/10`, `lun. 12/10/2026`."""
    label = f"{JOURS_COURTS[d.weekday()]} {d.day:02d}/{d.month:02d}"
    return f"{label}/{d.year:04d}" if year else label


def short_dt(dt: datetime, tz: tzinfo) -> str:
    """`lun. 12/10 11:00` (local time in `tz`; naive means already local)."""
    local = ensure_aware(dt, tz)
    return f"{short_day(local.date())} {_hm(local)}"


def long_dt(dt: datetime, tz: tzinfo) -> str:
    """`lundi 12 octobre 2026 à 11:00`."""
    local = ensure_aware(dt, tz)
    return f"{day_label(local.date(), year=True)} à {_hm(local)}"


def range_label(start: datetime, end: datetime, tz: tzinfo) -> str:
    """Header of a half-open range [start, end).

    End at local midnight: day-level, with the inclusive last day (`du lun. 12/10 au
    lun. 19/10/2026 inclus`); a start that is not at midnight keeps its time. Otherwise:
    `du lun. 12/10 14:00 au mar. 13/10/2026 09:00`. The start shows its year only when it
    differs from the end's.
    """
    first, last = ensure_aware(start, tz), ensure_aware(end, tz)
    if _is_midnight(last):
        last_day = last.date() - timedelta(days=1)
        label = short_day(first.date(), year=first.year != last_day.year)
        if not _is_midnight(first):
            label = f"{label} {_hm(first)}"
        return f"du {label} au {short_day(last_day, year=True)} inclus"
    label = short_day(first.date(), year=first.year != last.year)
    return f"du {label} {_hm(first)} au {short_day(last.date(), year=True)} {_hm(last)}"


def duration(minutes: int) -> str:
    """`45 min`, `1 h`, `1 h 30`, `1 h 05`, `10 h`."""
    minutes = max(0, int(minutes))
    if minutes < 60:
        return f"{minutes} min"
    hours, rest = divmod(minutes, 60)
    return f"{hours} h {rest:02d}" if rest else f"{hours} h"


def _decimal(value: float) -> str:
    """One French decimal below 10 (`3,4`), whole number above; a trailing `,0` is dropped."""
    if value >= 10:
        return str(int(value + 0.5))
    text = f"{value:.1f}".replace(".", ",")
    return text.removesuffix(",0")


def size(n_bytes: int) -> str:
    """`850 o`, `12 Ko`, `3,4 Mo`, `1,2 Go` (binary multiples, French units)."""
    n_bytes = max(0, int(n_bytes))
    if n_bytes < 1024:
        return f"{n_bytes} o"
    if n_bytes < 1024**2:
        return f"{int(n_bytes / 1024 + 0.5)} Ko"
    if n_bytes < 1024**3:
        return f"{_decimal(n_bytes / 1024**2)} Mo"
    return f"{_decimal(n_bytes / 1024**3)} Go"


def ago(delta: timedelta) -> str:
    """`à l'instant`, `il y a 4 min`, `il y a 2 h`, `il y a 3 j` (rounded down; a negative
    delta, i.e. a clock moved back, reads as `à l'instant`)."""
    seconds = int(delta.total_seconds())
    if seconds < 60:
        return "à l'instant"
    minutes = seconds // 60
    if minutes < 60:
        return f"il y a {minutes} min"
    hours = minutes // 60
    if hours < 24:
        return f"il y a {hours} h"
    return f"il y a {hours // 24} j"


# ---------------------------------------------------------------------------- mails


def mail_marks(mail: MailSummary) -> str:
    """Compact state column: subset of `*!+` in that order (non lu, haute, pièce jointe)."""
    marks = "*" if mail.unread else ""
    marks += "!" if mail.importance == "high" else ""
    return marks + ("+" if mail.has_attachments else "")


def mail_flags(mail: MailSummary) -> list[str]:
    """Readable state: `["non lu", "importance haute", "pièce jointe"]` (subset)."""
    flags = ["non lu"] if mail.unread else []
    if importance := IMPORTANCE_FR.get(mail.importance):
        flags.append(importance)
    if mail.has_attachments:
        flags.append("pièce jointe")
    return flags


# ---------------------------------------------------------------------------- agenda


class EventLike(Protocol):
    """Fields read by the agenda helpers: `Event` and the snapshot event model both qualify.
    The online meeting is read from `online_meeting_url` or `has_online_meeting`."""

    @property
    def start(self) -> datetime: ...

    @property
    def end(self) -> datetime: ...

    @property
    def all_day(self) -> bool: ...

    @property
    def busy_status(self) -> str: ...

    @property
    def response_status(self) -> str: ...

    @property
    def is_recurring(self) -> bool: ...


def group_by_day[E: EventLike](
    events: Sequence[E], start: datetime, end: datetime, tz: tzinfo
) -> list[tuple[date, list[E]]]:
    """Every local day touching [start, end), each with its events sorted by start.

    An event is listed once: under its local start day, or under the first day when it
    started before the range. Days without events are kept (empty list), so a reader never
    mistakes a missing day for a free one. Events outside the range are ignored.
    """
    first, last = ensure_aware(start, tz), ensure_aware(end, tz)
    if last <= first:
        return []
    days: dict[date, list[E]] = {}
    day, last_day = first.date(), _last_day(last)
    while day <= last_day:
        days[day] = []
        day += timedelta(days=1)
    for event in sorted(events, key=lambda e: ensure_aware(e.start, tz)):
        ev_start, ev_end = ensure_aware(event.start, tz), ensure_aware(event.end, tz)
        if ev_start >= last or (ev_start < first and ev_end <= first):
            continue
        bucket = days.get(max(ev_start.date(), first.date()))
        if bucket is not None:
            bucket.append(event)
    return list(days.items())


def time_span(event: EventLike, day: date, tz: tzinfo) -> str:
    """When an event happens, seen from the day it is listed under.

    All-day: `journée` or `journées jusqu'au mer. 14/10`. Timed on that day: `14:00-15:00`.
    Otherwise each end that falls on another day carries its date: `22:00 au sam. 17/10 01:00`,
    `dim. 11/10 22:00 au 01:00`.
    """
    ev_start, ev_end = ensure_aware(event.start, tz), ensure_aware(event.end, tz)
    if event.all_day:
        last = _last_day(ev_end)
        if last <= ev_start.date():
            return "journée"
        return f"journées jusqu'au {short_day(last)}"
    if ev_start.date() == day == ev_end.date():
        return f"{_hm(ev_start)}-{_hm(ev_end)}"
    begin = _hm(ev_start) if ev_start.date() == day else short_dt(ev_start, tz)
    finish = _hm(ev_end) if ev_end.date() == day else short_dt(ev_end, tz)
    return f"{begin} au {finish}"


def event_flags(event: EventLike) -> list[str]:
    """Notable traits, in this order: availability unless busy, a declined or unanswered
    invitation, `récurrent`, `visio`."""
    flags = []
    if event.busy_status != "busy":
        flags.append(BUSY_FR.get(event.busy_status, event.busy_status))
    if response := RESPONSE_FR.get(event.response_status):
        flags.append(response)
    if event.is_recurring:
        flags.append("récurrent")
    if getattr(event, "online_meeting_url", None) or getattr(event, "has_online_meeting", False):
        flags.append("visio")
    return flags


# ---------------------------------------------------------------------------- paths


def _path_pattern() -> re.Pattern[str] | None:
    """Profile folders from the environment, longest first, matched case-insensitively on
    whole path segments, with `\\`, `/` or doubled separators (repr of a path)."""
    prefixes: list[tuple[str, str]] = []
    for name in PATH_VARS:
        parts = [p for p in re.split(r"[\\/]+", os.environ.get(name, "").strip()) if p]
        if parts:
            prefixes.append((r"[\\/]+".join(re.escape(p) for p in parts), name))
    if not prefixes:
        return None
    prefixes.sort(key=lambda item: len(item[0]), reverse=True)
    alternatives = "|".join(f"(?P<{name}>{regex})" for regex, name in prefixes)
    return re.compile(rf"(?<![\w-])(?:{alternatives})(?![\w-])", re.IGNORECASE)


def mask_paths(text: str) -> str:
    """Replaces %LOCALAPPDATA%, %APPDATA% and %USERPROFILE% anywhere in a message, so that no
    user name reaches the output (error messages embed paths)."""
    pattern = _path_pattern()
    if pattern is None or not text:
        return text
    return pattern.sub(lambda m: f"%{m.lastgroup}%", text)


def display_path(path: Path | str) -> str:
    """A path for display: its profile prefix becomes `%LOCALAPPDATA%`, `%APPDATA%` or
    `%USERPROFILE%`. Other paths (a synced drive...) are shown as is."""
    return mask_paths(str(path))


# ---------------------------------------------------------------------------- JSON


def to_json(obj: object, *, ascii: bool = True) -> str:
    """Models (or lists of models) as indented JSON. `ascii=True` escapes every non-ASCII
    character, which survives any console code page and PowerShell 5.1 pipes."""
    return json.dumps(to_jsonable_python(obj), ensure_ascii=ascii, indent=2)
