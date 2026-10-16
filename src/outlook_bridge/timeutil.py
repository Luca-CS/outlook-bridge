"""Date normalisation and date formats for Outlook filters.

Phase 0 findings (docs/phase0-report.md):
- pywin32 returns the *local wall-clock* time labelled as UTC for ReceivedTime, Start, End
  and Table columns referenced by name.
- True UTC sources exist: StartUTC/EndUTC, PR_MESSAGE_DELIVERY_TIME, Table columns referenced
  by their DASL name. They are preferred because they are unambiguous across DST changes.
- Jet `Restrict` only parses `dd/MM/yyyy HH:mm` correctly on this locale; other formats are
  accepted silently and give wrong results. DASL takes `yyyy-MM-dd HH:mm` in UTC.

Phase 2 adds the user-facing date grammar (`parse_when`, design section 5). DST rules:
- day arithmetic is done on `date` objects, then `day_start` (local midnight);
- hour offsets are elapsed time, computed in UTC;
- local wall times go through `normalize` (UTC round trip): a nonexistent spring-gap time
  (2026-03-29 02:30) becomes 03:30+02:00, an ambiguous autumn time takes its first
  occurrence (fold=0, +02:00).
"""

from __future__ import annotations

import re
from datetime import UTC, date, datetime, time, timedelta, tzinfo
from typing import Literal, get_args
from zoneinfo import ZoneInfo

from outlook_bridge.errors import InvalidRequest

DEFAULT_TZ = ZoneInfo("Europe/Paris")
# Outlook uses 01/01/4501 as "no date".
NO_DATE_YEAR = 4501
JET_FORMAT = "%d/%m/%Y %H:%M"
DASL_FORMAT = "%Y-%m-%d %H:%M"

Role = Literal["since", "from", "to"]
ROLES: tuple[str, ...] = get_args(Role)
MAX_WHEN_CHARS = 40
MIN_YEAR, MAX_YEAR = 1900, 2999  # far from Outlook's 4501 "no date"
WHEN_FORMATS = (
    "now, today (auj), demain, hier, 2d, +7d, -3h, 16/10, 16/10/2026, 2026-10-16, "
    "heure optionnelle 14h30 ou 14:30"
)

_EPOCH = datetime(1970, 1, 1, tzinfo=UTC)
_ONE_US = timedelta(microseconds=1)
# Keyword -> day offset from today, None for the current instant.
_KEYWORDS: dict[str, int | None] = {
    "now": None,
    "maintenant": None,
    "today": 0,
    "auj": 0,
    "aujourd'hui": 0,
    "demain": 1,
    "tomorrow": 1,
    "hier": -1,
    "yesterday": -1,
}
# All patterns are used with fullmatch on the normalised (casefolded) text.
_OFFSET = re.compile(r"([+-]?)(\d{1,3})([djh])")
_ISO_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
_FR_DATE = re.compile(r"(\d{1,2})/(\d{1,2})(?:/(\d{4}))?")
_TIME = re.compile(r"(\d{1,2})(?::(\d{2})|h(\d{2})?)")
_DATE_TIME = re.compile(r"([\d/-]{3,10})(?:[ t]([\d:h]{2,5}))?")
# MCP tools (Phase 3): ISO 8601 only, extended format. An offset needs a time.
_ISO = re.compile(
    r"\d{4}-\d{2}-\d{2}"
    r"(?P<time>[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d{1,6})?)?(?:Z|[+-]\d{2}:\d{2})?)?",
    re.ASCII,
)
ISO_FORMATS = "2026-10-16 (jour entier) ou 2026-10-16T14:30 (heure de Paris si aucun fuseau)"


def to_aware(
    value: datetime | None, *, utc: bool = False, tz: tzinfo = DEFAULT_TZ
) -> datetime | None:
    """The only conversion point for COM dates.

    `utc=True` when the source is a true UTC value (StartUTC, DASL column...), otherwise the
    value is taken as local wall-clock time. pywin32's tzinfo is ignored in both cases since
    it is wrong for local values. Returns an aware datetime in `tz`, or None for "no date".
    """
    if value is None:
        return None
    naive = datetime(
        value.year,
        value.month,
        value.day,
        value.hour,
        value.minute,
        value.second,
        value.microsecond,
    )
    if naive.year >= NO_DATE_YEAR:
        return None
    if utc:
        return naive.replace(tzinfo=UTC).astimezone(tz)
    return naive.replace(tzinfo=tz)


def to_com(value: datetime, tz: tzinfo = DEFAULT_TZ) -> datetime:
    """Inverse of `to_aware` for COM arguments (e.g. GetOccurrence): pywin32 expects the local
    wall-clock time tagged UTC; a naive value would be shifted (Phase 1 probe)."""
    return ensure_aware(value, tz).replace(tzinfo=None).replace(tzinfo=UTC)


def normalize(value: datetime, tz: tzinfo = DEFAULT_TZ) -> datetime:
    """Round trip through UTC so the wall time shown matches the instant.

    A naive value is first read as local time in `tz`. A nonexistent spring-gap time
    (2026-03-29 02:30, fold=0) becomes 03:30+02:00; existing times are unchanged.
    """
    if value.tzinfo is None:
        value = value.replace(tzinfo=tz)
    return value.astimezone(UTC).astimezone(tz)


def ensure_aware(value: datetime, tz: tzinfo = DEFAULT_TZ) -> datetime:
    """User-supplied datetimes: naive means local time in `tz`; the result is normalised."""
    return normalize(value, tz)


def day_start(day: date, tz: tzinfo = DEFAULT_TZ) -> datetime:
    """Local midnight of `day` (always exists in Paris: transitions happen at 02:00/03:00)."""
    return normalize(datetime.combine(day, time(0), tzinfo=tz), tz)


def ceil_to(value: datetime, minutes: int = 15) -> datetime:
    """Round an aware datetime up to the next multiple of `minutes` (unchanged if aligned).

    Computed on the UTC timeline, so it stays correct across DST changes.
    """
    if value.tzinfo is None or value.utcoffset() is None:
        raise ValueError("ceil_to needs an aware datetime")
    if minutes < 1:
        raise ValueError("minutes must be >= 1")
    step = minutes * 60 * 1_000_000
    remainder = ((value - _EPOCH) // _ONE_US) % step
    if remainder == 0:
        return value
    return (value.astimezone(UTC) + (step - remainder) * _ONE_US).astimezone(value.tzinfo)


def parse_when(
    text: str,
    role: Role,
    *,
    now: datetime,
    label: str | None = None,
    after: datetime | None = None,
) -> datetime:
    """Parse a CLI date expression (design section 5) into an aware datetime in `now.tzinfo`.

    `role` sets the meaning of unsigned offsets (past for "since", future for "from"/"to")
    and of day values: a day gives its midnight for "since"/"from" and the next midnight for
    "to" (exclusive bound, so the day itself is included). An instant is the same for every
    role. Raises InvalidRequest (French message naming `label`, default `--<role>`).

    `after` ("to" only) is the resolved start of the range. A yearless `DD/MM` (with or
    without a time) then takes the occurrence nearest to that start instead of today, so
    both bounds land in the same year (`--from 01/04 --to 30/04` six months ahead) and
    `--from 20/12 --to 05/01` crosses New Year. For any range the core accepts (62 days at
    most) this is the first occurrence whose bound is after `after`; a reversed input stays
    reversed and gets the core's "end before start" error, not a range one year long.
    """
    if role not in ROLES:
        raise ValueError(f"unknown role: {role!r}")
    if now.tzinfo is None or now.utcoffset() is None:
        raise ValueError("parse_when needs an aware `now`")
    if after is not None and (role != "to" or after.tzinfo is None or after.utcoffset() is None):
        raise ValueError("`after` must be aware and only applies to role 'to'")
    tz = now.tzinfo
    label = label or f"--{role}"
    shown = _shown(text)
    anchor = after.astimezone(tz).date() if after is not None else None
    value = _parse_value(_normalize_text(text), role, now, anchor)
    if value is None:
        raise InvalidRequest(f"Date invalide pour {label} : « {shown} ». Formats : {WHEN_FORMATS}.")
    if isinstance(value, datetime):
        result = normalize(value, tz)
    else:
        result = day_start(value + timedelta(days=1) if role == "to" else value, tz)
    # Compare instants: same-tzinfo comparison would ignore the offset in the repeated hour.
    if role == "since" and result.astimezone(UTC) > now.astimezone(UTC):
        raise InvalidRequest(
            f"{label} désigne une date future (« {shown} ») : utilisez 2d pour « il y a 2 jours »."
        )
    return result


def parse_iso(
    text: str, role: Role, *, tz: tzinfo = DEFAULT_TZ, label: str | None = None
) -> datetime:
    """Parse an MCP tool date: ISO 8601 only (Phase 3 decision), aware in `tz`.

    A day alone resolves by role like `parse_when`: its midnight for "since"/"from", the next
    midnight for "to" (exclusive bound, so the day itself is included). A time without offset
    is local time in `tz`; an explicit offset or Z is honoured. No relative grammar: the model
    knows the date, and "2d" or "16/10" would be ambiguous. Raises InvalidRequest.
    """
    if role not in ROLES:
        raise ValueError(f"unknown role: {role!r}")
    label = label or role
    candidate = text.strip()
    match = _ISO.fullmatch(candidate) if len(candidate) <= MAX_WHEN_CHARS else None
    try:
        parsed = datetime.fromisoformat(candidate) if match else None
    except ValueError:
        parsed = None
    if parsed is None or not MIN_YEAR <= parsed.year <= MAX_YEAR:
        raise InvalidRequest(
            f"Date invalide pour {label} : « {_shown(text)} ». Format ISO 8601 : {ISO_FORMATS}."
        )
    if match["time"] is None:
        day = parsed.date()
        return day_start(day + timedelta(days=1) if role == "to" else day, tz)
    return normalize(parsed, tz)


def _normalize_text(text: str) -> str:
    """strip + casefold, typographic apostrophe -> ', any whitespace run (NBSP...) -> ' '."""
    return " ".join(text.replace("\u2019", "'").casefold().split())


def _shown(text: str) -> str:
    """The user's input, echoed safely in an error message."""
    shown = " ".join(text.split())
    shown = "".join(c if c.isprintable() else "?" for c in shown)
    return shown if len(shown) <= MAX_WHEN_CHARS else shown[:MAX_WHEN_CHARS] + "..."


def _parse_value(
    text: str, role: Role, now: datetime, anchor: date | None = None
) -> date | datetime | None:
    """A `date` (day value, resolved by role), an aware or local `datetime` (instant) or None.

    `anchor` is the day a yearless `DD/MM` is resolved around (default: today)."""
    if not text or len(text) > MAX_WHEN_CHARS:
        return None
    today = now.date()
    if text in _KEYWORDS:
        days = _KEYWORDS[text]
        return now if days is None else today + timedelta(days=days)

    if match := _OFFSET.fullmatch(text):
        sign_char, digits, unit = match.groups()
        if sign_char:
            sign = 1 if sign_char == "+" else -1
        else:
            sign = -1 if role == "since" else 1
        amount = sign * int(digits)
        if unit == "h":  # elapsed time: computed in UTC, correct across DST
            return (now.astimezone(UTC) + timedelta(hours=amount)).astimezone(now.tzinfo)
        return today + timedelta(days=amount)

    if match := _DATE_TIME.fullmatch(text):
        day = _parse_date(match[1], anchor or today)
        if day is not None:
            if match[2] is None:
                return day
            at = _parse_time(match[2])
            if at is not None:
                return datetime.combine(day, at, tzinfo=now.tzinfo)

    if (at := _parse_time(text)) is not None:
        return datetime.combine(today, at, tzinfo=now.tzinfo)

    try:  # ISO 8601 with an explicit offset or Z ("2026-10-15T14:00+00:00")
        parsed = datetime.fromisoformat(text.upper())
    except ValueError:
        return None
    if parsed.utcoffset() is None or not MIN_YEAR <= parsed.year <= MAX_YEAR:
        return None
    return parsed


def _parse_date(token: str, anchor: date) -> date | None:
    """`YYYY-MM-DD`, `DD/MM/YYYY`, or `DD/MM` (occurrence nearest to `anchor`, ties go to the
    future)."""
    try:
        if match := _ISO_DATE.fullmatch(token):
            year, month, day = (int(g) for g in match.groups())
        elif match := _FR_DATE.fullmatch(token):
            day, month = int(match[1]), int(match[2])
            if match[3] is None:
                return _nearest(day, month, anchor)
            year = int(match[3])
        else:
            return None
        if not MIN_YEAR <= year <= MAX_YEAR:
            return None
        return date(year, month, day)
    except ValueError:
        return None


def _nearest(day: int, month: int, anchor: date) -> date | None:
    candidates = []
    for year in (anchor.year - 1, anchor.year, anchor.year + 1):
        try:
            if MIN_YEAR <= year <= MAX_YEAR:
                candidates.append(date(year, month, day))
        except ValueError:
            continue
    if not candidates:
        return None
    return min(candidates, key=lambda c: (abs((c - anchor).days), c < anchor))


def _parse_time(token: str) -> time | None:
    """`HH:MM`, `HHhMM` or `HHh` (hours on 1 or 2 digits)."""
    if not (match := _TIME.fullmatch(token)):
        return None
    try:
        return time(int(match[1]), int(match[2] or match[3] or 0))
    except ValueError:
        return None


def jet_datetime(value: datetime, tz: tzinfo = DEFAULT_TZ) -> str:
    """Local-time literal for a Jet `Restrict` filter (`[Start] < '...'`)."""
    return ensure_aware(value, tz).strftime(JET_FORMAT)


def dasl_datetime(value: datetime, tz: tzinfo = DEFAULT_TZ) -> str:
    """UTC literal for a DASL `@SQL=` filter."""
    return ensure_aware(value, tz).astimezone(UTC).strftime(DASL_FORMAT)


def parse_working_hours(spec: str) -> tuple[time, time]:
    """`"09:00-19:00"` -> (09:00, 19:00)."""
    try:
        first, last = (time.fromisoformat(part.strip()) for part in spec.split("-"))
    except ValueError as exc:
        raise InvalidRequest(f"Plage horaire invalide : {spec!r} (attendu HH:MM-HH:MM).") from exc
    if first >= last:
        raise InvalidRequest(f"Plage horaire invalide : {spec!r} (début après la fin).")
    return first, last
