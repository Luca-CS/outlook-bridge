"""Date normalisation and date formats for Outlook filters.

Phase 0 findings (docs/phase0-report.md):
- pywin32 returns the *local wall-clock* time labelled as UTC for ReceivedTime, Start, End
  and Table columns referenced by name.
- True UTC sources exist: StartUTC/EndUTC, PR_MESSAGE_DELIVERY_TIME, Table columns referenced
  by their DASL name. They are preferred because they are unambiguous across DST changes.
- Jet `Restrict` only parses `dd/MM/yyyy HH:mm` correctly on this locale; other formats are
  accepted silently and give wrong results. DASL takes `yyyy-MM-dd HH:mm` in UTC.
"""

from __future__ import annotations

from datetime import UTC, datetime, time, tzinfo
from zoneinfo import ZoneInfo

from outlook_bridge.errors import InvalidRequest

DEFAULT_TZ = ZoneInfo("Europe/Paris")
# Outlook uses 01/01/4501 as "no date".
NO_DATE_YEAR = 4501
JET_FORMAT = "%d/%m/%Y %H:%M"
DASL_FORMAT = "%Y-%m-%d %H:%M"


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


def ensure_aware(value: datetime, tz: tzinfo = DEFAULT_TZ) -> datetime:
    """User-supplied datetimes: naive means local time in `tz`."""
    return value.replace(tzinfo=tz) if value.tzinfo is None else value.astimezone(tz)


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
