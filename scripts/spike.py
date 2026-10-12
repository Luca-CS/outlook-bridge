"""Phase 0 feasibility spike: classic Outlook via COM, read-only.

Usage:
    uv run python scripts/spike.py           # masked output (safe to paste / share)
    uv run python scripts/spike.py --show    # real subjects / names, local eyes only

Never commit the output of this script (see .gitignore and CLAUDE.md).
Read-only: no item is modified. The only property assignment is
`Items.IncludeRecurrences` on the calendar *collection*, as required by the spec (§7.5).
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pythoncom
import pywintypes
import win32com.client

PARIS = ZoneInfo("Europe/Paris")
OL_FOLDER_INBOX = 6
OL_FOLDER_CALENDAR = 9
PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
SAFETY_LIMIT = 2000
SLOW_CALL_S = 2.0  # a COM call this slow likely waited on an Object Model Guard prompt

SHOW = False


def mask(value: object) -> str:
    s = "" if value is None else str(value)
    if SHOW or not s:
        return s
    return f"{s[0]}…({len(s)} ch)"


def mask_addr(addr: object) -> str:
    s = "" if addr is None else str(addr)
    if SHOW or not s:
        return s
    if s.upper().startswith("/O="):
        return "<X500 address>"
    if "@" in s:
        tld = s.rsplit(".", 1)[-1] if "." in s.split("@", 1)[1] else "?"
        return f"{s[0]}***@***.{tld}"
    return mask(s)


def section(title: str) -> None:
    print(f"\n=== {title} ===")


def timed(label: str, fn):
    t0 = time.perf_counter()
    try:
        result = fn()
    except pywintypes.com_error as exc:
        dt = time.perf_counter() - t0
        msg = exc.excepinfo[2] if exc.excepinfo else exc
        print(f"  [{label}] COM error after {dt:.2f}s: {msg}")
        return None
    dt = time.perf_counter() - t0
    if dt > SLOW_CALL_S:
        print(f"  [{label}] SLOW ({dt:.1f}s): possible security prompt, did a dialog appear?")
    return result


def iterate(items, limit: int = SAFETY_LIMIT):
    """GetFirst/GetNext with a safety cap (never trust items.Count with recurrences)."""
    item = items.GetFirst()
    n = 0
    while item is not None and n < limit:
        yield item
        n += 1
        item = items.GetNext()
    if n >= limit:
        print(f"  ! safety limit {limit} reached")


def describe_dt(dt) -> str:
    return f"{dt.isoformat()} tz={dt.tzinfo!r}"


def naive(dt) -> datetime:
    return datetime(dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


# --------------------------------------------------------------------------- checks


def check_environment() -> None:
    section("Environment")
    print(f"  Python {sys.version.split()[0]} ({'64' if sys.maxsize > 2**32 else '32'}-bit)")
    procs = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"], capture_output=True, text=True, check=False
    ).stdout.lower()
    print(f"  OUTLOOK.EXE running (classic): {'outlook.exe' in procs}")
    print(f"  olk.exe running (new Outlook): {'olk.exe' in procs}")


def connect():
    section("Connection")
    app = timed("Dispatch", lambda: win32com.client.Dispatch("Outlook.Application"))
    if app is None:
        print("  NO-GO: cannot reach Outlook.Application (new Outlook? elevation mismatch?)")
        sys.exit(1)
    print(f"  Outlook version: {app.Version}")
    ns = app.GetNamespace("MAPI")
    mode = timed("ExchangeConnectionMode", lambda: ns.ExchangeConnectionMode)
    print(f"  ExchangeConnectionMode: {mode}  (700=cached connected full, 800=online)")
    return app, ns


def check_accounts(ns) -> None:
    section("Accounts and stores")
    for acc in iterate(_Seq(ns.Accounts)):
        smtp = timed("Account.SmtpAddress", lambda a=acc: a.SmtpAddress)
        print(f"  account: {mask(acc.DisplayName)}  smtp={mask_addr(smtp)}  type={acc.AccountType}")
    for store in iterate(_Seq(ns.Stores)):
        cached = timed("IsCachedExchange", lambda s=store: s.IsCachedExchange)
        inbox = timed("Store.GetDefaultFolder(Inbox)", lambda s=store: s.GetDefaultFolder(6))
        inbox_info = f"inbox items={inbox.Items.Count}" if inbox is not None else "no inbox"
        print(
            f"  store: {mask(store.DisplayName)}  exchange_type={store.ExchangeStoreType}"
            f"  cached={cached}  store_id_len={len(store.StoreID)}  {inbox_info}"
        )


class _Seq:
    """Adapter so 1-based COM collections fit iterate()."""

    def __init__(self, coll):
        self._coll, self._i = coll, 0

    def GetFirst(self):
        self._i = 0
        return self.GetNext()

    def GetNext(self):
        self._i += 1
        return self._coll.Item(self._i) if self._i <= self._coll.Count else None


def sender_smtp(mail) -> tuple[str, str]:
    etype = mail.SenderEmailType
    if etype == "EX":
        sender = timed("Sender", lambda: mail.Sender)
        user = timed("GetExchangeUser", lambda: sender.GetExchangeUser()) if sender else None
        if user is not None:
            return etype, timed("PrimarySmtpAddress", lambda: user.PrimarySmtpAddress) or ""
        return etype, f"(fallback SenderName) {mail.SenderName}"
    return etype, timed("SenderEmailAddress", lambda: mail.SenderEmailAddress) or ""


def check_inbox(ns) -> None:
    section("Inbox: 5 most recent (subject, sender, date only)")
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    for i, mail in enumerate(iterate(items, 5), 1):
        if mail.Class != 43:  # olMail
            print(f"  #{i} non-mail item class={mail.Class}")
            continue
        etype, smtp = sender_smtp(mail)
        recv = mail.ReceivedTime
        pa = mail.PropertyAccessor
        utc = timed("delivery time", lambda p=pa: p.GetProperty(PR_MESSAGE_DELIVERY_TIME))
        delta = (naive(recv) - naive(utc)) if utc is not None else None
        print(
            f"  #{i} subj={mask(mail.Subject)}  from={mask(mail.SenderName)}"
            f"  type={etype}  smtp={mask_addr(smtp)}"
        )
        print(f"      ReceivedTime={describe_dt(recv)}")
        print(f"      delivery(UTC)={describe_dt(utc) if utc else None}  wallclock-utc={delta}")
        if i == 1:
            n = timed("Body length (guard probe)", lambda m=mail: len(m.Body or ""))
            print(f"      body length={n}  (content not read out)")


def ground_truth_since(items, threshold: datetime) -> int:
    items.Sort("[ReceivedTime]", True)
    n = 0
    for mail in iterate(items):
        if naive(mail.ReceivedTime) < threshold:
            break
        n += 1
    return n


def check_restrict_formats(ns) -> str | None:
    section("Restrict: date format on this locale")
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    threshold = (datetime.now() - timedelta(days=3)).replace(hour=0, minute=0, second=0)
    truth = ground_truth_since(inbox.Items, threshold)
    print(f"  threshold (local) {threshold:%Y-%m-%d %H:%M}, ground truth by iteration: {truth}")

    jet_formats = {
        "dd/MM/yyyy HH:mm": "%d/%m/%Y %H:%M",
        "MM/dd/yyyy HH:mm": "%m/%d/%Y %H:%M",
        "MM/dd/yyyy hh:mm AM": "%m/%d/%Y %I:%M %p",
        "yyyy-MM-dd HH:mm": "%Y-%m-%d %H:%M",
    }
    winner = None
    for label, fmt in jet_formats.items():
        flt = f"[ReceivedTime] >= '{threshold.strftime(fmt)}'"
        n = _count_restrict(inbox.Items, flt)
        ok = n == truth
        winner = winner or (fmt if ok else None)
        print(f"  Jet  {label:<22} -> {n!s:>5}  {'OK' if ok else ''}   {flt}")

    utc_threshold = threshold.replace(tzinfo=PARIS).astimezone(ZoneInfo("UTC"))
    for label, value in {
        "DASL UTC": utc_threshold.strftime("%Y-%m-%d %H:%M"),
        "DASL local (wrong?)": threshold.strftime("%Y-%m-%d %H:%M"),
    }.items():
        flt = f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{value}'"
        n = _count_restrict(inbox.Items, flt)
        print(f"  {label:<27} -> {n!s:>5}  {'OK' if n == truth else ''}   {flt}")
    return winner


def _count_restrict(items, flt: str):
    try:
        restricted = items.Restrict(flt)
    except pywintypes.com_error as exc:
        return f"ERR {exc.excepinfo[2] if exc.excepinfo else exc}"[:60]
    return sum(1 for _ in iterate(restricted))


def check_table(ns) -> None:
    section("Folder.GetTable: date columns")
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    table = timed("GetTable", lambda: inbox.GetTable())
    if table is None:
        return
    cols = table.Columns
    wanted = [
        "ReceivedTime",
        "urn:schemas:httpmail:datereceived",
        "SenderName",
        "UnRead",
        "ConversationID",
        "http://schemas.microsoft.com/mapi/proptag/0x30130102",  # PR_CONVERSATION_ID
        "ConversationTopic",
        "SenderEmailType",
    ]
    added = []
    for name in wanted:
        try:
            cols.Add(name)
            added.append(name)
            print(f"  column OK      {name}")
        except pywintypes.com_error:
            print(f"  column REFUSED {name}")
    table.Sort("ReceivedTime", True)
    if table.EndOfTable:
        print("  empty inbox")
        return
    row = table.GetNextRow()
    by_name = row.Item("ReceivedTime")
    by_dasl = row.Item("urn:schemas:httpmail:datereceived")
    item = ns.GetItemFromID(row.Item("EntryID"), inbox.StoreID)
    item_recv = item.ReceivedTime
    print(f"  table 'ReceivedTime'          = {describe_dt(by_name)}")
    print(f"  table 'urn:...datereceived'   = {describe_dt(by_dasl)}")
    print(f"  item.ReceivedTime (same item) = {describe_dt(item_recv)}")
    print(f"  name-dasl={naive(by_name) - naive(by_dasl)}")
    print(f"  name-item={naive(by_name) - naive(item_recv)}")
    for name in added[2:]:
        value = row.Item(name)
        kind = type(value).__name__
        size = len(value) if isinstance(value, (str, bytes, memoryview)) else value
        print(f"  {name[-30:]:<30} type={kind} len/value={size if kind != 'str' else len(value)}")
    conv_tag = bytes(row.Item("http://schemas.microsoft.com/mapi/proptag/0x30130102")).hex().upper()
    print(f"  PR_CONVERSATION_ID hex == item.ConversationID: {conv_tag == item.ConversationID}")
    print("  GetItemFromID(entry_id, store_id): OK")


def check_calendar(ns, fmt: str | None, days: int = 7) -> None:
    section(f"Calendar: next {days} days, recurrences expanded")
    fmt = fmt or "%d/%m/%Y %H:%M"
    start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end = start + timedelta(days=days)
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items
    items.Sort("[Start]")
    items.IncludeRecurrences = True  # collection setting, not an item write
    flt = f"[Start] < '{end.strftime(fmt)}' AND [End] > '{start.strftime(fmt)}'"
    print(f"  filter: {flt}")
    restricted = items.Restrict(flt)
    n = recurring = 0
    for ev in iterate(restricted, 500):
        n += 1
        recurring += bool(ev.IsRecurring)
        s, s_utc = ev.Start, ev.StartUTC
        if n <= 15:
            print(
                f"  {naive(s):%a %d/%m %H:%M}  allday={ev.AllDayEvent}  recurring={ev.IsRecurring}"
                f"  busy={ev.BusyStatus}  subj={mask(ev.Subject)}"
            )
            print(
                f"      Start tz={s.tzinfo!r}  StartUTC={describe_dt(s_utc)}"
                f"  Start-StartUTC={naive(s) - naive(s_utc)}"
            )
    print(f"  total={n}  recurring occurrences={recurring}")
    if recurring == 0 and days < 31:
        print("  no recurring occurrence found, widening the window")
        check_calendar(ns, fmt, 31)
    elif days >= 31 and n == 0:
        diagnose_calendars(ns)


def diagnose_calendars(ns) -> None:
    section("Calendar diagnostic (no event found)")
    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    raw = cal.Items  # no recurrence expansion: Count is reliable here
    print(f"  default calendar: {mask(cal.Name)}  raw items={raw.Count}")
    raw.Sort("[Start]", True)
    latest = None
    for i, ev in enumerate(iterate(raw, 3), 1):
        latest = latest or naive(ev.Start)
        print(f"  latest #{i}: start={naive(ev.Start):%Y-%m-%d}  recurring={ev.IsRecurring}")
    masters = [ev for ev in iterate(cal.Items) if ev.IsRecurring]
    for ev in masters:
        pattern = ev.GetRecurrencePattern()
        print(
            f"  recurring series: first={naive(ev.Start):%Y-%m-%d}"
            f"  pattern_end={naive(pattern.PatternEndDate):%Y-%m-%d}  noend={pattern.NoEndDate}"
        )
    print(f"  recurring series total={len(masters)}")
    if latest is not None:
        # Validate the Jet filter on a window that must contain the latest known event.
        lo, hi = latest - timedelta(days=10), latest + timedelta(days=2)
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        fmt = "%d/%m/%Y %H:%M"
        flt = f"[Start] < '{hi.strftime(fmt)}' AND [End] > '{lo.strftime(fmt)}'"
        past = list(iterate(items.Restrict(flt), 500))
        print(f"  past-window check: {len(past)} event(s) for {flt}")
        for ev in past:
            s, s_utc = ev.Start, ev.StartUTC
            print(
                f"      Start={describe_dt(s)}  StartUTC={describe_dt(s_utc)}"
                f"  Start-StartUTC={naive(s) - naive(s_utc)}"
            )
    for store in iterate(_Seq(ns.Stores)):
        _walk_calendars(store.GetRootFolder(), depth=0)


def _walk_calendars(folder, depth: int) -> None:
    if depth > 6:
        return
    if folder.DefaultItemType == 1:  # olAppointmentItem
        masters = sum(1 for ev in iterate(folder.Items) if ev.IsRecurring)
        start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=31)
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True
        fmt = "%d/%m/%Y %H:%M"
        flt = f"[Start] < '{end.strftime(fmt)}' AND [End] > '{start.strftime(fmt)}'"
        upcoming = list(iterate(items.Restrict(flt), 500))
        occurrences = sum(1 for ev in upcoming if ev.IsRecurring)
        print(
            f"  calendar folder: {mask(folder.Name)}  depth={depth}  items={folder.Items.Count}"
            f"  recurring_series={masters}  next31d={len(upcoming)}  occurrences={occurrences}"
        )
        for ev in upcoming[:3]:
            s, s_utc = ev.Start, ev.StartUTC
            print(
                f"      {naive(s):%a %d/%m %H:%M}  allday={ev.AllDayEvent}"
                f"  Start-StartUTC={naive(s) - naive(s_utc)}  subj={mask(ev.Subject)}"
            )
    for sub in iterate(_Seq(folder.Folders)):
        _walk_calendars(sub, depth + 1)


def main() -> None:
    global SHOW
    parser = argparse.ArgumentParser()
    parser.add_argument("--show", action="store_true", help="print real values (local only)")
    SHOW = parser.parse_args().show
    sys.stdout.reconfigure(encoding="utf-8")
    pythoncom.CoInitialize()
    try:
        check_environment()
        _app, ns = connect()
        check_accounts(ns)
        check_inbox(ns)
        fmt = check_restrict_formats(ns)
        check_table(ns)
        check_calendar(ns, fmt)
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
