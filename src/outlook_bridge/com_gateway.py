"""OutlookGateway over COM (pywin32). The only module that touches COM.

Every COM call runs on one dedicated STA worker thread; callers only receive pydantic
models. Strictly read-only: the only assignment on a COM object is the collection setting
`Items.IncludeRecurrences`, required to expand recurrences (spec §7.5).
"""

from __future__ import annotations

import queue
import re
import subprocess
import threading
from collections.abc import Callable, Iterator, Sequence
from concurrent.futures import Future
from concurrent.futures import TimeoutError as FutureTimeout
from datetime import datetime, tzinfo
from typing import Any, NamedTuple, TypeVar

import pythoncom
import pywintypes
import win32com.client

from outlook_bridge.errors import (
    NewOutlookDetected,
    NotFound,
    OutlookBridgeError,
    OutlookNotRunning,
    OutlookUnavailable,
)
from outlook_bridge.events import build_jet_range, find_meeting_url
from outlook_bridge.gateway import AcceptEvent, AcceptMail, MailQuery
from outlook_bridge.mail import build_dasl_filter
from outlook_bridge.models import (
    Account,
    Attachment,
    BusyStatus,
    Event,
    Folder,
    Importance,
    MailDetail,
    MailSummary,
    ResponseStatus,
    Status,
)
from outlook_bridge.redact import SNIPPET_MAX
from outlook_bridge.timeutil import DEFAULT_TZ, to_aware, to_com

T = TypeVar("T")

# Outlook enums
OL_FOLDER_DELETED, OL_FOLDER_SENT, OL_FOLDER_INBOX = 3, 5, 6
OL_FOLDER_CALENDAR, OL_FOLDER_DRAFTS, OL_FOLDER_JUNK = 9, 16, 23
OL_APPOINTMENT_ITEM = 1  # Folder.DefaultItemType
OL_CLASS_APPOINTMENT = 26
NON_MAIL_CLASSES = {OL_CLASS_APPOINTMENT, 40, 48, 49}  # appointment, contact, task, task request
CACHED_MODES = {200, 400, 500, 600, 700}

WELL_KNOWN_FOLDERS = {
    "inbox": OL_FOLDER_INBOX,
    "boîte de réception": OL_FOLDER_INBOX,
    "sent": OL_FOLDER_SENT,
    "sent items": OL_FOLDER_SENT,
    "éléments envoyés": OL_FOLDER_SENT,
    "drafts": OL_FOLDER_DRAFTS,
    "brouillons": OL_FOLDER_DRAFTS,
    "deleted items": OL_FOLDER_DELETED,
    "éléments supprimés": OL_FOLDER_DELETED,
    "junk email": OL_FOLDER_JUNK,
    "courrier indésirable": OL_FOLDER_JUNK,
    "calendar": OL_FOLDER_CALENDAR,
    "calendrier": OL_FOLDER_CALENDAR,
}
IMPORTANCE: dict[int, Importance] = {0: "low", 1: "normal", 2: "high"}
BUSY: dict[int, BusyStatus] = {
    0: "free",
    1: "tentative",
    2: "busy",
    3: "oof",
    4: "working_elsewhere",
}
RESPONSE: dict[int, ResponseStatus] = {
    0: "none",
    1: "organizer",
    2: "tentative",
    3: "accepted",
    4: "declined",
    5: "not_responded",
}
ATTACHMENT_TYPES = {1: "file", 4: "link", 5: "item", 6: "ole"}

# Table columns (Phase 0: ConversationID is refused by name, DASL dates are true UTC)
PT = "http://schemas.microsoft.com/mapi/proptag/"
COL_RECEIVED_UTC = "urn:schemas:httpmail:datereceived"
COL_CONVERSATION = PT + "0x30130102"
COL_SENDER_SMTP = PT + "0x5D01001F"
COL_PREVIEW = PT + "0x3FD9001F"
COL_HAS_ATTACH = "urn:schemas:httpmail:hasattachment"
DASL_THREAD_TOPIC = "urn:schemas:httpmail:thread-topic"
# Table rows carry *short-term* EntryIDs in cached mode (valid for the session only).
COL_LONGTERM_ID = PT + "0x66700102"
PR_MESSAGE_DELIVERY_TIME = PT + "0x0E060040"
MAIL_COLUMNS = (
    "EntryID",
    COL_LONGTERM_ID,
    "Subject",
    "SenderName",
    "SenderEmailType",
    "SenderEmailAddress",
    "To",
    "UnRead",
    "Importance",
    "CreationTime",
    "ReceivedTime",
    COL_RECEIVED_UTC,
    COL_CONVERSATION,
    COL_SENDER_SMTP,
    COL_PREVIEW,
    COL_HAS_ATTACH,
)
_COL = {name: i for i, name in enumerate(MAIL_COLUMNS)}

SCAN_LIMIT = 20000  # rows scanned per folder table
EVENTS_LIMIT = 2000  # occurrences read per calendar folder
FOLDER_DEPTH = 5
SHORT_LIST = 10
DEFAULT_TIMEOUT_S = 120.0


# ---------------------------------------------------------------- STA worker


class _StaWorker:
    """One thread, COM-initialised once, consuming a queue of callables."""

    def __init__(self, timeout_s: float) -> None:
        self._timeout_s = timeout_s
        self._queue: queue.Queue[tuple[Callable[[], Any], Future[Any]] | None] = queue.Queue()
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None

    def run(self, fn: Callable[[], T]) -> T:
        with self._lock:
            if self._thread is None or not self._thread.is_alive():
                self._thread = threading.Thread(
                    target=self._loop, name="outlook-com-sta", daemon=True
                )
                self._thread.start()
        future: Future[T] = Future()
        self._queue.put((fn, future))
        try:
            return future.result(timeout=self._timeout_s)
        except FutureTimeout as exc:
            raise OutlookUnavailable(
                "Outlook ne répond pas (délai dépassé). Une fenêtre de sécurité ou une boîte de "
                "dialogue est peut-être ouverte dans Outlook."
            ) from exc

    def stop(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            self._queue.put(None)
            self._thread.join(timeout=5)

    def _loop(self) -> None:
        pythoncom.CoInitialize()
        try:
            while (job := self._queue.get()) is not None:
                fn, future = job
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    future.set_result(fn())
                except BaseException as exc:  # noqa: BLE001 - forwarded to the caller
                    future.set_exception(exc)
        finally:
            pythoncom.CoUninitialize()


# ---------------------------------------------------------------- helpers (worker thread only)


def _running_processes() -> set[str]:
    out = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        encoding="mbcs",
        errors="replace",
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        check=False,
    ).stdout
    return {line.split(",", 1)[0].strip('"').casefold() for line in out.splitlines() if line}


def _com_message(exc: pywintypes.com_error) -> str:
    details = exc.excepinfo[2] if exc.excepinfo and exc.excepinfo[2] else exc.strerror
    return str(details or exc)


def _get(obj: Any, name: str, default: Any = None) -> Any:
    try:
        value = getattr(obj, name)
    except (pywintypes.com_error, AttributeError):
        return default
    return default if value is None else value


def _int(value: Any, default: int) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _smtp_by_store(ns: Any) -> dict[str, str]:
    return {
        str(_get(_get(acc, "DeliveryStore"), "StoreID", "")): str(_get(acc, "SmtpAddress", ""))
        for acc in _each(ns.Accounts)
    }


def _parent_folder(item: Any) -> Any:
    """Parent of an occurrence or exception is its master appointment, not the folder."""
    parent = item.Parent
    while _get(parent, "Class") == OL_CLASS_APPOINTMENT:
        parent = parent.Parent
    return parent


def _each(collection: Any) -> Iterator[Any]:
    """1-based COM collections without recurrence expansion (Count is reliable)."""
    for i in range(1, collection.Count + 1):
        yield collection.Item(i)


def _iter_items(items: Any, limit: int) -> Iterator[Any]:
    """GetFirst/GetNext with a safety cap: never trust Count with recurrences."""
    item = items.GetFirst()
    n = 0
    while item is not None and n < limit:
        yield item
        n += 1
        item = items.GetNext()


def _split_names(value: str | None, limit: int | None = None) -> list[str]:
    names = [part.strip() for part in (value or "").split(";") if part.strip()]
    return names[:limit] if limit else names


def _snippet(text: str | None) -> str:
    return " ".join((text or "").split())[:SNIPPET_MAX]


def _clean_body(text: str | None) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _hex(binary: Any) -> str:
    """Binary MAPI column (conversation id, long-term EntryID) -> uppercase hex string."""
    return bytes(binary).hex().upper() if binary else ""


def _folder_path(folder: Any) -> str:
    """`\\\\Store\\Boîte de réception\\Sub` -> `Boîte de réception/Sub`."""
    parts = [p for p in str(_get(folder, "FolderPath", "")).split("\\") if p]
    return "/".join(parts[1:])


class _Location(NamedTuple):
    """Store and folder facts, read once per folder rather than once per row or item
    (each COM property read is a cross-process call)."""

    store_id: str
    account: str
    path: str


def _location(store: Any, folder: Any) -> _Location:
    return _Location(str(store.StoreID), str(store.DisplayName), _folder_path(folder))


def _table(folder: Any, columns: Sequence[str], flt: str = "") -> Iterator[tuple[Any, ...]]:
    """Rows of `folder` restricted to `columns`, newest first, capped at SCAN_LIMIT."""
    table = folder.GetTable(flt) if flt else folder.GetTable()
    table_columns = table.Columns
    table_columns.RemoveAll()
    for name in columns:
        table_columns.Add(name)
    table.Sort("ReceivedTime", True)
    n = 0
    while not table.EndOfTable and n < SCAN_LIMIT:
        n += 1
        yield tuple(table.GetNextRow().GetValues())


def _default_folder(store: Any, folder_id: int) -> Any:
    try:
        return store.GetDefaultFolder(folder_id)
    except pywintypes.com_error:
        return None


def _resolve_folder(store: Any, path: str) -> Any:
    parts = [p for p in path.replace("\\", "/").split("/") if p]
    if not parts:
        return None
    well_known = WELL_KNOWN_FOLDERS.get(parts[0].casefold())
    if well_known is not None:
        folder, rest = _default_folder(store, well_known), parts[1:]
    else:
        folder, rest = store.GetRootFolder(), parts
    for name in rest:
        if folder is None:
            return None
        folder = next(
            (f for f in _each(folder.Folders) if str(f.Name).casefold() == name.casefold()),
            None,
        )
    return folder


def _find_topic(folder: Any, target: str) -> str | None:
    """Conversation topic of `target`, scanning a slim table (3 columns) newest first."""
    for conversation, topic, _ in _table(
        folder, (COL_CONVERSATION, "ConversationTopic", "ReceivedTime")
    ):
        if _hex(conversation) == target:
            return str(topic or "")
    return None


class ComGateway:
    def __init__(self, tz: tzinfo = DEFAULT_TZ, timeout_s: float = DEFAULT_TIMEOUT_S) -> None:
        self._tz = tz
        self._worker = _StaWorker(timeout_s)
        self._last_access: datetime | None = None
        # Only touched on the worker thread:
        self._app: Any = None
        self._ns: Any = None

    def close(self) -> None:
        self._worker.stop()

    # ------------------------------------------------------------ plumbing

    def _call(self, fn: Callable[[Any], T]) -> T:
        def task() -> T:
            try:
                return fn(self._namespace())
            except pywintypes.com_error as exc:
                self._ns = None  # reconnect on next call (Outlook restarted...)
                raise OutlookUnavailable(f"Erreur COM Outlook : {_com_message(exc)}") from exc

        result = self._worker.run(task)
        self._last_access = datetime.now(self._tz)
        return result

    def _namespace(self) -> Any:
        if self._ns is None:
            processes = _running_processes()
            if "outlook.exe" not in processes:
                if "olk.exe" in processes:
                    raise NewOutlookDetected()
                raise OutlookNotRunning()
            try:
                self._app = win32com.client.Dispatch("Outlook.Application")
                self._ns = self._app.GetNamespace("MAPI")
            except pywintypes.com_error as exc:
                raise OutlookUnavailable(
                    "Impossible de joindre Outlook via COM : Python et Outlook doivent tourner "
                    f"tous deux sans élévation ({_com_message(exc)})."
                ) from exc
        return self._ns

    def _stores(self, ns: Any, account: str | None = None) -> Iterator[Any]:
        if not account:
            yield from _each(ns.Stores)
            return
        smtp_by_store = _smtp_by_store(ns)  # only needed to match an account by address
        wanted = account.casefold()
        for store in _each(ns.Stores):
            names = (str(store.DisplayName), smtp_by_store.get(str(store.StoreID), ""))
            if wanted in (name.casefold() for name in names):
                yield store

    # ------------------------------------------------------------ status / accounts / folders

    def status(self) -> Status:
        try:
            return self._call(self._status)
        except NewOutlookDetected as exc:
            return Status(reachable=False, classic=False, message=str(exc))
        except OutlookBridgeError as exc:
            return Status(reachable=False, message=str(exc), last_access=self._last_access)

    def _status(self, ns: Any) -> Status:
        mode = _get(ns, "ExchangeConnectionMode", 0)
        cached = mode in CACHED_MODES or any(
            bool(_get(s, "IsCachedExchange", False)) for s in _each(ns.Stores)
        )
        return Status(
            reachable=True,
            outlook_version=str(self._app.Version),
            classic=True,
            cached_mode=cached,
            accounts=self._accounts(ns),
            last_access=datetime.now(self._tz),
            message=(
                "Mode Exchange mis en cache : les mails anciens peuvent être absents localement."
                if cached
                else None
            ),
        )

    def list_accounts(self) -> list[Account]:
        return self._call(self._accounts)

    def _accounts(self, ns: Any) -> list[Account]:
        smtp_by_store = _smtp_by_store(ns)
        return [
            Account(
                display_name=str(store.DisplayName),
                smtp_address=smtp_by_store.get(str(store.StoreID), ""),
                store_id=str(store.StoreID),
            )
            for store in _each(ns.Stores)
        ]

    def list_folders(self, account: str | None = None) -> list[Folder]:
        def run(ns: Any) -> list[Folder]:
            folders: list[Folder] = []
            for store in self._stores(ns, account):
                self._walk_folders(store.GetRootFolder(), str(store.DisplayName), folders, 0)
            return folders

        return self._call(run)

    def _walk_folders(self, folder: Any, account: str, out: list[Folder], depth: int) -> None:
        for sub in _each(folder.Folders):
            out.append(
                Folder(
                    name=str(sub.Name),
                    path=_folder_path(sub),
                    item_count=int(_get(sub.Items, "Count", 0)),
                    unread_count=int(_get(sub, "UnReadItemCount", 0)),
                    account=account,
                )
            )
            if depth < FOLDER_DEPTH:
                self._walk_folders(sub, account, out, depth + 1)

    # ------------------------------------------------------------ mail

    def search_mail(self, query: MailQuery, accept: AcceptMail) -> list[MailSummary]:
        def run(ns: Any) -> list[MailSummary]:
            found: list[MailSummary] = []
            resolved = 0
            for store in self._stores(ns, query.account):
                folder = _resolve_folder(store, query.folder)
                if folder is None:
                    continue
                resolved += 1
                location = _location(store, folder)
                rows = _table(folder, MAIL_COLUMNS, build_dasl_filter(query, self._tz))
                taken = 0
                for values in rows:
                    if taken >= query.limit:
                        break
                    mail = self._summary(values, location)
                    if mail is not None and accept(mail):
                        found.append(mail)
                        taken += 1
            if resolved == 0:
                raise NotFound(f"Dossier « {query.folder} »")
            found.sort(key=lambda m: m.received, reverse=True)
            return found[: query.limit]

        return self._call(run)

    def _summary(self, values: tuple[Any, ...], location: _Location) -> MailSummary | None:
        def v(name: str) -> Any:
            return values[_COL[name]]

        received = to_aware(v(COL_RECEIVED_UTC), utc=True, tz=self._tz) or to_aware(
            v("CreationTime"), tz=self._tz
        )
        if received is None:
            return None
        sender_email = v(COL_SENDER_SMTP) or ""
        if not sender_email and v("SenderEmailType") != "EX":
            sender_email = v("SenderEmailAddress") or ""
        return MailSummary(
            id=_hex(v(COL_LONGTERM_ID)) or str(v("EntryID")),
            store_id=location.store_id,
            conversation_id=_hex(v(COL_CONVERSATION)),
            subject=str(v("Subject") or ""),
            sender_name=str(v("SenderName") or ""),
            sender_email=str(sender_email),
            to=_split_names(v("To"), SHORT_LIST),
            received=received,
            unread=bool(v("UnRead")),
            importance=IMPORTANCE.get(_int(v("Importance"), 1), "normal"),
            has_attachments=bool(v(COL_HAS_ATTACH)),
            folder=location.path,
            account=location.account,
            snippet=_snippet(v(COL_PREVIEW)),
        )

    def get_mail(self, entry_id: str, store_id: str, max_chars: int) -> MailDetail | None:
        def run(ns: Any) -> MailDetail | None:
            item = self._item(ns, entry_id, store_id)
            if item is None or _get(item, "Class") in NON_MAIL_CLASSES:
                return None
            return self._detail(item, max_chars)

        return self._call(run)

    def _item(self, ns: Any, entry_id: str, store_id: str) -> Any:
        try:
            return ns.GetItemFromID(entry_id, store_id)
        except pywintypes.com_error:
            return None

    def _sender_email(self, item: Any) -> str:
        if _get(item, "SenderEmailType") == "EX":
            user = None
            sender = _get(item, "Sender")
            if sender is not None:
                try:
                    user = sender.GetExchangeUser()
                except pywintypes.com_error:
                    user = None
            return str(_get(user, "PrimarySmtpAddress", "")) if user is not None else ""
        return str(_get(item, "SenderEmailAddress", ""))

    def _detail(self, item: Any, max_chars: int) -> MailDetail:
        folder = item.Parent
        location = _location(folder.Store, folder)
        try:
            delivered = item.PropertyAccessor.GetProperty(PR_MESSAGE_DELIVERY_TIME)
            received = to_aware(delivered, utc=True, tz=self._tz)
        except pywintypes.com_error:
            received = None
        received = (
            received
            or to_aware(_get(item, "ReceivedTime"), tz=self._tz)
            or to_aware(_get(item, "CreationTime"), tz=self._tz)
        )
        body = _clean_body(_get(item, "Body", ""))
        attachments = [
            Attachment(
                name=str(_get(att, "FileName", "") or _get(att, "DisplayName", "")),
                size=int(_get(att, "Size", 0)),
                type=ATTACHMENT_TYPES.get(int(_get(att, "Type", 0)), "other"),
            )
            for att in _each(item.Attachments)
        ]
        return MailDetail(
            id=str(item.EntryID),
            store_id=location.store_id,
            conversation_id=str(_get(item, "ConversationID", "")),
            subject=str(_get(item, "Subject", "")),
            sender_name=str(_get(item, "SenderName", "")),
            sender_email=self._sender_email(item),
            to=_split_names(_get(item, "To"), SHORT_LIST),
            cc=_split_names(_get(item, "CC")),
            received=received,
            unread=bool(_get(item, "UnRead", False)),
            importance=IMPORTANCE.get(_int(_get(item, "Importance"), 1), "normal"),
            has_attachments=bool(attachments),
            folder=location.path,
            account=location.account,
            snippet=_snippet(body),
            body_text=body[:max_chars],
            body_truncated=len(body) > max_chars,
            attachments=attachments,
        )

    def get_thread(self, conversation_id: str, limit: int, accept: AcceptMail) -> list[MailSummary]:
        """COM cannot filter a Table on the binary conversation id (Phase 1 probe). Step 1 scans
        slim Inbox/Sent tables, newest first, to find the conversation topic; step 2 filters on
        that topic in DASL and keeps the rows whose conversation id matches."""
        target = conversation_id.strip().upper()

        def run(ns: Any) -> list[MailSummary]:
            folders = [
                (folder, _location(store, folder))
                for store in self._stores(ns)
                for folder_id in (OL_FOLDER_INBOX, OL_FOLDER_SENT)
                if (folder := _default_folder(store, folder_id)) is not None
            ]
            topic = next((t for f, _ in folders if (t := _find_topic(f, target)) is not None), None)
            if topic is None:
                return []
            flt = f"@SQL=\"{DASL_THREAD_TOPIC}\" = '{topic.replace(chr(39), chr(39) * 2)}'"
            found: list[MailSummary] = []
            for folder, location in folders:
                for values in _table(folder, MAIL_COLUMNS, flt):
                    if _hex(values[_COL[COL_CONVERSATION]]) != target:
                        continue
                    mail = self._summary(values, location)
                    if mail is not None and accept(mail):
                        found.append(mail)
            found.sort(key=lambda m: m.received)
            return found[-limit:]

        return self._call(run)

    # ------------------------------------------------------------ calendar

    def list_events(
        self, start: datetime, end: datetime, calendar: str | None, accept: AcceptEvent
    ) -> list[Event]:
        flt = build_jet_range(start, end, self._tz)

        def run(ns: Any) -> list[Event]:
            events: list[Event] = []
            for store in self._stores(ns):
                root = _default_folder(store, OL_FOLDER_CALENDAR)
                for folder in self._calendars(root, 0):
                    location = _location(store, folder)
                    if calendar and calendar.casefold() not in (
                        str(folder.Name).casefold(),
                        location.path.casefold(),
                    ):
                        continue
                    items = folder.Items
                    items.Sort("[Start]")
                    items.IncludeRecurrences = True
                    for item in _iter_items(items.Restrict(flt), EVENTS_LIMIT):
                        event = self._event(item, location, max_chars=None)
                        if event is not None and accept(event):
                            events.append(event)
            events.sort(key=lambda e: e.start)
            return events

        return self._call(run)

    def _calendars(self, folder: Any, depth: int) -> Iterator[Any]:
        if folder is None or depth > FOLDER_DEPTH:
            return
        if _get(folder, "DefaultItemType") == OL_APPOINTMENT_ITEM:
            yield folder
        for sub in _each(folder.Folders):
            yield from self._calendars(sub, depth + 1)

    def _event(self, item: Any, location: _Location, max_chars: int | None) -> Event | None:
        if _get(item, "Class") != OL_CLASS_APPOINTMENT:
            return None
        start = to_aware(_get(item, "StartUTC"), utc=True, tz=self._tz)
        end = to_aware(_get(item, "EndUTC"), utc=True, tz=self._tz)
        if start is None or end is None:
            return None
        venue = str(_get(item, "Location", ""))
        body = None
        truncated = False
        if max_chars is not None:
            full = _clean_body(_get(item, "Body", ""))
            body, truncated = full[:max_chars], len(full) > max_chars
        return Event(
            id=str(item.EntryID),
            store_id=location.store_id,
            calendar=location.path,
            account=location.account,
            subject=str(_get(item, "Subject", "")),
            start=start,
            end=end,
            all_day=bool(_get(item, "AllDayEvent", False)),
            location=venue,
            organizer=str(_get(item, "Organizer", "")),
            required_attendees=_split_names(_get(item, "RequiredAttendees")),
            optional_attendees=_split_names(_get(item, "OptionalAttendees")),
            busy_status=BUSY.get(int(_get(item, "BusyStatus", 2)), "busy"),
            is_recurring=bool(_get(item, "IsRecurring", False)),
            response_status=RESPONSE.get(int(_get(item, "ResponseStatus", 0)), "none"),
            online_meeting_url=find_meeting_url(venue, body),
            body_text=body,
            body_truncated=truncated,
        )

    def get_event(
        self,
        entry_id: str,
        store_id: str,
        max_chars: int,
        occurrence_start: datetime | None = None,
    ) -> Event | None:
        def run(ns: Any) -> Event | None:
            item = self._item(ns, entry_id, store_id)
            if item is None or _get(item, "Class") != OL_CLASS_APPOINTMENT:
                return None
            if occurrence_start is not None and _get(item, "IsRecurring", False):
                try:
                    pattern = item.GetRecurrencePattern()
                    item = pattern.GetOccurrence(to_com(occurrence_start, self._tz))
                except pywintypes.com_error:
                    return None
            folder = _parent_folder(item)
            return self._event(item, _location(folder.Store, folder), max_chars)

        return self._call(run)
