"""Audit log (Phase 2 design §7): one scrubbed record per tool call, never any content."""

from __future__ import annotations

import contextlib
import hashlib
import inspect
import json
import os
import re
import subprocess
import sys
import threading
import time
from datetime import UTC, datetime, timedelta

import pytest

from outlook_bridge import audit as audit_module
from outlook_bridge import core
from outlook_bridge.audit import (
    TOOL_NAMES,
    Auditor,
    JsonlSink,
    NullSink,
    fingerprint,
    result_count,
    scrub_params,
)
from outlook_bridge.config import Config, data_dir, log_dir
from outlook_bridge.core import Bridge
from outlook_bridge.errors import (
    InvalidRequest,
    NotFound,
    OutlookBridgeError,
    OutlookNotRunning,
)
from outlook_bridge.events import CalendarService
from outlook_bridge.mail import MailService
from outlook_bridge.models import Folder
from outlook_bridge.redact import Policy
from tests.fakes import (
    ACCOUNT,
    NOW,
    PARIS,
    STORE,
    FakeGateway,
    MemorySink,
    StepTimer,
    make_event,
    make_mail,
)

DAY = datetime(2026, 10, 13, tzinfo=PARIS)  # Tuesday
NEXT_DAY = DAY + timedelta(days=1)
RECORD_KEYS = ["ts", "run", "source", "tool", "params", "count", "ms", "outcome"]
# Spec §5.2, written out independently of audit.TOOL_NAMES.
SPEC_TOOLS = (
    "status",
    "list_accounts",
    "list_folders",
    "search_mail",
    "get_mail",
    "get_thread",
    "list_events",
    "get_event",
    "find_free_slots",
)
GATEWAY_METHODS = (
    "status",
    "list_accounts",
    "list_folders",
    "search_mail",
    "get_mail",
    "get_thread",
    "list_events",
    "get_event",
)


@pytest.fixture
def gateway() -> FakeGateway:
    """Overrides the conftest fixture: some data, so that counts are meaningful."""
    return FakeGateway(
        mails=[make_mail(n) for n in range(1, 4)],
        events=[make_event(1, DAY.replace(hour=10)), make_event(2, DAY.replace(hour=14))],
        folders=[Folder(name="Inbox", path="Inbox", item_count=3, unread_count=1, account=ACCOUNT)],
    )


def make_bridge(gateway: FakeGateway, sink, config: Config | None = None) -> Bridge:
    auditor = Auditor(sink, source="test", clock=lambda: NOW, timer=StepTimer(), run_id="run00001")
    return Bridge(gateway, config or Config(), audit=auditor, clock=lambda: NOW)


# tool -> (call, params expected in the record)
TOOL_CALLS = {
    "status": (lambda b: b.status(), {}),
    "list_accounts": (lambda b: b.list_accounts(), {}),
    "list_folders": (lambda b: b.list_folders(), {}),
    "search_mail": (
        lambda b: b.mail.search_mail(),
        {"folder": "Inbox", "unread_only": False, "limit": 25},
    ),
    "get_mail": (
        lambda b: b.mail.get_mail("MAIL-1", STORE),
        {"id": fingerprint("MAIL-1"), "store_id": fingerprint(STORE)},
    ),
    "get_thread": (
        lambda b: b.mail.get_thread("CONV-1"),
        {"conversation_id": fingerprint("CONV-1"), "limit": 20},
    ),
    "list_events": (
        lambda b: b.calendar.list_events(DAY, NEXT_DAY),
        {"start": "2026-10-13T00:00+02:00", "end": "2026-10-14T00:00+02:00"},
    ),
    "get_event": (
        lambda b: b.calendar.get_event("EVT-1", STORE),
        {"id": fingerprint("EVT-1"), "store_id": fingerprint(STORE)},
    ),
    "find_free_slots": (
        lambda b: b.calendar.find_free_slots(DAY, NEXT_DAY),
        {
            "start": "2026-10-13T00:00+02:00",
            "end": "2026-10-14T00:00+02:00",
            "min_duration_min": 30,
            "include_weekends": False,
            "treat_tentative_as_busy": True,
        },
    ),
}


# ---------------------------------------------------------------- one record per tool


@pytest.mark.parametrize("tool", SPEC_TOOLS)
def test_one_record_per_tool(audited_bridge, sink, tool):
    call, expected_params = TOOL_CALLS[tool]
    result = call(audited_bridge)

    assert len(sink.records) == 1
    record = sink.records[0]
    assert list(record) == RECORD_KEYS
    assert record["tool"] == tool
    assert record["source"] == "test"
    assert record["run"] == "run00001"
    assert record["ts"] == "2026-10-12T12:00:00.000+02:00"
    assert record["ms"] == 250
    assert record["outcome"] == "ok"
    expected_count = len(result) if isinstance(result, list) else 1
    assert expected_count >= 1
    assert record["count"] == expected_count
    assert record["params"] == expected_params
    assert sink.whens == [NOW]


def test_params_bound_with_defaults_without_self_or_none(audited_bridge, sink):
    audited_bridge.mail.search_mail(since=NOW - timedelta(days=2), unread_only=True)
    params = sink.records[0]["params"]
    assert params == {
        "folder": "Inbox",
        "since": "2026-10-10T12:00+02:00",
        "unread_only": True,
        "limit": 25,
    }
    assert list(params) == ["folder", "since", "unread_only", "limit"]


def test_ids_and_accounts_are_hashed(audited_bridge, sink):
    audited_bridge.mail.get_mail("MAIL-1", STORE)
    audited_bridge.list_folders(account=ACCOUNT)
    assert sink.records[0]["params"] == {
        "id": fingerprint("MAIL-1"),
        "store_id": fingerprint(STORE),
    }
    assert sink.records[1]["params"] == {"account": fingerprint(ACCOUNT)}
    dump = json.dumps(sink.records, ensure_ascii=False)
    for raw in ("MAIL-1", STORE, ACCOUNT, "example.com"):
        assert raw not in dump


def test_fingerprint_is_short_and_stable():
    assert re.fullmatch(r"#[0-9a-f]{10}", fingerprint("MAIL-1"))
    assert fingerprint("MAIL-1") == fingerprint("MAIL-1") != fingerprint("MAIL-2")


def test_fingerprint_is_documented_as_a_guessable_pseudonym():
    """Unsalted on purpose (the log stays local): a guessed short name can be confirmed,
    so the docstring must not promise more than a pseudonym."""
    guess = "Inbox/Dossier test"
    assert fingerprint(guess) == "#" + hashlib.sha256(guess.encode("utf-8")).hexdigest()[:10]
    doc = " ".join((fingerprint.__doc__ or "").split()).casefold()
    assert "non-reversible" not in doc
    assert "pseudonym" in doc and "guess" in doc


@pytest.mark.parametrize(
    ("folder", "kept"),
    [
        ("Inbox", True),
        ("INBOX", True),
        ("Boîte de réception", True),
        ("Éléments envoyés", True),
        ("Projets/X", False),
        ("Inbox/Projets", False),
    ],
)
def test_folder_kept_only_when_well_known(audited_bridge, sink, folder, kept):
    audited_bridge.mail.search_mail(folder=folder)
    assert sink.records[0]["params"]["folder"] == (folder if kept else fingerprint(folder))


def test_calendar_kept_only_when_well_known(audited_bridge, sink):
    audited_bridge.calendar.list_events(DAY, NEXT_DAY, calendar="Calendrier")
    audited_bridge.calendar.list_events(DAY, NEXT_DAY, calendar="Agenda projet test")
    calendars = [r["params"]["calendar"] for r in sink.records]
    assert calendars == ["Calendrier", fingerprint("Agenda projet test")]


def test_working_hours_kept_only_when_well_formed(audited_bridge, sink):
    audited_bridge.calendar.find_free_slots(DAY, NEXT_DAY, working_hours="09:00-18:00")
    with pytest.raises(InvalidRequest):
        audited_bridge.calendar.find_free_slots(DAY, NEXT_DAY, working_hours="lundi")
    assert [r["params"]["working_hours"] for r in sink.records] == ["09:00-18:00", "<texte:5>"]
    assert [r["outcome"] for r in sink.records] == ["ok", "InvalidRequest"]


def test_scrub_params_rules():
    scrubbed = scrub_params(
        {
            "self": object(),
            "flag": True,
            "n": 3,
            "absent": None,
            "naive": datetime(2026, 10, 12, 9, 30),
            "utc": datetime(2026, 10, 12, 10, 0, tzinfo=UTC),
            "query": "Réunion test",
            "sender": "bob@example.com",
            "values": [1, 2],
            "ratio": 0.5,
        },
        PARIS,
    )
    assert scrubbed == {
        "flag": True,
        "n": 3,
        "naive": "2026-10-12T09:30+02:00",
        "utc": "2026-10-12T12:00+02:00",
        "query": "<texte:12>",
        "sender": "<texte:15>",
        "values": "<list>",
        "ratio": "<float>",
    }
    assert scrubbed["flag"] is True  # bool tested before int


def test_result_count_is_duck_typed():
    assert result_count([1, 2, 3]) == 3
    assert result_count(()) == 0
    assert result_count(None) == 0
    assert result_count("une chaîne") == 1
    assert result_count(object()) == 1


def test_explicit_export_record(audited_bridge, sink):
    """Contract used by export.py: `bridge.audit.call("export", ...)` with a custom count."""
    files = ["a", "b", "c"]
    assert (
        audited_bridge.audit.call("export", {"custom_out": True}, lambda: files, count=len) is files
    )
    record = sink.records[0]
    assert (record["tool"], record["params"], record["count"]) == (
        "export",
        {"custom_out": True},
        3,
    )


# ---------------------------------------------------------------- canary: no content, ever


class TeeSink:
    def __init__(self, *sinks) -> None:
        self.sinks = sinks

    def write(self, record, when) -> None:
        for sink in self.sinks:
            sink.write(record, when)


def canary_gateway(errors=None) -> FakeGateway:
    mails = [
        make_mail(
            n,
            id=f"MARQUEUR-ID-{n}",
            store_id="MARQUEUR-STORE",
            conversation_id="MARQUEUR-CONV",
            subject=f"MARQUEUR-SUJET {n}",
            sender_name="MARQUEUR-EXP",
            sender_email="marqueur@example.com",
            to=["MARQUEUR-DEST"],
            snippet="MARQUEUR-EXTRAIT",
            body_text="MARQUEUR-CORPS",
        )
        for n in (1, 2)
    ]
    events = [
        make_event(
            1,
            DAY.replace(hour=10),
            id="MARQUEUR-EVT",
            store_id="MARQUEUR-STORE",
            subject="MARQUEUR-SUJET",
            location="MARQUEUR-LIEU",
            organizer="MARQUEUR-ORG",
            required_attendees=["MARQUEUR-PART"],
            body_text="MARQUEUR-CORPS",
        )
    ]
    folders = [
        Folder(
            name="MARQUEUR-DOSSIER",
            path="Inbox/MARQUEUR-DOSSIER",
            item_count=1,
            unread_count=0,
            account=ACCOUNT,
        )
    ]
    return FakeGateway(mails=mails, events=events, folders=folders, errors=errors)


def exercise_every_tool(bridge: Bridge) -> tuple[int, int]:
    calls = [
        lambda: bridge.status(),
        lambda: bridge.list_accounts(),
        lambda: bridge.list_folders(),
        lambda: bridge.list_folders(account="MARQUEUR-COMPTE"),
        lambda: bridge.mail.search_mail(query="MARQUEUR-SUJET", sender="MARQUEUR-EXP"),
        lambda: bridge.mail.search_mail(folder="MARQUEUR-DOSSIER", account="MARQUEUR-COMPTE"),
        lambda: bridge.mail.get_mail("MARQUEUR-ID-1", "MARQUEUR-STORE"),
        lambda: bridge.mail.get_mail("MARQUEUR-ABSENT", "MARQUEUR-STORE"),
        lambda: bridge.mail.get_thread("MARQUEUR-CONV"),
        lambda: bridge.calendar.list_events(DAY, NEXT_DAY),
        lambda: bridge.calendar.list_events(DAY, NEXT_DAY, calendar="MARQUEUR-AGENDA"),
        lambda: bridge.calendar.get_event("MARQUEUR-EVT", "MARQUEUR-STORE"),
        lambda: bridge.calendar.find_free_slots(DAY, NEXT_DAY),
        # InvalidRequest whose message quotes the value: only the class name may be logged.
        lambda: bridge.calendar.find_free_slots(DAY, NEXT_DAY, working_hours="MARQUEUR-HEURES"),
    ]
    failures = 0
    for call in calls:
        try:
            call()
        except (OutlookBridgeError, ValueError):
            failures += 1
    return len(calls), failures


def test_canary_never_reaches_the_log(tmp_path, capsys):
    memory = MemorySink()
    files = JsonlSink(tmp_path / "logs")
    sink = TeeSink(memory, files)
    ok_calls, ok_failures = exercise_every_tool(make_bridge(canary_gateway(), sink))
    failing = canary_gateway(
        errors={name: ValueError("MARQUEUR-ERREUR") for name in GATEWAY_METHODS}
    )
    err_calls, err_failures = exercise_every_tool(make_bridge(failing, sink))

    assert len(memory.records) == ok_calls + err_calls
    assert ok_failures >= 2 and err_failures > ok_failures  # error paths really exercised
    outcomes = {r["outcome"] for r in memory.records}
    assert {"ok", "NotFound", "InvalidRequest", "ValueError"} <= outcomes
    assert {r["tool"] for r in memory.records} == set(SPEC_TOOLS)

    lines = (tmp_path / "logs" / "audit-2026-10.jsonl").read_text(encoding="utf-8").splitlines()
    assert len(lines) == len(memory.records)
    assert "marqueur" not in "\n".join(lines).casefold()
    assert "marqueur" not in json.dumps(memory.records, ensure_ascii=False).casefold()
    captured = capsys.readouterr()
    assert "marqueur" not in (captured.out + captured.err).casefold()


# ---------------------------------------------------------------- errors


def test_error_is_reraised_unchanged_and_logged_by_class(sink):
    boom = NotFound("Mail")
    bridge = make_bridge(FakeGateway(errors={"get_mail": boom}), sink)
    with pytest.raises(NotFound) as info:
        bridge.mail.get_mail("MAIL-1", STORE)
    assert info.value is boom
    record = sink.records[0]
    assert (record["outcome"], record["count"]) == ("NotFound", None)
    assert "introuvable" not in json.dumps(record, ensure_ascii=False)


def test_missing_mail_is_logged_as_not_found(audited_bridge, sink):
    with pytest.raises(NotFound):
        audited_bridge.mail.get_mail("NOPE", STORE)
    record = sink.records[0]
    assert (record["tool"], record["outcome"], record["count"]) == ("get_mail", "NotFound", None)
    assert record["params"]["id"] == fingerprint("NOPE")


def test_invalid_range_is_logged(audited_bridge, sink):
    with pytest.raises(InvalidRequest):
        audited_bridge.calendar.list_events(DAY, DAY + timedelta(days=90))
    record = sink.records[0]
    assert (record["outcome"], record["count"]) == ("InvalidRequest", None)
    assert record["params"]["end"] == "2027-01-11T00:00+01:00"


def test_gateway_failure_is_logged(sink):
    bridge = make_bridge(FakeGateway(errors={"search_mail": OutlookNotRunning()}), sink)
    with pytest.raises(OutlookNotRunning):
        bridge.mail.search_mail()
    assert sink.records[0]["outcome"] == "OutlookNotRunning"


# ---------------------------------------------------------------- silent by default


def test_plain_bridge_writes_nothing(bridge):
    assert isinstance(bridge.audit.sink, NullSink)
    assert bridge.audit.source == "core"
    for call, _ in TOOL_CALLS.values():
        call(bridge)
    assert not data_dir().exists()


def test_standalone_services_are_silent(gateway):
    config = Config()
    for service in (
        MailService(gateway, config, Policy(config)),
        CalendarService(gateway, config, Policy(config)),
    ):
        assert isinstance(service._audit.sink, NullSink)
    MailService(gateway, config, Policy(config)).search_mail()
    assert not data_dir().exists()


# ---------------------------------------------------------------- JsonlSink


def test_jsonl_file_named_after_the_clock_month(tmp_path):
    auditor = Auditor(JsonlSink(tmp_path / "logs"), source="test", clock=lambda: NOW)
    assert auditor.call("status", {}, lambda: "résultat") == "résultat"
    path = tmp_path / "logs" / "audit-2026-10.jsonl"
    raw = path.read_bytes()
    assert raw.endswith(b"\n") and b"\r\n" not in raw
    record = json.loads(raw.decode("utf-8"))
    assert list(record) == RECORD_KEYS
    assert record["source"] == "test"


def test_month_rollover_uses_local_time(tmp_path):
    # 01/11 00:30 Paris is still 31/10 in UTC: the file follows the local month.
    times = iter(
        [datetime(2026, 10, 31, 23, 59, tzinfo=PARIS), datetime(2026, 11, 1, 0, 30, tzinfo=PARIS)]
    )
    auditor = Auditor(JsonlSink(tmp_path), source="test", clock=lambda: next(times))
    auditor.call("status", {}, lambda: None)
    auditor.call("status", {}, lambda: None)
    assert sorted(p.name for p in tmp_path.iterdir()) == [
        "audit-2026-10.jsonl",
        "audit-2026-11.jsonl",
    ]
    november = json.loads((tmp_path / "audit-2026-11.jsonl").read_text(encoding="utf-8"))
    assert november["ts"] == "2026-11-01T00:30:00.000+01:00"


def test_unwritable_log_fails_open_with_one_warning(tmp_path, capsys):
    blocker = tmp_path / "logs"
    blocker.write_text("pas un dossier", encoding="utf-8")
    auditor = Auditor(JsonlSink(blocker), source="test", clock=lambda: NOW)
    assert auditor.call("status", {}, lambda: 41) == 41
    assert auditor.call("status", {}, lambda: 42) == 42
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err.count("journal d'audit indisponible") == 1
    assert str(tmp_path) not in captured.err  # class name only, never the path


class BrokenSink:
    def write(self, record, when) -> None:
        raise RuntimeError("MARQUEUR-SINK")


def test_broken_sink_or_counter_never_breaks_the_call(capsys):
    auditor = Auditor(BrokenSink(), source="test", clock=lambda: NOW)
    assert auditor.call("status", {}, lambda: 1) == 1
    assert auditor.call("status", {}, lambda: 2, count=lambda r: 1 // 0) == 2
    with pytest.raises(KeyError):
        auditor.call("status", {}, lambda: {}["x"])
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err.count("journal d'audit indisponible (RuntimeError)") == 1
    assert "MARQUEUR" not in captured.err


def test_concurrent_writes_stay_whole_lines(tmp_path):
    auditor = Auditor(JsonlSink(tmp_path), source="test", clock=lambda: NOW)

    def worker(i: int) -> None:
        for j in range(50):
            auditor.call("search_mail", {"limit": j, "query": "x" * (i + 1)}, lambda: [])

    threads = [threading.Thread(target=worker, args=(i,)) for i in range(8)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    lines = (tmp_path / "audit-2026-10.jsonl").read_text(encoding="utf-8").splitlines()
    assert len(lines) == 400
    assert all(json.loads(line)["tool"] == "search_mail" for line in lines)


# One writer process: signals it is ready, waits for the start file, then appends.
_WRITER = """
import sys, time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
from outlook_bridge.audit import JsonlSink

directory, signals = Path(sys.argv[1]), Path(sys.argv[4])
worker, count = int(sys.argv[2]), int(sys.argv[3])
sink = JsonlSink(directory)
when = datetime(2026, 10, 12, 12, 0, tzinfo=ZoneInfo("Europe/Paris"))
(signals / f"ready-{worker}").touch()
deadline = time.monotonic() + 60
while not (signals / "go").exists() and time.monotonic() < deadline:
    time.sleep(0.001)
for n in range(count):
    sink.write({"worker": worker, "n": n, "pad": "x" * 300}, when)
"""


def test_concurrent_processes_never_lose_a_record(tmp_path):
    """MCP server, scheduled export and manual CLI may append to the same file at once.

    On Windows a CRT append is seek-to-end then write: without a cross-process lock, two
    writers land on the same offset and one record silently replaces the other.
    """
    workers, count = 6, 300
    signals = tmp_path / "signals"
    signals.mkdir()
    logs = tmp_path / "logs"
    procs = [
        subprocess.Popen(
            [sys.executable, "-c", _WRITER, str(logs), str(w), str(count), str(signals)]
        )
        for w in range(workers)
    ]
    try:
        deadline = time.monotonic() + 60
        while len(list(signals.glob("ready-*"))) < workers and time.monotonic() < deadline:
            time.sleep(0.01)
        (signals / "go").touch()
        assert [proc.wait(timeout=120) for proc in procs] == [0] * workers
    finally:
        for proc in procs:
            if proc.poll() is None:
                proc.kill()
    lines = (logs / "audit-2026-10.jsonl").read_text(encoding="utf-8").splitlines()
    records = [json.loads(line) for line in lines]  # every line whole
    assert len(records) == workers * count
    assert {(r["worker"], r["n"]) for r in records} == {
        (w, n) for w in range(workers) for n in range(count)
    }


@pytest.mark.skipif(sys.platform != "win32", reason="msvcrt byte-range locks are Windows-only")
def test_lock_held_elsewhere_fails_open(tmp_path, monkeypatch, capsys):
    """A writer stuck holding the lock delays a record a little, never drops it."""
    import msvcrt

    monkeypatch.setattr(audit_module, "_LOCK_WAIT_S", 0.05)
    sink = JsonlSink(tmp_path)
    path = sink.path_for(NOW)
    with path.open("ab", buffering=0) as other:
        os.lseek(other.fileno(), audit_module._LOCK_OFFSET, os.SEEK_SET)
        msvcrt.locking(other.fileno(), msvcrt.LK_NBLCK, 1)
        sink.write({"n": 1}, NOW)  # waits, then writes unlocked
        sink.write({"n": 2}, NOW)  # locking now off: no further wait
        os.lseek(other.fileno(), audit_module._LOCK_OFFSET, os.SEEK_SET)
        msvcrt.locking(other.fileno(), msvcrt.LK_UNLCK, 1)
    assert path.read_bytes() == b'{"n":1}\n{"n":2}\n'
    assert sink._locking is False
    assert capsys.readouterr().err == ""


# ---------------------------------------------------------------- wiring and lifecycle


class _RolloverDatetime(datetime):
    """`datetime.now` frozen at 01/11/2026 00:30 Paris, which is still 31/10 in UTC."""

    INSTANT = datetime(2026, 10, 31, 23, 30, tzinfo=UTC)

    @classmethod
    def now(cls, tz=None):
        return cls.INSTANT.astimezone(tz) if tz is not None else cls.INSTANT.replace(tzinfo=None)


@pytest.mark.skipif(sys.platform != "win32", reason="pywin32 is Windows-only")
def test_from_environment_wires_the_audit_log(tmp_path, monkeypatch):
    monkeypatch.setattr(core, "datetime", _RolloverDatetime)
    bridge = Bridge.from_environment(source="cli")
    try:
        assert bridge.audit.source == "cli"
        assert isinstance(bridge.audit.sink, JsonlSink)
        assert bridge.audit.sink.directory == tmp_path / "data" / "logs"
        assert re.fullmatch(r"[0-9a-f]{8}", bridge.audit.run_id)
        assert bridge.config.timezone == "Europe/Paris"  # isolated: absent config.toml
        assert bridge.now().tzinfo == PARIS
        assert bridge.gateway._worker._thread is None  # lazy: no STA thread, no COM yet
        assert not (tmp_path / "data").exists()  # log directory created lazily

        # The audit reads the bridge's own config.tz clock: `ts` and the monthly file are
        # local, so 00:30 on 1 November goes to November, not to UTC's October.
        assert bridge.audit._clock is bridge.clock
        bridge.audit.call("status", {}, lambda: None)  # a dummy call: no COM involved
        assert [p.name for p in log_dir().iterdir()] == ["audit-2026-11.jsonl"]
        record = json.loads((log_dir() / "audit-2026-11.jsonl").read_text(encoding="utf-8"))
        assert record["ts"] == "2026-11-01T00:30:00.000+01:00"
        assert (record["source"], record["tool"], record["outcome"]) == ("cli", "status", "ok")
        assert bridge.gateway._worker._thread is None  # still no COM
    finally:
        bridge.close()
    bridge.close()  # idempotent


def test_from_environment_requires_a_source():
    with pytest.raises(TypeError):
        Bridge.from_environment()  # type: ignore[call-arg]
    with pytest.raises(ValueError):
        Bridge.from_environment(source="web")  # type: ignore[arg-type]


def test_each_bridge_has_its_own_run_id(gateway, config):
    first, second = Bridge(gateway, config), Bridge(gateway, config)
    assert re.fullmatch(r"[0-9a-f]{8}", first.audit.run_id)
    assert first.audit.run_id != second.audit.run_id


def test_close_and_context_manager(config):
    gateway = FakeGateway()
    Bridge(gateway, config).close()
    assert gateway.closed

    gateway = FakeGateway()
    with contextlib.suppress(RuntimeError), Bridge(gateway, config) as bridge:
        assert not gateway.closed
        assert bridge.gateway is gateway
        raise RuntimeError("sortie anticipée")
    assert gateway.closed


def test_now_uses_the_clock(audited_bridge, bridge):
    assert audited_bridge.now() == NOW
    assert bridge.now().tzinfo == PARIS


# ---------------------------------------------------------------- coverage guard


def test_every_tool_method_is_audited():
    methods = (
        Bridge.status,
        Bridge.list_accounts,
        Bridge.list_folders,
        MailService.search_mail,
        MailService.get_mail,
        MailService.get_thread,
        CalendarService.list_events,
        CalendarService.get_event,
        CalendarService.find_free_slots,
    )
    assert all(hasattr(m, "__wrapped__") for m in methods)
    assert tuple(m.__name__ for m in methods) == TOOL_NAMES == SPEC_TOOLS
    for method in methods:
        # `__wrapped__` points at the decorated method itself, not at some other function.
        assert method.__wrapped__.__qualname__ == method.__qualname__
        # The wrapper carries the real signature itself (front ends such as FastMCP or typer
        # introspect it). follow_wrapped=False: otherwise both sides follow `__wrapped__`
        # and the check could never fail.
        own = inspect.signature(method, follow_wrapped=False)
        assert own == inspect.signature(method.__wrapped__)
        assert next(iter(own.parameters)) == "self"


def test_no_tool_calls_another_audited_tool(audited_bridge, sink):
    for call, _ in TOOL_CALLS.values():
        call(audited_bridge)
    assert [r["tool"] for r in sink.records] == list(TOOL_CALLS)
