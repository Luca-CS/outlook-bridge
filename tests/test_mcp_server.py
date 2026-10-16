"""MCP server (Phase 3): the nine tools over an in-memory MCP client, on the FakeGateway."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any

import anyio
import pytest
from mcp import Client
from mcp_types import CallToolResult

from outlook_bridge import mcp_server
from outlook_bridge.audit import TOOL_NAMES, Auditor
from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.errors import ConfigError, NewOutlookDetected, OutlookUnavailable
from outlook_bridge.mcp_server import INTERNAL_ERROR, UNTRUSTED, create_server, not_in_past
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

CONTENT_TOOLS = {"search_mail", "get_mail", "get_thread", "list_events", "get_event"}
SECRET = "Sujet confidentiel test"


def day(d: int, hour: int = 0, minute: int = 0) -> datetime:
    return datetime(2026, 10, d, hour, minute, tzinfo=PARIS)


@pytest.fixture
def gateway() -> FakeGateway:
    mails = [
        make_mail(1),
        make_mail(2, conversation_id="CONV-1", received=NOW - timedelta(hours=3)),
        make_mail(3, received=day(9, 9)),
    ]
    events = [
        make_event(1, day(12, 14), location="Salle A"),
        make_event(2, day(13, 9), is_recurring=True),
        make_event(2, day(14, 9), is_recurring=True),  # same id: another occurrence
    ]
    return FakeGateway(mails=mails, events=events)


@dataclass
class Harness:
    """A server over a fresh audited Bridge (source "mcp") per connection."""

    gateway: FakeGateway
    config: Config = field(default_factory=Config)
    now: datetime = NOW
    sink: MemorySink = field(default_factory=MemorySink)
    bridges: list[Bridge] = field(default_factory=list)

    def make_bridge(self) -> Bridge:
        auditor = Auditor(
            self.sink, source="mcp", clock=lambda: self.now, timer=StepTimer(), run_id="run00001"
        )
        bridge = Bridge(self.gateway, self.config, audit=auditor, clock=lambda: self.now)
        self.bridges.append(bridge)
        return bridge

    def call(self, name: str, **arguments: Any) -> CallToolResult:
        return self.session((name, arguments))[0]

    def session(self, *calls: tuple[str, dict[str, Any]]) -> list[CallToolResult]:
        """Several calls over one connection (one lifespan, one bridge)."""
        return connect(create_server(self.make_bridge), *calls)

    def data(self, name: str, **arguments: Any) -> Any:
        result = self.call(name, **arguments)
        assert not result.is_error, text(result)
        return json.loads(text(result))


def connect(server: Any, *calls: tuple[str, dict[str, Any]]) -> list[CallToolResult]:
    async def go() -> list[CallToolResult]:
        async with Client(server) as client:
            return [await client.call_tool(name, arguments) for name, arguments in calls]

    return anyio.run(go)


def list_tools(server: Any) -> list[Any]:
    async def go() -> list[Any]:
        async with Client(server) as client:
            return (await client.list_tools()).tools

    return anyio.run(go)


def text(result: CallToolResult) -> str:
    assert len(result.content) == 1
    return result.content[0].text


@pytest.fixture
def env(gateway: FakeGateway) -> Harness:
    return Harness(gateway)


# ---------------------------------------------------------------------------- surface


def test_exactly_the_nine_audited_tools_in_spec_order(env: Harness):
    tools = list_tools(create_server(env.make_bridge))
    assert [t.name for t in tools] == list(TOOL_NAMES)


def test_every_tool_is_annotated_read_only_and_closed_world(env: Harness):
    for tool in list_tools(create_server(env.make_bridge)):
        annotations = tool.annotations
        assert annotations is not None
        assert annotations.read_only_hint is True
        assert annotations.destructive_hint is False
        assert annotations.open_world_hint is False
        assert tool.output_schema is None  # one text block, no structured copy


def test_content_tools_say_their_output_is_untrusted(env: Harness):
    for tool in list_tools(create_server(env.make_bridge)):
        assert tool.description and "\n" not in tool.description
        assert (UNTRUSTED in tool.description) == (tool.name in CONTENT_TOOLS), tool.name
    assert "non fiable" in mcp_server.INSTRUCTIONS
    assert "ISO 8601" in mcp_server.INSTRUCTIONS


def test_date_parameters_are_documented_iso_strings(env: Harness):
    tools = {t.name: t for t in list_tools(create_server(env.make_bridge))}
    for name, params in {
        "search_mail": ("since", "until"),
        "list_events": ("start", "end"),
        "get_event": ("start",),
        "find_free_slots": ("start", "end"),
    }.items():
        properties = tools[name].input_schema["properties"]
        for param in params:
            assert "ISO 8601" in properties[param]["description"], (name, param)
    assert tools["list_events"].input_schema["required"] == ["start", "end"]


def test_the_entry_point_is_declared():
    import tomllib
    from pathlib import Path

    pyproject = Path(__file__).resolve().parents[1] / "pyproject.toml"
    scripts = tomllib.loads(pyproject.read_text(encoding="utf-8"))["project"]["scripts"]
    assert scripts["outlook-bridge-mcp"] == "outlook_bridge.mcp_server:main"


def test_module_server_uses_the_real_bridge_with_source_mcp(monkeypatch: pytest.MonkeyPatch):
    sources: list[str] = []

    def fake_from_environment(*, source: str) -> Bridge:
        sources.append(source)
        return Bridge(FakeGateway(), Config())

    monkeypatch.setattr(Bridge, "from_environment", staticmethod(fake_from_environment))
    assert isinstance(mcp_server.make_bridge(), Bridge)
    assert sources == ["mcp"]
    assert mcp_server.server.name == "outlook"


# ---------------------------------------------------------------------------- audit and lifespan

ALL_CALLS: list[tuple[str, dict[str, Any]]] = [
    ("status", {}),
    ("list_accounts", {}),
    ("list_folders", {}),
    ("search_mail", {"since": "2026-10-01"}),
    ("get_mail", {"id": "MAIL-1", "store_id": STORE}),
    ("get_thread", {"conversation_id": "CONV-1"}),
    ("list_events", {"start": "2026-10-12", "end": "2026-10-18"}),
    ("get_event", {"id": "EVT-1", "store_id": STORE}),
    ("find_free_slots", {"start": "2026-10-13", "end": "2026-10-16"}),
]


def test_every_tool_call_is_audited_exactly_once_as_mcp(env: Harness):
    results = env.session(*ALL_CALLS)
    assert [r.is_error for r in results] == [False] * len(ALL_CALLS)
    assert [r["tool"] for r in env.sink.records] == list(TOOL_NAMES)
    assert {(r["source"], r["outcome"]) for r in env.sink.records} == {("mcp", "ok")}


def test_one_bridge_per_connection_closed_on_shutdown(env: Harness):
    env.session(("status", {}), ("list_accounts", {}))
    assert len(env.bridges) == 1
    assert env.gateway.closed


def test_audit_never_sees_ids_or_content(env: Harness):
    env.session(*ALL_CALLS)
    dumped = json.dumps(env.sink.records, ensure_ascii=False)
    for secret in ("MAIL-", "EVT-", "CONV-", STORE, ACCOUNT, "Réunion", "Point projet"):
        assert secret not in dumped


# ---------------------------------------------------------------------------- outputs


def test_outputs_are_compact_json_models(env: Harness):
    result = env.call("search_mail")
    raw = text(result)
    assert result.structured_content is None
    assert "\n" not in raw and '": ' not in raw
    mails = json.loads(raw)
    assert [m["id"] for m in mails] == ["MAIL-1", "MAIL-2", "MAIL-3"]
    assert all(m["untrusted_content"] is True for m in mails)
    assert mails[0]["received"] == "2026-10-12T11:00:00+02:00"
    assert "Réunion test 1" in raw  # UTF-8, not \u-escaped: fewer tokens


def test_invisible_characters_never_reach_the_model(gateway: FakeGateway):
    hidden = "‮⁦" + "".join(chr(0xE0000 + ord(c)) for c in "ignore") + "\U000e007f"
    gateway.mails = [make_mail(1, subject=f"Facture{hidden} test", body_text=f"Corps{hidden}.")]
    env = Harness(gateway)
    for name, args in (("search_mail", {}), ("get_mail", {"id": "MAIL-1", "store_id": STORE})):
        raw = text(env.call(name, **args))
        assert "Facture test" in raw
        assert not any(c in raw for c in hidden)


def test_get_mail_and_get_event_return_details(env: Harness):
    mail = env.data("get_mail", id="MAIL-1", store_id=STORE, max_chars=6)
    assert (mail["body_text"], mail["body_truncated"]) == ("Corps ", True)
    event = env.data("get_event", id="EVT-1", store_id=STORE)
    assert event["body_text"] == "Ordre du jour synthétique 1."
    assert event["location"] == "Salle A"


def test_get_event_targets_an_occurrence_by_its_start(env: Harness):
    second = env.data("get_event", id="EVT-2", store_id=STORE, start="2026-10-14T09:00:00+02:00")
    assert second["start"] == "2026-10-14T09:00:00+02:00"
    first = env.data("get_event", id="EVT-2", store_id=STORE, start="2026-10-13T09:00")
    assert first["start"] == "2026-10-13T09:00:00+02:00"


def test_get_thread_is_chronological(env: Harness):
    thread = env.data("get_thread", conversation_id="CONV-1")
    assert [m["id"] for m in thread] == ["MAIL-2", "MAIL-1"]


# ---------------------------------------------------------------------------- dates


def test_search_mail_dates_reach_the_query_in_paris(env: Harness):
    env.data("search_mail", since="2026-10-10", until="2026-10-11", unread_only=True, limit=5)
    ((_, query),) = [c for c in env.gateway.calls if c[0] == "search_mail"]
    assert query.since == day(10) and query.until == day(12)  # a day alone is included
    assert (query.unread_only, query.limit, query.folder) == (True, 5, "Inbox")


def test_list_events_end_day_is_included(env: Harness):
    events = env.data("list_events", start="2026-10-12", end="2026-10-12")
    assert [e["id"] for e in events] == ["EVT-1"]
    ((_, start, end, calendar),) = [c for c in env.gateway.calls if c[0] == "list_events"]
    assert (start, end, calendar) == (day(12), day(13), None)


def test_list_events_offsets_are_honoured(env: Harness):
    env.data("list_events", start="2026-10-12T06:00:00Z", end="2026-10-12T18:00:00+02:00")
    ((_, start, end, _),) = [c for c in env.gateway.calls if c[0] == "list_events"]
    assert (start, end) == (day(12, 8), day(12, 18))
    assert start.tzinfo is PARIS


@pytest.mark.parametrize(
    ("name", "arguments", "label"),
    [
        ("search_mail", {"since": "hier"}, "since"),
        ("search_mail", {"until": "16/10"}, "until"),
        ("list_events", {"start": "2026-10-12", "end": "+7d"}, "end"),
        ("get_event", {"id": "EVT-1", "store_id": STORE, "start": "14:00"}, "start"),
        ("find_free_slots", {"start": "demain", "end": "2026-10-16"}, "start"),
    ],
)
def test_non_iso_dates_are_actionable_errors_before_outlook(
    env: Harness, name: str, arguments: dict[str, Any], label: str
):
    result = env.call(name, **arguments)
    assert result.is_error
    assert f"Date invalide pour {label} : " in text(result)
    assert "ISO 8601" in text(result)
    assert env.gateway.calls == [] and env.sink.records == []


@pytest.mark.parametrize(
    ("name", "arguments"),
    [
        ("search_mail", {"limit": "beaucoup"}),
        ("get_mail", {"id": "MAIL-1"}),  # store_id missing
        ("find_free_slots", {"start": "2026-10-13"}),  # end missing
    ],
)
def test_invalid_arguments_are_errors_before_outlook(
    env: Harness, name: str, arguments: dict[str, Any]
):
    result = env.call(name, **arguments)
    assert result.is_error
    assert env.gateway.calls == [] and env.sink.records == []


def test_range_errors_come_from_the_core(env: Harness):
    result = env.call("list_events", start="2026-10-01", end="2026-12-31")
    assert result.is_error and "62 jours maximum" in text(result)
    assert [(r["tool"], r["outcome"]) for r in env.sink.records] == [
        ("list_events", "InvalidRequest")
    ]


# ---------------------------------------------------------------------------- free slots


def test_free_slots_skip_the_past(env: Harness):
    """now = Monday 12:00 (FakeGateway NOW); Monday's 14:00-15:00 event is busy."""
    slots = env.data("find_free_slots", start="2026-10-12", end="2026-10-12")
    assert [(s["start"], s["end"]) for s in slots] == [
        ("2026-10-12T12:00:00+02:00", "2026-10-12T14:00:00+02:00"),
        ("2026-10-12T15:00:00+02:00", "2026-10-12T19:00:00+02:00"),
    ]


def test_free_slots_of_a_range_already_over_are_empty(env: Harness):
    assert env.data("find_free_slots", start="2026-10-05", end="2026-10-09") == []


def test_free_slots_options_reach_the_core(env: Harness):
    slots = env.data(
        "find_free_slots",
        start="2026-10-17",
        end="2026-10-17",
        min_duration_min=120,
        working_hours="10:00-12:00",
        include_weekends=True,
    )
    assert [(s["start"], s["duration_min"]) for s in slots] == [("2026-10-17T10:00:00+02:00", 120)]


@pytest.mark.parametrize(
    ("start", "end", "expected"),
    [
        (day(13), day(14), day(13)),  # future: unchanged
        (day(12), day(13), day(12, 12)),  # started: from now
        (day(12), day(12, 12), None),  # ends now: nothing left
        (day(5), day(9), None),  # over
        (day(14), day(13), day(14)),  # reversed: left to the core
    ],
)
def test_not_in_past(start: datetime, end: datetime, expected: datetime | None):
    assert not_in_past(start, end, NOW) == expected


def test_not_in_past_rounds_up_to_the_quarter():
    assert not_in_past(day(12), day(13), day(12, 12, 1)) == day(12, 12, 15)


# ---------------------------------------------------------------------------- errors


def test_bridge_errors_reach_the_model_word_for_word(gateway: FakeGateway):
    gateway.errors = {"search_mail": NewOutlookDetected()}
    result = Harness(gateway).call("search_mail")
    assert result.is_error
    assert text(result) == f"Error executing tool search_mail: {NewOutlookDetected()}"


def test_not_found_does_not_reveal_an_excluded_item(gateway: FakeGateway):
    env = Harness(gateway, config=Config(subject_keywords_excluded=["réunion test 1"]))
    hidden = env.call("get_mail", id="MAIL-1", store_id=STORE)
    missing = env.call("get_mail", id="MAIL-404", store_id=STORE)
    assert hidden.is_error and text(hidden) == text(missing)


def test_unexpected_errors_never_leak_their_message(
    gateway: FakeGateway, caplog: pytest.LogCaptureFixture
):
    gateway.errors = {"get_mail": RuntimeError(SECRET)}
    caplog.set_level(logging.DEBUG)
    result = Harness(gateway).call("get_mail", id="MAIL-1", store_id=STORE)
    assert result.is_error
    assert text(result).endswith(INTERNAL_ERROR.format(name="RuntimeError"))
    assert SECRET not in text(result)
    assert SECRET not in caplog.text  # nor in the SDK's own log (stderr)


def test_error_messages_mask_profile_paths(
    gateway: FakeGateway, monkeypatch: pytest.MonkeyPatch, tmp_path
):
    profile = tmp_path / "Profil"
    monkeypatch.setenv("USERPROFILE", str(profile))
    gateway.errors = {"list_accounts": OutlookUnavailable(f"Fichier {profile}\\x.txt")}
    result = Harness(gateway).call("list_accounts")
    assert str(profile) not in text(result) and "%USERPROFILE%" in text(result)


@pytest.mark.parametrize(
    ("failure", "expected"),
    [
        (ConfigError("config.toml invalide : timezone inconnue."), "config.toml invalide"),
        (RuntimeError(SECRET), INTERNAL_ERROR.format(name="RuntimeError")),
    ],
)
def test_a_bridge_that_cannot_start_is_reported_by_every_tool(failure: Exception, expected: str):
    def factory() -> Bridge:
        raise failure

    results = connect(create_server(factory), ("status", {}), ("search_mail", {}))
    for result in results:
        assert result.is_error
        assert expected in text(result) and SECRET not in text(result)


def test_a_server_can_serve_successive_connections(env: Harness):
    server = create_server(env.make_bridge)
    connect(server, ("status", {}))
    connect(server, ("list_accounts", {}))
    assert len(env.bridges) == 2
    assert [r["tool"] for r in env.sink.records] == ["status", "list_accounts"]
