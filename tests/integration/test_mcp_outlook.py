"""PYTEST_DONT_REWRITE: Phase 3 against the real classic Outlook (`uv run pytest -m outlook`,
run by the user only): the MCP server over real stdio, shortened links, Teams detection.

The server runs as the real console script, in a subprocess that inherits the isolated
environment (tests/conftest.py): default Config, audit log in tmp_path. A failure must never
print mail content, an ID or an address (rule 5): assertion rewriting is off for this module
and every assert tests a boolean or a count with a literal message.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import sys
from collections.abc import Iterator
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

import anyio
import pytest
from mcp import Client, StdioServerParameters

from outlook_bridge.audit import TOOL_NAMES
from outlook_bridge.config import Config, log_dir
from outlook_bridge.core import Bridge
from outlook_bridge.redact import LONG_URL
from outlook_bridge.timeutil import DEFAULT_TZ as PARIS

if sys.platform == "win32":
    from outlook_bridge.com_gateway import ComGateway

pytestmark = [
    pytest.mark.outlook,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows + Outlook only"),
]

LINK = re.compile(r"https?://[^\s<>\"']+")
LINK_TAIL = ".,;:!?)]}'\""


@pytest.fixture(autouse=True)
def _wipe_real_outputs(tmp_path: Path) -> Iterator[None]:
    yield
    shutil.rmtree(tmp_path, ignore_errors=True)


@pytest.fixture(scope="module")
def bridge() -> Iterator[Bridge]:
    gateway = ComGateway()
    yield Bridge(gateway, Config())
    gateway.close()


def no_long_link(text: str | None) -> bool:
    return all(len(m.rstrip(LINK_TAIL)) <= LONG_URL for m in LINK.findall(text or ""))


def stdio_session(*calls: tuple[str, dict[str, Any]]) -> list[Any]:
    script = Path(sys.executable).with_name("outlook-bridge-mcp.exe")
    script_found = script.exists()
    assert script_found, "script outlook-bridge-mcp introuvable (uv sync ?)"
    params = StdioServerParameters(command=str(script), args=[], env=dict(os.environ))

    async def go() -> list[Any]:
        async with Client(params, read_timeout_seconds=120) as client:
            return [await client.call_tool(name, arguments) for name, arguments in calls]

    return anyio.run(go)


def test_stdio_server_answers_every_kind_of_call():
    today = datetime.now(PARIS).date()
    results = stdio_session(
        ("status", {}),
        ("search_mail", {"limit": 3}),
        ("list_events", {"start": today.isoformat(), "end": (today + timedelta(13)).isoformat()}),
        (
            "find_free_slots",
            {"start": today.isoformat(), "end": (today + timedelta(6)).isoformat()},
        ),
    )
    no_error = not any(r.is_error for r in results)
    assert no_error, "un outil a renvoyé une erreur"
    status, mails, events, slots = (json.loads(r.content[0].text) for r in results)
    reachable = status["reachable"] is True
    assert reachable, "Outlook injoignable"
    within_limit = isinstance(mails, list) and len(mails) <= 3
    assert within_limit, "search_mail dépasse limit"
    untrusted = all(m["untrusted_content"] is True for m in mails + events)
    assert untrusted, "untrusted_content manquant"
    now = datetime.now(PARIS)
    future = all(datetime.fromisoformat(s["start"]) >= now - timedelta(minutes=1) for s in slots)
    assert future, "créneau libre dans le passé"


def test_stdio_server_audits_as_mcp():
    stdio_session(("status", {}), ("list_accounts", {}))
    files = sorted(log_dir().glob("*.jsonl"))
    records = [json.loads(line) for f in files for line in f.read_text("utf-8").splitlines()]
    tools_ok = [r["tool"] for r in records] == ["status", "list_accounts"]
    assert tools_ok, "enregistrements d'audit inattendus"
    source_ok = {r["source"] for r in records} == {"mcp"} and set(TOOL_NAMES) >= {
        r["tool"] for r in records
    }
    assert source_ok, "source d'audit inattendue"


def test_bodies_and_snippets_have_no_long_link(bridge: Bridge):
    mails = bridge.mail.search_mail(limit=10)
    if not mails:
        pytest.skip("empty inbox")
    snippets_ok = all(no_long_link(m.snippet) for m in mails)
    assert snippets_ok, "lien long dans un extrait"
    bodies_ok = all(no_long_link(bridge.mail.get_mail(m.id, m.store_id).body_text) for m in mails)
    assert bodies_ok, "lien long dans un corps"


def test_teams_meetings_are_flagged_and_get_event_finds_their_link(bridge: Bridge):
    start = datetime.now(PARIS).replace(hour=0, minute=0, second=0, microsecond=0)
    events = bridge.calendar.list_events(start - timedelta(days=31), start + timedelta(days=31))
    consistent = all(e.has_online_meeting for e in events if e.online_meeting_url)
    assert consistent, "lien de visio sans has_online_meeting"
    flagged = [e for e in events if e.has_online_meeting][:3]
    if not flagged:
        pytest.skip("no online meeting in the window")
    details = [bridge.calendar.get_event(e.id, e.store_id, start=e.start) for e in flagged]
    with_link = sum(1 for d in details if d.online_meeting_url)
    assert with_link >= 1, "aucun lien de visio trouvé par get_event"
    bodies_ok = all(no_long_link(d.body_text) for d in details)
    assert bodies_ok, "lien long dans le corps d'un événement"
