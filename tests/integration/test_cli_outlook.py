"""CLI against the real classic Outlook (`uv run pytest -m outlook`, run by the user only).

The real `make_bridge` is used: COM through the STA worker, the audit log and the ordinal
cache. The autouse isolation fixture (tests/conftest.py) points OUTLOOK_BRIDGE_CONFIG at a
missing file (default Config) and OUTLOOK_BRIDGE_DATA_DIR at tmp_path, so the user's
config.toml and local profile are never read or written.

Assertions only compare booleans, counts and dates: a failure never prints mail content,
an ID or an address. CliRunner keeps the command output in its own buffers.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pytest
from typer.testing import CliRunner, Result

from outlook_bridge import cli
from outlook_bridge.audit import WELL_KNOWN_NAMES
from outlook_bridge.config import log_dir
from outlook_bridge.export import CALENDAR_JSON, SNAPSHOT_FILES

pytestmark = [
    pytest.mark.outlook,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows + Outlook only"),
]

SCRUBBED = re.compile(
    r"^(#[0-9a-f]{10}|<texte:\d+>|\d{4}-\d{2}-\d{2}T.*|\d{2}:\d{2}-\d{2}:\d{2}|<[A-Za-z]+>)$"
)


def run(*args: str) -> Result:
    return CliRunner().invoke(cli.app, list(args), prog_name="outlook-bridge")


def parsed_json(result: Result) -> object:
    assert result.exit_code == 0
    assert result.stdout.isascii()
    return json.loads(result.stdout)


def front_matter_value(text: str, key: str) -> str | None:
    lines = text.split("\n")
    end = lines.index("---", 1)
    for line in lines[1:end]:
        if line.startswith(f"{key}: "):
            return line.removeprefix(f"{key}: ")
    return None


def is_aware(text: str | None) -> bool:
    return text is not None and datetime.fromisoformat(text).utcoffset() is not None


def test_status_exits_0():
    assert run("status").exit_code == 0


def test_mails_json_respects_the_limit():
    items = parsed_json(run("mails", "--json", "--limit", "3"))
    assert isinstance(items, list)
    assert len(items) <= 3
    assert all(item["untrusted_content"] is True for item in items)


def test_events_and_slots_json_parse():
    events = parsed_json(run("events", "--json"))
    assert isinstance(events, list)
    assert all(event["untrusted_content"] is True for event in events)
    slots = parsed_json(run("slots", "--json"))
    assert isinstance(slots, list)
    assert all(slot["duration_min"] >= 60 for slot in slots)


def test_export_writes_three_dated_files(tmp_path: Path):
    out = tmp_path / "snap"
    assert run("export", "--out", str(out), "--quiet").exit_code == 0
    assert sorted(p.name for p in out.iterdir()) == sorted(SNAPSHOT_FILES)
    for name in SNAPSHOT_FILES:
        text = (out / name).read_text(encoding="utf-8")
        if name == CALENDAR_JSON:
            stamp = json.loads(text).get("generated_at")
        else:
            stamp = front_matter_value(text, "generated_at")
        assert is_aware(stamp)


def test_audit_log_holds_only_scrubbed_values():
    assert run("status").exit_code == 0
    assert run("mails", "--limit", "3").exit_code == 0
    assert run("events").exit_code == 0
    assert run("slots").exit_code == 0
    files = list(log_dir().glob("audit-*.jsonl"))
    lines = [line for f in files for line in f.read_text(encoding="utf-8").splitlines()]
    assert len(lines) >= 4
    records = [json.loads(line) for line in lines]
    assert all(r["source"] == "cli" for r in records)
    values = [v for r in records for v in r["params"].values() if isinstance(v, str)]
    assert all(SCRUBBED.match(v) or v.casefold() in WELL_KNOWN_NAMES for v in values)
