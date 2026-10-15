"""PYTEST_DONT_REWRITE: CLI against the real classic Outlook (`uv run pytest -m outlook`, run
by the user only).

The real `make_bridge` is used: COM through the STA worker, the audit log and the ordinal
cache. The autouse isolation fixture (tests/conftest.py) points OUTLOOK_BRIDGE_CONFIG at a
missing file (default Config) and OUTLOOK_BRIDGE_DATA_DIR at tmp_path, so the user's
config.toml and local profile are never read or written.

A failure must never print mail content, an ID or an address (rule 5). The marker above
turns off pytest's assertion rewriting for this module, and every assert tests a boolean or
a count computed beforehand, with a literal message: pytest has no operand to show. CliRunner
keeps the command output in its own buffers. Default Config means no exclusion: every test
deletes its tmp_path (snapshots, ordinal cache with real IDs, audit log) on the way out.
"""

from __future__ import annotations

import json
import re
import shutil
import sys
from collections.abc import Iterator
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


@pytest.fixture(autouse=True)
def _wipe_real_outputs(tmp_path: Path) -> Iterator[None]:
    """Nothing written from the real mailbox outlives the test (pytest keeps tmp_path)."""
    yield
    shutil.rmtree(tmp_path, ignore_errors=True)


def run(*args: str) -> Result:
    return CliRunner().invoke(cli.app, list(args), prog_name="outlook-bridge")


def succeeded(*args: str) -> bool:
    return run(*args).exit_code == 0


def parsed_json(result: Result) -> object:
    exit_ok = result.exit_code == 0
    assert exit_ok, "code de sortie non nul"
    ascii_ok = result.stdout.isascii()
    assert ascii_ok, "sortie JSON non ASCII"
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
    ok = succeeded("status")
    assert ok, "status en échec"


def test_mails_json_respects_the_limit():
    items = parsed_json(run("mails", "--json", "--limit", "3"))
    is_list = isinstance(items, list)
    assert is_list, "la sortie JSON n'est pas une liste"
    count = len(items)
    assert count <= 3, "plus de 3 mails"
    flagged = all(item["untrusted_content"] is True for item in items)
    assert flagged, "untrusted_content absent"


def test_events_and_slots_json_parse():
    events = parsed_json(run("events", "--json"))
    is_list = isinstance(events, list)
    assert is_list, "la sortie JSON des événements n'est pas une liste"
    flagged = all(event["untrusted_content"] is True for event in events)
    assert flagged, "untrusted_content absent"
    slots = parsed_json(run("slots", "--json"))
    is_list = isinstance(slots, list)
    assert is_list, "la sortie JSON des créneaux n'est pas une liste"
    long_enough = all(slot["duration_min"] >= 60 for slot in slots)
    assert long_enough, "créneau de moins de 60 min"


def test_export_writes_three_dated_files(tmp_path: Path):
    out = tmp_path / "snap"
    try:
        ok = succeeded("export", "--out", str(out), "--quiet")
        assert ok, "export en échec"
        expected_files = sorted(p.name for p in out.iterdir()) == sorted(SNAPSHOT_FILES)
        assert expected_files, "fichiers de snapshot inattendus"
        for name in SNAPSHOT_FILES:
            text = (out / name).read_text(encoding="utf-8")
            if name == CALENDAR_JSON:
                stamp = json.loads(text).get("generated_at")
            else:
                stamp = front_matter_value(text, "generated_at")
            aware = is_aware(stamp)
            assert aware, "generated_at sans fuseau"
    finally:
        # Real, unexcluded content (default Config): never left behind in %TEMP%.
        shutil.rmtree(out, ignore_errors=True)


def test_audit_log_holds_only_scrubbed_values():
    status_ok = succeeded("status")
    assert status_ok, "status en échec"
    mails_ok = succeeded("mails", "--limit", "3")
    assert mails_ok, "mails en échec"
    events_ok = succeeded("events")
    assert events_ok, "events en échec"
    slots_ok = succeeded("slots")
    assert slots_ok, "slots en échec"
    files = list(log_dir().glob("audit-*.jsonl"))
    lines = [line for f in files for line in f.read_text(encoding="utf-8").splitlines()]
    count = len(lines)
    assert count >= 4, "journal d'audit incomplet"
    records = [json.loads(line) for line in lines]
    from_cli = all(r["source"] == "cli" for r in records)
    assert from_cli, "source inattendue dans le journal d'audit"
    values = [v for r in records for v in r["params"].values() if isinstance(v, str)]
    scrubbed = all(SCRUBBED.match(v) or v.casefold() in WELL_KNOWN_NAMES for v in values)
    assert scrubbed, "valeur non masquée dans le journal d'audit"
