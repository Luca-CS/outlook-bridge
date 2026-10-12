"""Read-only guard (spec §6.1): scans src/ for any COM write, network access or stray COM use."""

from __future__ import annotations

import re
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parents[1] / "src"

FORBIDDEN_CALLS = [
    ".Send(",
    ".Delete(",
    ".Move(",
    ".Copy(",
    ".Save(",
    ".SaveAs(",
    ".Reply(",
    ".ReplyAll(",
    ".Forward(",
    ".Close(",
    "CreateItem(",
    "Items.Add(",
    ".Respond(",
    ".Display(",
    # Other writes reachable from the object model
    ".SetProperty(",
    ".SetProperties(",
    ".DeleteProperties(",
    ".MarkAsTask(",
    ".ClearTaskFlag(",
    ".ClearRecurrencePattern(",
    ".PrintOut(",
    ".Execute(",
    ".AddStore(",
    ".RemoveStore(",
    "setattr(",
]
# Attribute assignment on a PascalCase name = COM property write (`item.UnRead = False`).
PROPERTY_WRITE = re.compile(r"(?<=[\w)\]])\.([A-Z]\w*)[ \t]*(?:[-+*/|&]?=)(?!=)")
ALLOWED_PROPERTY_WRITES = {"IncludeRecurrences"}  # Items collection setting (spec §7.5)
NETWORK_IMPORT = re.compile(
    r"^\s*(?:from|import)\s+(socket|ssl|urllib|http|requests|httpx|aiohttp|ftplib|smtplib|"
    r"imaplib|poplib|websocket|websockets)\b",
    re.MULTILINE,
)
COM_IMPORT = re.compile(r"^\s*(?:from|import)\s+(win32com|pythoncom|pywintypes|win32\w*)\b", re.M)
COM_MODULE = "com_gateway.py"


def violations(source: str) -> list[str]:
    found = [f"forbidden call {token}" for token in FORBIDDEN_CALLS if token in source]
    for match in PROPERTY_WRITE.finditer(source):
        if match.group(1) not in ALLOWED_PROPERTY_WRITES:
            found.append(f"COM property write .{match.group(1)}")
    found += [f"network import {m.group(1)}" for m in NETWORK_IMPORT.finditer(source)]
    return found


def source_files() -> list[Path]:
    files = sorted(SRC.rglob("*.py"))
    assert files, "no source file found"
    return files


@pytest.mark.parametrize("path", source_files(), ids=lambda p: p.name)
def test_source_is_read_only(path: Path):
    assert violations(path.read_text(encoding="utf-8")) == []


@pytest.mark.parametrize("path", source_files(), ids=lambda p: p.name)
def test_only_com_gateway_touches_com(path: Path):
    if path.name != COM_MODULE:
        assert not COM_IMPORT.search(path.read_text(encoding="utf-8"))


@pytest.mark.parametrize(
    "snippet",
    [
        "mail.Send()",
        "item.UnRead = False",
        "item.Categories += 'x'",
        "ns.GetDefaultFolder(6).Items.Add('IPM.Note')",
        "app.CreateItem(0)",
        "item.PropertyAccessor.SetProperty(tag, 1)",
        "setattr(item, 'Subject', 'x')",
        "import socket",
        "from urllib.request import urlopen",
    ],
)
def test_guard_catches_writes(snippet: str):
    assert violations(snippet)


@pytest.mark.parametrize(
    "snippet",
    [
        "items.IncludeRecurrences = True",
        "self._ns = None",
        "if item.UnRead == True:",
        "value = item.Subject",
        "x: dict[str, int] = {}",
        "# a comment ending with a dot.\nCONSTANT = 1",
    ],
)
def test_guard_allows_reads(snippet: str):
    assert violations(snippet) == []
