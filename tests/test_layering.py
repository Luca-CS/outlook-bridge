"""Layering rules (Phase 2 design §3, §7): front ends never reach the gateway or COM, only the
CLI writes to stdout, and the audit module cannot see any model or logging handler."""

from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

PACKAGE = Path(__file__).resolve().parents[1] / "src" / "outlook_bridge"

FRONT_ENDS = ("cli.py", "export.py", "mcp_server.py")  # missing ones are skipped
FRONT_END_FORBIDDEN = ("com_gateway", ".gateway.", "outlook_bridge.gateway", "win32")
STDOUT_OWNER = "cli.py"
PRINT_CALL = re.compile(r"\bprint\s*\(")
AUDIT_FORBIDDEN_IMPORTS = {
    "outlook_bridge.models",
    "outlook_bridge.mail",
    "outlook_bridge.events",
    "logging",
}


def _modules() -> list[Path]:
    files = sorted(PACKAGE.glob("*.py"))
    assert files, "no source file found"
    return files


@pytest.mark.parametrize("name", FRONT_ENDS)
def test_front_ends_never_reach_the_gateway(name: str):
    path = PACKAGE / name
    if not path.exists():
        pytest.skip(f"{name} not written yet")
    source = path.read_text(encoding="utf-8")
    assert [token for token in FRONT_END_FORBIDDEN if token in source] == []


@pytest.mark.parametrize("path", _modules(), ids=lambda p: p.name)
def test_only_the_cli_writes_to_stdout(path: Path):
    """stdout is the MCP stdio channel: a stray print would corrupt the protocol."""
    if path.name == STDOUT_OWNER:
        return
    source = path.read_text(encoding="utf-8")
    assert not PRINT_CALL.search(source)
    assert "sys.stdout" not in source


def _imported_modules(source: str, package: str = "outlook_bridge") -> set[str]:
    found: set[str] = set()
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.Import):
            found.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            base = node.module or ""
            if node.level:  # relative import inside the package
                base = f"{package}.{base}" if base else package
            found.add(base)
            # `from pkg import mod` imports the submodule pkg.mod
            found.update(f"{base}.{alias.name}" for alias in node.names)
    return found


def test_audit_imports_no_model_and_no_logging():
    imported = _imported_modules((PACKAGE / "audit.py").read_text(encoding="utf-8"))
    leaks = {
        name
        for name in imported
        for forbidden in AUDIT_FORBIDDEN_IMPORTS
        if name == forbidden or name.startswith(forbidden + ".")
    }
    assert leaks == set()
    assert "outlook_bridge.timeutil" in imported  # sanity: the walker sees our imports


@pytest.mark.parametrize(
    "source",
    [
        "import logging",
        "from logging import getLogger",
        "from outlook_bridge.models import Event",
        "from outlook_bridge import mail",
        "from . import events",
        "from .models import Event",
    ],
)
def test_import_walker_catches_forbidden_imports(source: str):
    imported = _imported_modules(source)
    assert any(
        name == forbidden or name.startswith(forbidden + ".")
        for name in imported
        for forbidden in AUDIT_FORBIDDEN_IMPORTS
    )
