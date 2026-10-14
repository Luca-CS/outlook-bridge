"""Audit log of tool calls (spec rule 4): tool, scrubbed parameters, result count, duration.

Never a subject, body, address or raw ID: the auditor only receives `count(result)` (an int),
every parameter goes through the deny-by-default `scrub_params`, errors are logged by class
name only and IDs are hashed. This module imports no models and is not a `logging` handler,
so third-party loggers can never reach the audit file. It never writes to stdout (the MCP
stdio channel); failures give a single stderr warning and the call proceeds (fail-open).

The `@audited` decorator sits on the core methods that map 1:1 to the MCP tools, so every
front end (CLI, export, MCP server) is audited without any code of its own.
"""

from __future__ import annotations

import contextlib
import functools
import hashlib
import inspect
import json
import re
import secrets
import sys
import threading
import time
from collections.abc import Callable, Mapping
from datetime import UTC, datetime, tzinfo
from pathlib import Path
from typing import Protocol

from outlook_bridge.timeutil import ensure_aware

Clock = Callable[[], datetime]  # returns an aware datetime in config.tz

TOOL_NAMES: tuple[str, ...] = (
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
HASHED_PARAMS = frozenset({"id", "store_id", "conversation_id", "account"})
NAME_PARAMS = frozenset({"folder", "calendar"})
HOURS_PARAMS = frozenset({"working_hours"})
WELL_KNOWN_NAMES = frozenset(  # casefolded
    {
        "inbox",
        "sent items",
        "drafts",
        "calendar",
        "boîte de réception",
        "éléments envoyés",
        "brouillons",
        "calendrier",
    }
)
_HOURS = re.compile(r"^\d{2}:\d{2}-\d{2}:\d{2}$")


def fingerprint(value: str) -> str:
    """Short, stable, non-reversible tag for an ID or a personal name: `#` + 10 hex chars."""
    return "#" + hashlib.sha256(value.encode("utf-8")).hexdigest()[:10]


def _scrub(key: str, value: object, tz: tzinfo) -> bool | int | str:
    if isinstance(value, bool):  # before int: bool is an int subclass
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, datetime):
        return ensure_aware(value, tz).isoformat(timespec="minutes")
    if isinstance(value, str):
        if key in HASHED_PARAMS:
            return fingerprint(value)
        if key in NAME_PARAMS:
            return value if value.casefold() in WELL_KNOWN_NAMES else fingerprint(value)
        if key in HOURS_PARAMS and _HOURS.match(value):
            return value
        return f"<texte:{len(value)}>"
    return f"<{type(value).__name__}>"


def scrub_params(arguments: Mapping[str, object], tz: tzinfo) -> dict[str, bool | int | str]:
    """Deny by default: only booleans, integers, ISO dates, hashes, well-known folder names
    and `HH:MM-HH:MM` survive; any other string is reduced to its length."""
    return {
        str(key): _scrub(str(key), value, tz)
        for key, value in arguments.items()
        if key != "self" and value is not None
    }


def result_count(result: object) -> int:
    """Duck-typed: a list gives its length, None gives 0, anything else counts as 1."""
    if isinstance(result, list | tuple):
        return len(result)
    return 0 if result is None else 1


def _warn(exc: BaseException) -> None:
    """One line on stderr, class name only (an OSError message may embed a user path)."""
    stream = sys.stderr
    if stream is None:  # pythonw: no console at all
        return
    with contextlib.suppress(Exception):
        stream.write(
            f"Avertissement : journal d'audit indisponible ({type(exc).__name__}) : "
            "les appels continuent sans journal.\n"
        )
        stream.flush()


class AuditSink(Protocol):
    def write(self, record: dict[str, object], when: datetime) -> None: ...


class NullSink:
    """Discards everything: the default for a hand-built Bridge (tests)."""

    def write(self, record: dict[str, object], when: datetime) -> None:
        return None


class JsonlSink:
    """Monthly JSON Lines files `audit-YYYY-MM.jsonl`, one write per record, thread-safe."""

    def __init__(self, directory: Path) -> None:
        self.directory = directory
        self._lock = threading.Lock()
        self._warned = False

    def path_for(self, when: datetime) -> Path:
        return self.directory / f"audit-{when:%Y-%m}.jsonl"

    def write(self, record: dict[str, object], when: datetime) -> None:
        line = json.dumps(record, ensure_ascii=False, separators=(",", ":")) + "\n"
        with self._lock:
            try:
                self.directory.mkdir(parents=True, exist_ok=True)
                with self.path_for(when).open("a", encoding="utf-8", newline="\n") as fh:
                    fh.write(line)
            except OSError as exc:
                if not self._warned:
                    self._warned = True
                    _warn(exc)


class Auditor:
    """Times a call and writes one scrubbed record for it, whatever the outcome."""

    def __init__(
        self,
        sink: AuditSink,
        *,
        source: str,
        clock: Clock,
        timer: Callable[[], float] = time.perf_counter,
        run_id: str | None = None,
    ) -> None:
        self.sink = sink
        self.source = source
        self.run_id = run_id or secrets.token_hex(4)
        self._clock = clock
        self._timer = timer
        self._warned = False

    @classmethod
    def disabled(cls) -> Auditor:
        return cls(NullSink(), source="core", clock=lambda: datetime.now(UTC))

    def call[T](
        self,
        tool: str,
        params: Mapping[str, object],
        fn: Callable[[], T],
        *,
        count: Callable[[T], int] = result_count,
    ) -> T:
        """Runs `fn`, logs it, returns its result or re-raises its exception unchanged."""
        started = self._timer()
        try:
            result = fn()
        except BaseException as exc:
            self._log(tool, params, started, type(exc).__name__, None)
            raise
        self._log(tool, params, started, "ok", lambda: count(result))
        return result

    def _log(
        self,
        tool: str,
        params: Mapping[str, object],
        started: float,
        outcome: str,
        counter: Callable[[], int] | None,
    ) -> None:
        try:
            elapsed = self._timer() - started
            when = self._clock()
            record: dict[str, object] = {
                "ts": when.isoformat(timespec="milliseconds"),
                "run": self.run_id,
                "source": self.source,
                "tool": tool,
                "params": scrub_params(params, when.tzinfo or UTC),
                "count": None if counter is None else int(counter()),
                "ms": round(elapsed * 1000),
                "outcome": outcome,
            }
            self.sink.write(record, when)
        except Exception as exc:  # fail-open: the audit never breaks a call
            if not self._warned:
                self._warned = True
                _warn(exc)


def audited[**P, R](fn: Callable[P, R]) -> Callable[P, R]:
    """Audits a method of an object holding `self._audit: Auditor`; tool name = method name.

    Parameters are the bound arguments with defaults applied, minus `self`.
    """
    signature = inspect.signature(fn)
    tool = fn.__name__

    @functools.wraps(fn)
    def wrapper(*args: P.args, **kwargs: P.kwargs) -> R:
        bound = signature.bind(*args, **kwargs)
        bound.apply_defaults()
        auditor: Auditor = args[0]._audit  # type: ignore[attr-defined]
        return auditor.call(tool, bound.arguments, lambda: fn(*args, **kwargs))

    return wrapper
