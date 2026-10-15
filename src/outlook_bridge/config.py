"""Exposure configuration: %APPDATA%\\outlook-bridge\\config.toml (see config.example.toml)."""

from __future__ import annotations

import os
import re
import tomllib
from functools import cached_property
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

from outlook_bridge.errors import ConfigError, InvalidRequest
from outlook_bridge.timeutil import parse_working_hours

ENV_CONFIG = "OUTLOOK_BRIDGE_CONFIG"
ENV_DATA_DIR = "OUTLOOK_BRIDGE_DATA_DIR"
MAX_BODY_CHARS_LIMIT = 20000


class Config(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    timezone: str = "Europe/Paris"
    accounts_allowed: list[str] = Field(default_factory=list)  # empty = all
    folders_excluded: list[str] = Field(
        default_factory=lambda: [
            "Deleted Items",
            "Junk Email",
            "Éléments supprimés",
            "Courrier indésirable",
        ]
    )
    senders_excluded: list[str] = Field(default_factory=list)  # addresses or domains
    subject_keywords_excluded: list[str] = Field(default_factory=list)
    # French IBAN: FR + 2 check digits + 23 alphanumerics, optionally grouped with spaces.
    redact_patterns: list[str] = Field(default_factory=lambda: [r"\bFR\d{2}(?: ?[0-9A-Z]){23}\b"])
    max_body_chars: int = Field(default=4000, ge=1, le=MAX_BODY_CHARS_LIMIT)
    snapshot_dir: str = ""
    working_hours: str = "09:00-19:00"

    @field_validator("timezone")
    @classmethod
    def _known_timezone(cls, value: str) -> str:
        try:
            ZoneInfo(value)
        except (ZoneInfoNotFoundError, ValueError) as exc:
            raise ValueError("fuseau inconnu (nom IANA attendu, ex. Europe/Paris)") from exc
        return value

    @field_validator("redact_patterns")
    @classmethod
    def _valid_regexes(cls, patterns: list[str]) -> list[str]:
        # By rank and position only: a pattern names what the user hides (re.error quotes
        # fragments of it, e.g. a group name).
        for rank, pattern in enumerate(patterns, 1):
            try:
                re.compile(pattern)
            except re.error as exc:
                where = "" if exc.pos is None else f", position {exc.pos}"
                raise ValueError(f"regex invalide (motif {rank}{where})") from exc
        return patterns

    @field_validator("working_hours")
    @classmethod
    def _valid_working_hours(cls, value: str) -> str:
        try:
            parse_working_hours(value)
        except InvalidRequest as exc:  # its message quotes the value
            raise ValueError("plage invalide (attendu HH:MM-HH:MM, début avant fin)") from exc
        return value

    @cached_property
    def tz(self) -> ZoneInfo:
        return ZoneInfo(self.timezone)


def config_path() -> Path:
    if override := os.environ.get(ENV_CONFIG):
        return Path(override)
    appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    return Path(appdata) / "outlook-bridge" / "config.toml"


def data_dir() -> Path:
    """Local, per-machine data (audit log, CLI state): %LOCALAPPDATA%\\outlook-bridge."""
    if override := os.environ.get(ENV_DATA_DIR):
        return Path(override)
    local = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(local) / "outlook-bridge"


def log_dir() -> Path:
    return data_dir() / "logs"


def state_dir() -> Path:
    return data_dir() / "state"


def _describe(exc: ValidationError) -> str:
    """`champ : message` per error, never the offending value: the lists hold what the user
    hides from Claude (senders, keywords, patterns)."""
    return " ; ".join(
        f"{'.'.join(map(str, error['loc'])) or 'configuration'} : {error['msg']}"
        for error in exc.errors(include_input=False, include_url=False, include_context=False)
    )


def load_config(path: Path | None = None) -> Config:
    """Missing file = defaults. Invalid file = explicit error (never silently ignored)."""
    path = path or config_path()
    if not path.exists():
        return Config()
    try:
        # utf-8-sig: also accepts the BOM that Notepad and PowerShell 5.1 write.
        text = path.read_bytes().decode("utf-8-sig")
    except UnicodeDecodeError as exc:  # ANSI (cp1252) or UTF-16 file
        raise ConfigError(
            f"Configuration invalide ({path}) : fichier non encodé en UTF-8, "
            "réenregistrez-le en UTF-8."
        ) from exc
    except OSError as exc:  # class name only: its message repeats the path
        raise ConfigError(f"Configuration illisible ({path}) : {type(exc).__name__}.") from exc
    try:
        return Config.model_validate(tomllib.loads(text))
    except tomllib.TOMLDecodeError as exc:  # reason + position (at most a key or one char)
        raise ConfigError(f"Configuration invalide ({path}) : {exc}") from exc
    except ValidationError as exc:
        raise ConfigError(f"Configuration invalide ({path}) : {_describe(exc)}") from exc
