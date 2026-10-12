"""Exposure configuration: %APPDATA%\\outlook-bridge\\config.toml (see config.example.toml)."""

from __future__ import annotations

import os
import re
import tomllib
from functools import cached_property
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

from outlook_bridge.errors import InvalidRequest, OutlookBridgeError
from outlook_bridge.timeutil import parse_working_hours

ENV_CONFIG = "OUTLOOK_BRIDGE_CONFIG"
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
            raise ValueError(f"fuseau inconnu : {value!r}") from exc
        return value

    @field_validator("redact_patterns")
    @classmethod
    def _valid_regexes(cls, patterns: list[str]) -> list[str]:
        for pattern in patterns:
            try:
                re.compile(pattern)
            except re.error as exc:
                raise ValueError(f"regex invalide {pattern!r} : {exc}") from exc
        return patterns

    @field_validator("working_hours")
    @classmethod
    def _valid_working_hours(cls, value: str) -> str:
        try:
            parse_working_hours(value)
        except InvalidRequest as exc:
            raise ValueError(str(exc)) from exc
        return value

    @cached_property
    def tz(self) -> ZoneInfo:
        return ZoneInfo(self.timezone)


def config_path() -> Path:
    if override := os.environ.get(ENV_CONFIG):
        return Path(override)
    appdata = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    return Path(appdata) / "outlook-bridge" / "config.toml"


def load_config(path: Path | None = None) -> Config:
    """Missing file = defaults. Invalid file = explicit error (never silently ignored)."""
    path = path or config_path()
    if not path.exists():
        return Config()
    try:
        data = tomllib.loads(path.read_text(encoding="utf-8"))
        return Config.model_validate(data)
    except (tomllib.TOMLDecodeError, ValidationError) as exc:
        raise OutlookBridgeError(f"Configuration invalide ({path}) : {exc}") from exc
