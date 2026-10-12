from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from outlook_bridge.config import ENV_CONFIG, Config, config_path, load_config
from outlook_bridge.errors import OutlookBridgeError


def test_defaults_match_spec():
    config = Config()
    assert config.timezone == "Europe/Paris"
    assert config.tz == ZoneInfo("Europe/Paris")
    assert "Courrier indésirable" in config.folders_excluded
    assert config.max_body_chars == 4000
    assert config.working_hours == "09:00-19:00"


def test_example_file_matches_defaults():
    example = Path(__file__).resolve().parents[1] / "config.example.toml"
    assert load_config(example) == Config()


def test_missing_file_gives_defaults(tmp_path):
    assert load_config(tmp_path / "absent.toml") == Config()


def test_load_file(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text(
        'senders_excluded = ["example.net"]\nmax_body_chars = 1000\n'
        'working_hours = "08:00-18:00"\n',
        encoding="utf-8",
    )
    config = load_config(path)
    assert config.senders_excluded == ["example.net"]
    assert config.max_body_chars == 1000


@pytest.mark.parametrize(
    "content",
    [
        "redact_patterns = ['(']",
        "timezone = 'Mars/Olympus'",
        "working_hours = '19:00-09:00'",
        "max_body_chars = 50000",
        "typo_key = 1",
        "not toml at all [",
    ],
)
def test_invalid_file_is_an_explicit_error(tmp_path, content):
    path = tmp_path / "config.toml"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(OutlookBridgeError, match="Configuration invalide"):
        load_config(path)


def test_env_override(tmp_path, monkeypatch):
    monkeypatch.setenv(ENV_CONFIG, str(tmp_path / "x.toml"))
    assert config_path() == tmp_path / "x.toml"


def test_default_path_under_appdata(tmp_path, monkeypatch):
    monkeypatch.delenv(ENV_CONFIG, raising=False)
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert config_path() == tmp_path / "outlook-bridge" / "config.toml"
