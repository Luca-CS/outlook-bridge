from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from outlook_bridge.config import (
    ENV_CONFIG,
    ENV_DATA_DIR,
    Config,
    config_path,
    data_dir,
    load_config,
    log_dir,
    state_dir,
)
from outlook_bridge.errors import ConfigError, OutlookBridgeError


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
        'snapshot_dir = "G:\\Mon Drive"',  # basic string: `\M` is an invalid escape
    ],
)
def test_invalid_file_is_an_explicit_error(tmp_path, content):
    path = tmp_path / "config.toml"
    path.write_text(content, encoding="utf-8")
    with pytest.raises(ConfigError, match="Configuration invalide") as info:
        load_config(path)
    assert isinstance(info.value, OutlookBridgeError)


def test_snapshot_dir_as_literal_string(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text("snapshot_dir = 'G:\\Mon Drive\\outlook-bridge'\n", encoding="utf-8")
    assert load_config(path).snapshot_dir == "G:\\Mon Drive\\outlook-bridge"


def test_env_override(tmp_path, monkeypatch):
    monkeypatch.setenv(ENV_CONFIG, str(tmp_path / "x.toml"))
    assert config_path() == tmp_path / "x.toml"


def test_default_path_under_appdata(tmp_path, monkeypatch):
    monkeypatch.delenv(ENV_CONFIG, raising=False)
    monkeypatch.setenv("APPDATA", str(tmp_path))
    assert config_path() == tmp_path / "outlook-bridge" / "config.toml"


def test_tests_are_isolated_from_the_real_profile(tmp_path):
    """Autouse fixture (conftest): config and data dir point into the test's tmp_path."""
    assert config_path() == tmp_path / "absent.toml"
    assert load_config() == Config()
    assert data_dir() == tmp_path / "data"


def test_data_dir_env_override(tmp_path, monkeypatch):
    monkeypatch.setenv(ENV_DATA_DIR, str(tmp_path / "ailleurs"))
    assert data_dir() == tmp_path / "ailleurs"


def test_data_dir_under_localappdata(tmp_path, monkeypatch):
    monkeypatch.delenv(ENV_DATA_DIR, raising=False)
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert data_dir() == tmp_path / "outlook-bridge"


def test_data_dir_home_fallback(tmp_path, monkeypatch):
    monkeypatch.delenv(ENV_DATA_DIR, raising=False)
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setenv("USERPROFILE", str(tmp_path))  # Path.home() on Windows
    monkeypatch.setenv("HOME", str(tmp_path))  # Path.home() elsewhere
    assert data_dir() == tmp_path / "AppData" / "Local" / "outlook-bridge"


def test_log_and_state_dirs(tmp_path):
    assert log_dir() == tmp_path / "data" / "logs"
    assert state_dir() == tmp_path / "data" / "state"
