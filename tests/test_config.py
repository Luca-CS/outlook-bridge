import os
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from outlook_bridge.cli import ENV_DEBUG, exit_code_for
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
from tests.conftest import isolate_environment


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


# What PowerShell 5.1 or Notepad may write: the default excluded folders hold accents.
ACCENTED = 'folders_excluded = ["Éléments supprimés", "Courrier indésirable"]\n'


@pytest.mark.parametrize("encoding", ["cp1252", "utf-16"])
def test_non_utf8_file_is_a_config_error(tmp_path, encoding):
    path = tmp_path / "config.toml"
    path.write_bytes(ACCENTED.encode(encoding))
    with pytest.raises(ConfigError) as info:
        load_config(path)
    assert str(info.value) == (
        f"Configuration invalide ({path}) : fichier non encodé en UTF-8, réenregistrez-le en UTF-8."
    )
    assert exit_code_for(info.value) == 5  # the CLI's config exit code, not 1 (internal)


def test_utf8_with_bom_is_accepted(tmp_path):
    path = tmp_path / "config.toml"
    path.write_bytes(ACCENTED.encode("utf-8-sig"))  # PowerShell 5.1 `-Encoding UTF8`
    assert load_config(path).folders_excluded == ["Éléments supprimés", "Courrier indésirable"]


def test_unreadable_file_is_a_config_error(tmp_path):
    path = tmp_path / "config.toml"
    path.mkdir()  # a folder named config.toml: reading it raises an OSError
    with pytest.raises(ConfigError) as info:
        load_config(path)
    prefix = f"Configuration illisible ({path}) : "
    message = str(info.value)
    assert message.startswith(prefix)
    # Class name only: an OSError message repeats the path.
    assert message.removeprefix(prefix) in {"PermissionError.", "IsADirectoryError."}
    assert exit_code_for(info.value) == 5


@pytest.mark.parametrize(
    ("content", "field"),
    [
        ('senders_exclude = ["marqueur@example.com"]', "senders_exclude"),  # key typo
        ('subject_keywords_excluded = "MARQUEUR"', "subject_keywords_excluded"),
        ('accounts_allowed = ["ok", ["MARQUEUR"]]', "accounts_allowed.1"),
        ("redact_patterns = ['ok', 'MARQUEUR-(']", "redact_patterns"),
        ("redact_patterns = ['(?P=MARQUEUR)']", "redact_patterns"),  # re.error quotes names
        ("timezone = 'MARQUEUR/Zone'", "timezone"),
        ("working_hours = 'MARQUEUR'", "working_hours"),
        ("max_body_chars = 'MARQUEUR'", "max_body_chars"),
    ],
)
def test_invalid_values_are_never_echoed(tmp_path, content, field):
    """Exclusions say what the user hides from Claude: the error names the field only."""
    path = tmp_path / "config.toml"
    path.write_text(content + "\n", encoding="utf-8")
    with pytest.raises(ConfigError) as info:
        load_config(path)
    message = str(info.value)
    assert message.startswith(f"Configuration invalide ({path}) : {field} : ")
    assert "marqueur" not in message.casefold()


def test_validation_errors_are_listed_by_field(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text(
        "max_body_chars = 50000\nredact_patterns = ['ok', '(']\ntypo_key = 1\n",
        encoding="utf-8",
    )
    with pytest.raises(ConfigError) as info:
        load_config(path)
    assert str(info.value) == (
        f"Configuration invalide ({path}) : "
        "redact_patterns : Value error, regex invalide (motif 2, position 0) ; "
        "max_body_chars : Input should be less than or equal to 20000 ; "
        "typo_key : Extra inputs are not permitted"
    )


def test_toml_syntax_error_keeps_only_its_position(tmp_path):
    path = tmp_path / "config.toml"
    path.write_text('senders_excluded = ["a@example.com"]\ntimezone = = 1\n', encoding="utf-8")
    with pytest.raises(ConfigError, match=r"\(at line \d+, column \d+\)$") as info:
        load_config(path)
    assert "example.com" not in str(info.value)


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


def test_isolation_clears_the_debug_switch(tmp_path):
    """OUTLOOK_BRIDGE_DEBUG left set in the developer's shell must not leak into the tests."""
    with pytest.MonkeyPatch.context() as patch:
        patch.setenv(ENV_DEBUG, "1")
        isolate_environment(patch, tmp_path)
        assert ENV_DEBUG not in os.environ
    assert ENV_DEBUG not in os.environ  # the autouse fixture already cleared it here


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
