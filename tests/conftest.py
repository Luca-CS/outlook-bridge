from __future__ import annotations

import pytest

from outlook_bridge.audit import Auditor
from outlook_bridge.cli import ENV_DEBUG
from outlook_bridge.config import ENV_CONFIG, ENV_DATA_DIR, Config
from outlook_bridge.core import Bridge
from tests.fakes import NOW, FakeGateway, MemorySink, StepTimer


def isolate_environment(monkeypatch: pytest.MonkeyPatch, tmp_path) -> None:
    """Config and data dir inside `tmp_path`; the developer's shell switches are cleared."""
    monkeypatch.setenv(ENV_CONFIG, str(tmp_path / "absent.toml"))
    monkeypatch.setenv(ENV_DATA_DIR, str(tmp_path / "data"))
    # --debug equivalent: it appends tracebacks, so exact-stderr assertions would fail.
    monkeypatch.delenv(ENV_DEBUG, raising=False)


@pytest.fixture(autouse=True)
def _isolated_environment(tmp_path, monkeypatch):
    """No unit test may read the real config.toml or write to the real local profile."""
    isolate_environment(monkeypatch, tmp_path)


@pytest.fixture
def config() -> Config:
    return Config()


@pytest.fixture
def gateway() -> FakeGateway:
    return FakeGateway()


@pytest.fixture
def bridge(gateway: FakeGateway, config: Config) -> Bridge:
    return Bridge(gateway, config)


@pytest.fixture
def sink() -> MemorySink:
    return MemorySink()


@pytest.fixture
def audited_bridge(gateway: FakeGateway, config: Config, sink: MemorySink) -> Bridge:
    auditor = Auditor(sink, source="test", clock=lambda: NOW, timer=StepTimer(), run_id="run00001")
    return Bridge(gateway, config, audit=auditor, clock=lambda: NOW)
