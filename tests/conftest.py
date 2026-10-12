from __future__ import annotations

import pytest

from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from tests.fakes import FakeGateway


@pytest.fixture
def config() -> Config:
    return Config()


@pytest.fixture
def gateway() -> FakeGateway:
    return FakeGateway()


@pytest.fixture
def bridge(gateway: FakeGateway, config: Config) -> Bridge:
    return Bridge(gateway, config)
