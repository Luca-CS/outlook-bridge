"""COM gateway error paths, without Outlook: process detection is faked."""

import sys

import pytest

pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="pywin32 is Windows-only")

if sys.platform == "win32":
    from outlook_bridge import com_gateway
    from outlook_bridge.errors import NewOutlookDetected, OutlookNotRunning


@pytest.fixture
def gateway():
    gw = com_gateway.ComGateway(timeout_s=10)
    yield gw
    gw.close()


def test_new_outlook_is_reported(gateway, monkeypatch):
    monkeypatch.setattr(com_gateway, "_running_processes", lambda: {"olk.exe"})
    status = gateway.status()
    assert status.reachable is False
    assert status.classic is False
    assert "Nouvel Outlook" in status.message
    with pytest.raises(NewOutlookDetected):
        gateway.list_accounts()


def test_outlook_not_running(gateway, monkeypatch):
    monkeypatch.setattr(com_gateway, "_running_processes", lambda: {"explorer.exe"})
    status = gateway.status()
    assert status.reachable is False
    assert "pas lancé" in status.message
    with pytest.raises(OutlookNotRunning):
        gateway.list_accounts()


def test_worker_runs_calls_on_one_thread(gateway):
    import threading

    names = {gateway._worker.run(lambda: threading.current_thread().name) for _ in range(5)}
    assert names == {"outlook-com-sta"}


def test_helpers():
    assert com_gateway._split_names("A; B ;; C", 2) == ["A", "B"]
    assert com_gateway._snippet("a\r\n\r\n  b\tc") == "a b c"
    assert com_gateway._clean_body("a\r\n\r\n\r\n\r\nb") == "a\n\nb"
    assert com_gateway._int(None, 1) == 1
    assert com_gateway._int(0, 1) == 0

    class FakeFolder:
        FolderPath = "\\\\alice@example.com\\Boîte de réception\\Projets"

    assert com_gateway._folder_path(FakeFolder()) == "Boîte de réception/Projets"
