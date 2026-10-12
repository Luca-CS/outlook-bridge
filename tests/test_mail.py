from datetime import datetime, timedelta

import pytest

from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.errors import InvalidRequest, NotFound
from outlook_bridge.gateway import MailQuery
from outlook_bridge.mail import build_dasl_filter
from outlook_bridge.redact import REDACTED
from tests.fakes import NOW, PARIS, STORE, FakeGateway, make_mail


def bridge_with(mails, **config) -> tuple[Bridge, FakeGateway]:
    gw = FakeGateway(mails=mails)
    return Bridge(gw, Config(**config)), gw


def test_search_newest_first_with_limit():
    bridge, _ = bridge_with([make_mail(n) for n in range(1, 6)])
    result = bridge.mail.search_mail(limit=3)
    assert [m.id for m in result] == ["MAIL-1", "MAIL-2", "MAIL-3"]
    assert all(m.untrusted_content for m in result)


def test_limit_counts_visible_items_only():
    mails = [make_mail(n, sender_email="spam@example.net") for n in range(1, 4)]
    mails += [make_mail(n) for n in range(4, 8)]
    bridge, _ = bridge_with(mails, senders_excluded=["example.net"])
    result = bridge.mail.search_mail(limit=3)
    assert [m.id for m in result] == ["MAIL-4", "MAIL-5", "MAIL-6"]


def test_limit_is_capped_and_validated():
    bridge, gw = bridge_with([])
    bridge.mail.search_mail(limit=1000)
    assert gw.calls[-1][1].limit == 100
    with pytest.raises(InvalidRequest):
        bridge.mail.search_mail(limit=0)


def test_filters_are_forwarded_with_aware_dates():
    bridge, gw = bridge_with([])
    bridge.mail.search_mail(
        query="budget", since=datetime(2026, 10, 1), sender="bob", unread_only=True
    )
    query = gw.calls[-1][1]
    assert query.since == datetime(2026, 10, 1, tzinfo=PARIS)
    assert (query.query, query.sender, query.unread_only) == ("budget", "bob", True)


def test_search_filters_on_fake():
    mails = [make_mail(1, subject="Budget 2027"), make_mail(2), make_mail(3, unread=True)]
    bridge, _ = bridge_with(mails)
    assert [m.id for m in bridge.mail.search_mail(query="budget")] == ["MAIL-1"]
    assert {m.id for m in bridge.mail.search_mail(unread_only=True)} == {"MAIL-2", "MAIL-3"}
    since = NOW - timedelta(hours=2, minutes=30)
    assert [m.id for m in bridge.mail.search_mail(since=since)] == ["MAIL-1", "MAIL-2"]


def test_since_after_until_is_rejected():
    bridge, _ = bridge_with([])
    with pytest.raises(InvalidRequest):
        bridge.mail.search_mail(since=NOW, until=NOW - timedelta(days=1))


def test_excluded_folder_is_refused():
    bridge, _ = bridge_with([])
    with pytest.raises(NotFound):
        bridge.mail.search_mail(folder="Junk Email")


def test_search_results_are_redacted():
    bridge, _ = bridge_with([make_mail(1, subject="RIB FR1234567890123456789012345")])
    assert bridge.mail.search_mail()[0].subject == f"RIB {REDACTED}"


def test_get_mail_truncates_body():
    bridge, gw = bridge_with([make_mail(1, body_text="x" * 5000)])
    mail = bridge.mail.get_mail("MAIL-1", STORE, max_chars=100)
    assert len(mail.body_text) == 100
    assert mail.body_truncated is True
    assert gw.calls[-1][3] > 100  # extra margin fetched for redaction


def test_get_mail_default_and_max_chars():
    bridge, gw = bridge_with([make_mail(1, body_text="court")], max_body_chars=1234)
    mail = bridge.mail.get_mail("MAIL-1", STORE)
    assert mail.body_text == "court"
    assert mail.body_truncated is False
    bridge.mail.get_mail("MAIL-1", STORE, max_chars=10**6)
    assert gw.calls[-1][3] <= 20000 + 1000


def test_redaction_applies_across_the_truncation_point():
    iban = "FR1234567890123456789012345"
    # The IBAN straddles the 100-char cut: it must be redacted, not left as a fragment.
    bridge, _ = bridge_with([make_mail(1, body_text="a" * 89 + " " + iban)])
    mail = bridge.mail.get_mail("MAIL-1", STORE, max_chars=100)
    assert "FR12" not in mail.body_text
    assert mail.body_text == "a" * 89 + " " + REDACTED


def test_get_mail_by_id_respects_exclusions():
    mails = [make_mail(1, sender_email="x@example.net"), make_mail(2, folder="Junk Email")]
    bridge, _ = bridge_with(mails, senders_excluded=["example.net"])
    for mail_id in ("MAIL-1", "MAIL-2"):
        with pytest.raises(NotFound):
            bridge.mail.get_mail(mail_id, STORE)


def test_get_mail_unknown_id():
    bridge, _ = bridge_with([])
    with pytest.raises(NotFound):
        bridge.mail.get_mail("nope", STORE)


def test_get_thread_chronological_and_filtered():
    mails = [
        make_mail(1, conversation_id="C"),
        make_mail(2, conversation_id="C", subject="Confidentiel"),
        make_mail(3, conversation_id="C"),
        make_mail(4, conversation_id="OTHER"),
    ]
    bridge, _ = bridge_with(mails, subject_keywords_excluded=["confidentiel"])
    assert [m.id for m in bridge.mail.get_thread("C")] == ["MAIL-3", "MAIL-1"]
    assert [m.id for m in bridge.mail.get_thread("C", limit=1)] == ["MAIL-1"]


def test_accounts_and_folders_are_filtered(config):
    from outlook_bridge.models import Folder

    folders = [
        Folder(
            name="Inbox", path="Inbox", item_count=1, unread_count=0, account="alice@example.com"
        ),
        Folder(
            name="Junk",
            path="Junk Email",
            item_count=1,
            unread_count=0,
            account="alice@example.com",
        ),
        Folder(
            name="Inbox", path="Inbox", item_count=1, unread_count=0, account="perso@example.org"
        ),
    ]
    bridge = Bridge(FakeGateway(folders=folders), Config(accounts_allowed=["alice@example.com"]))
    assert [a.display_name for a in bridge.list_accounts()] == ["alice@example.com"]
    assert [a.display_name for a in bridge.status().accounts] == ["alice@example.com"]
    assert [(f.account, f.path) for f in bridge.list_folders()] == [("alice@example.com", "Inbox")]


class TestDaslFilter:
    def test_empty(self):
        assert build_dasl_filter(MailQuery(), PARIS) == ""

    def test_dates_are_utc(self):
        flt = build_dasl_filter(
            MailQuery(
                since=datetime(2026, 10, 9, tzinfo=PARIS),
                until=datetime(2026, 10, 10, tzinfo=PARIS),
            ),
            PARIS,
        )
        assert flt.startswith("@SQL=")
        assert "\"urn:schemas:httpmail:datereceived\" >= '2026-10-08 22:00'" in flt
        assert "\"urn:schemas:httpmail:datereceived\" < '2026-10-09 22:00'" in flt

    def test_text_filters_and_quote_escaping(self):
        flt = build_dasl_filter(MailQuery(query="l'offre", sender="bob", unread_only=True), PARIS)
        assert "LIKE '%l''offre%'" in flt
        assert "LIKE '%bob%'" in flt
        assert '"urn:schemas:httpmail:read" = 0' in flt
        assert flt.count(" AND ") == 2
