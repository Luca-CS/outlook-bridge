import json
from datetime import datetime

import pytest
from pydantic import ValidationError

from outlook_bridge.models import REDACT_MARGIN, SNIPPET_MAX, Event, MailDetail, MailSummary
from tests.fakes import NOW, make_event, make_mail


def test_mail_and_event_are_flagged_untrusted():
    mail = make_mail(1)
    event = make_event(1, NOW)
    assert mail.untrusted_content is True
    assert event.untrusted_content is True
    assert json.loads(mail.model_dump_json())["untrusted_content"] is True
    assert json.loads(event.model_dump_json())["untrusted_content"] is True


def test_untrusted_flag_cannot_be_disabled():
    with pytest.raises(ValidationError):
        make_mail(1, untrusted_content=False)


def test_naive_dates_are_rejected():
    with pytest.raises(ValidationError):
        make_mail(1, received=datetime(2026, 10, 12, 9, 0))
    with pytest.raises(ValidationError):
        make_event(1, datetime(2026, 10, 12, 9, 0))


def test_unknown_fields_are_rejected():
    with pytest.raises(ValidationError):
        make_mail(1, body_html="<p>x</p>")


def test_snippet_length_is_bounded():
    # The gateway boundary keeps a redaction margin; the core clips outputs to SNIPPET_MAX.
    raw_max = SNIPPET_MAX + REDACT_MARGIN
    assert len(make_mail(1, snippet="x" * raw_max).snippet) == raw_max
    with pytest.raises(ValidationError):
        make_mail(1, snippet="x" * (raw_max + 1))


def test_dates_serialise_with_offset():
    data = json.loads(make_mail(1).model_dump_json())
    assert data["received"].endswith("+02:00")


def test_detail_extends_summary():
    assert set(MailSummary.model_fields) < set(MailDetail.model_fields)
    assert "body_text" in Event.model_fields
