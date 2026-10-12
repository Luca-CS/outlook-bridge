from outlook_bridge.config import Config
from outlook_bridge.redact import REDACTED, Policy
from tests.fakes import NOW, make_event, make_mail

IBAN = "FR12 3456 7890 1234 5678 9012 345"  # invented, IBAN-shaped


def policy(**kwargs) -> Policy:
    return Policy(Config(**kwargs))


def test_default_pattern_redacts_french_iban():
    assert policy().redact(f"Virement sur {IBAN} merci") == f"Virement sur {REDACTED} merci"


def test_custom_patterns():
    p = policy(redact_patterns=[r"\b\d{3}-\d{4}\b", r"secret\w*"])
    assert p.redact("code 123-4567 et secrets") == f"code {REDACTED} et {REDACTED}"


def test_redact_mail_covers_subject_snippet_and_body():
    mail = make_mail(1, subject=f"RIB {IBAN}", snippet=f"voir {IBAN}", body_text=f"IBAN : {IBAN}")
    redacted = policy().redact_mail(mail)
    assert IBAN not in redacted.subject + redacted.snippet + redacted.body_text
    assert redacted.untrusted_content is True


def test_redact_event_covers_location_and_body():
    event = make_event(1, NOW, location=f"Salle {IBAN}", body_text=f"x {IBAN}")
    redacted = policy().redact_event(event)
    assert IBAN not in redacted.location + (redacted.body_text or "")


def test_redacted_snippet_stays_bounded():
    p = policy(redact_patterns=["a"])
    mail = make_mail(1, snippet="a" * 200)
    assert len(p.redact_mail(mail).snippet) == 200


def test_sender_rules_address_and_domain():
    p = policy(senders_excluded=["boss@example.com", "@example.net", "example.org"])
    assert p.sender_excluded("Boss@Example.com")
    assert not p.sender_excluded("other@example.com")
    assert p.sender_excluded("x@example.net")
    assert p.sender_excluded("x@lists.example.org")
    assert not p.sender_excluded("x@notexample.org")
    assert not p.sender_excluded("")


def test_folder_rules_match_any_segment():
    p = policy()
    assert p.folder_excluded("Courrier indésirable")
    assert p.folder_excluded("Junk Email/Sub")
    assert p.folder_excluded("éléments supprimés")
    assert not p.folder_excluded("Inbox")


def test_subject_keywords_are_case_insensitive():
    p = policy(subject_keywords_excluded=["Confidentiel"])
    assert p.subject_excluded("CONFIDENTIEL : budget")
    assert not p.subject_excluded("Réunion test")


def test_accounts_allowed():
    assert policy().account_allowed("anything")
    p = policy(accounts_allowed=["alice@example.com"])
    assert p.account_allowed("Alice@example.com")
    assert p.account_allowed("Alice Exemple", "alice@example.com")
    assert not p.account_allowed("perso@example.org")


def test_mail_and_event_visibility():
    p = policy(senders_excluded=["example.net"], subject_keywords_excluded=["paie"])
    assert p.mail_visible(make_mail(1))
    assert not p.mail_visible(make_mail(1, sender_email="x@example.net"))
    assert not p.mail_visible(make_mail(1, subject="Bulletin de paie"))
    assert not p.mail_visible(make_mail(1, folder="Junk Email"))
    assert p.event_visible(make_event(1, NOW))
    assert not p.event_visible(make_event(1, NOW, subject="Entretien paie"))
    assert not p.event_visible(make_event(1, NOW, calendar="Éléments supprimés"))
