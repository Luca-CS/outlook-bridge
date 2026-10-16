import ast
import sys
from pathlib import Path

import pytest

from outlook_bridge.config import Config
from outlook_bridge.models import REDACT_MARGIN, SNIPPET_MAX
from outlook_bridge.redact import MAX_HOST_CHARS, REDACTED, Policy, shorten_urls
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
    mail = make_mail(1, snippet="a" * (SNIPPET_MAX + REDACT_MARGIN))
    assert len(p.redact_mail(mail).snippet) == SNIPPET_MAX


@pytest.mark.skipif(sys.platform != "win32", reason="pywin32 is Windows-only")
def test_gateway_snippet_keeps_the_redaction_margin():
    """The COM gateway must not cut the preview to SNIPPET_MAX before redaction: an IBAN
    straddling character 200 would lose its tail and escape the pattern."""
    from outlook_bridge.com_gateway import _snippet  # pure helper, no Outlook involved

    raw = "y" * 189 + " \r\n  " + IBAN + " fin " + "z" * 500
    snippet = _snippet(raw)
    assert len(snippet) == SNIPPET_MAX + REDACT_MARGIN
    assert IBAN in snippet
    shown = policy().redact_mail(make_mail(1, snippet=snippet)).snippet
    assert len(shown) == SNIPPET_MAX
    assert shown.startswith("y" * 189 + " " + REDACTED)
    assert not any(char.isdigit() for char in shown)


@pytest.mark.skipif(sys.platform != "win32", reason="pywin32 is Windows-only")
def test_outlook_capped_preview_drops_its_cut_tail():
    """Outlook cuts PR_PREVIEW at ~255 characters itself. Whitespace collapsing can pull an
    IBAN fragment cut there below SNIPPET_MAX, where redaction cannot recognise it."""
    from outlook_bridge.com_gateway import PREVIEW_CAP, _snippet

    head = "mot " + " \r\n " * 40  # collapses to a few characters
    fragment = IBAN[:14]  # "FR12 3456 7890": what Outlook's own cut leaves
    raw = (head + "x" * 255)[: 255 - len(fragment)] + fragment
    assert len(raw) >= PREVIEW_CAP
    shown = policy().redact_mail(make_mail(1, snippet=_snippet(raw, outlook_preview=True)))
    assert not any(char.isdigit() for char in shown.snippet)
    # A body-derived snippet is not capped by Outlook: nothing is dropped.
    assert _snippet("a b", outlook_preview=True) == "a b"


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


# ------------------------------------------------------------------ rule 5 in -m outlook tests
# The integration tests read the real mailbox and cannot run here: these static checks keep a
# failing `-m outlook` run from printing real content, and its output from staying on disk.

INTEGRATION = Path(__file__).resolve().parent / "integration"


def _bare(node: ast.expr) -> bool:
    """A local name, a literal, or a comparison or `and`/`or`/`not` of those: pytest has no
    sub-expression to introspect, so nothing read from Outlook can be printed."""
    if isinstance(node, ast.Name | ast.Constant):
        return True
    if isinstance(node, ast.UnaryOp) and isinstance(node.op, ast.Not):
        return _bare(node.operand)
    if isinstance(node, ast.BoolOp):
        return all(_bare(value) for value in node.values)
    if isinstance(node, ast.Compare):
        return all(_bare(side) for side in (node.left, *node.comparators))
    return False


def _unsafe_asserts(tree: ast.AST) -> list[int]:
    return [
        node.lineno
        for node in ast.walk(tree)
        if isinstance(node, ast.Assert)
        and not (_bare(node.test) and (node.msg is None or isinstance(node.msg, ast.Constant)))
    ]


@pytest.mark.parametrize("path", sorted(INTEGRATION.glob("test_*.py")), ids=lambda path: path.name)
def test_real_outlook_failures_print_no_content(path: Path):
    tree = ast.parse(path.read_text(encoding="utf-8"))
    assert "PYTEST_DONT_REWRITE" in (ast.get_docstring(tree) or "")
    assert _unsafe_asserts(tree) == []


@pytest.mark.parametrize(
    ("source", "safe"),
    [
        ("assert ok", True),
        ("assert count <= 3, 'trop'", True),
        ("assert not (a and b)", True),
        ("assert len(items) <= 3", False),
        ("assert result.stdout.isascii()", False),
        ("assert mail.received == wall_clock", False),
        ("assert ok, mail.subject", False),
    ],
)
def test_assert_checker(source: str, safe: bool):
    assert (_unsafe_asserts(ast.parse(source)) == []) is safe


def test_real_snapshots_are_deleted_even_on_failure():
    """The export test writes real, unexcluded snapshots: a `finally` removes them."""
    tree = ast.parse((INTEGRATION / "test_cli_outlook.py").read_text(encoding="utf-8"))
    exporting = [
        func
        for func in ast.walk(tree)
        if isinstance(func, ast.FunctionDef)
        and any(isinstance(n, ast.Constant) and n.value == "export" for n in ast.walk(func))
    ]
    assert exporting
    for func in exporting:
        cleaned = any(
            isinstance(node, ast.Try) and "rmtree(" in ast.unparse(node.finalbody)
            for node in ast.walk(func)
        )
        assert cleaned, func.name


# --- shorten_urls (Phase 3): long links become `<lien : domaine>` -------------------------

TRACKING = "https://click.news.example.com/ls/click?upn=" + "a1B2c3D4" * 30  # 285 chars


def test_long_bracketed_link_is_shortened_with_its_brackets():
    text = f"Lire l'article <{TRACKING}> puis la suite."
    assert shorten_urls(text) == "Lire l'article <lien : click.news.example.com> puis la suite."


def test_long_bare_link_keeps_its_trailing_punctuation():
    assert shorten_urls(f"Voir {TRACKING}.") == "Voir <lien : click.news.example.com>."
    assert shorten_urls(f"(voir {TRACKING}),") == "(voir <lien : click.news.example.com>),"


@pytest.mark.parametrize(
    "text",
    [
        "Le document : https://example.com/doc/42",
        "Lien <https://example.com/a?b=c>.",
        "mailto:alice@example.com et www.example.com/" + "x" * 100,  # no scheme: not a link
        "Aucun lien ici.",
        "",
    ],
)
def test_short_links_and_plain_text_are_unchanged(text: str):
    assert shorten_urls(text) == text


def test_threshold_is_inclusive():
    url = "https://example.com/" + "p" * 60
    assert shorten_urls(url, max_len=len(url)) == url
    assert shorten_urls(url, max_len=len(url) - 1) == "<lien : example.com>"


@pytest.mark.parametrize(
    ("url", "host"),
    [
        ("HTTPS://Tracking.Example.COM/" + "x" * 90, "tracking.example.com"),
        ("https://user:pw@example.com:8443/" + "x" * 90, "example.com"),
        ("https://[2001:db8::1]:443/" + "x" * 90, "[2001:db8::1]"),
        ("https://example.com?" + "x" * 90, "example.com"),
        (r"https://evil.example\@good.example/" + "x" * 90, "evil.example"),  # browsers: \ = /
        ("https:///" + "x" * 90, "?"),
        ("https://" + "h" * 100 + ".example.com/x", "h" * MAX_HOST_CHARS),
    ],
)
def test_host_is_lower_case_without_user_info_or_port(url: str, host: str):
    assert shorten_urls(url) == f"<lien : {host}>"


def test_many_links_in_a_newsletter_leave_room_for_the_text():
    body = "\n".join(f"Article {n} <{TRACKING}{n}>" for n in range(20))
    shortened = shorten_urls(body)
    assert "Article 19 <lien : click.news.example.com>" in shortened
    assert len(shortened) < len(body) // 5
    assert shorten_urls(shortened) == shortened  # idempotent


def test_shortening_drops_what_redaction_would_have_to_find_in_a_link():
    text = f"https://bank.example.com/rib?iban={IBAN.replace(' ', '')}&" + "x" * 60
    assert IBAN.replace(" ", "") not in shorten_urls(text)
