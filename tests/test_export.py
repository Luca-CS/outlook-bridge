"""Snapshot export: digest partition, calendar files, target checks, atomic writes, audit.

Synthetic data only (example.com, "Réunion test", the invented IBAN of test_redact.py).
"""

from __future__ import annotations

import contextlib
import itertools
import json
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pytest

from outlook_bridge import export, models
from outlook_bridge.audit import Auditor
from outlook_bridge.config import Config
from outlook_bridge.core import Bridge
from outlook_bridge.errors import ConfigError, ExportError, OutlookNotRunning
from outlook_bridge.events import find_meeting_url
from outlook_bridge.export import (
    CALENDAR_JSON,
    CALENDAR_MD,
    INBOX_FILE,
    SNAPSHOT_FILES,
    collect_calendar,
    collect_inbox,
    export_snapshots,
    inside_git_worktree,
    last_export,
    render_calendar_md,
    render_inbox_md,
    resolve_out_dir,
    snapshot_event,
    write_files,
)
from outlook_bridge.models import CalendarSnapshot, SnapshotEvent
from tests.fakes import NOW, PARIS, FakeGateway, MemorySink, StepTimer, make_event, make_mail

IBAN = "FR12 3456 7890 1234 5678 9012 345"  # invented, IBAN-shaped (as in test_redact.py)
HOSTILE = "Ligne1\n---\ngenerated_at: faux\n# Titre\x1b[2J"
HOSTILE_INLINE = r"Ligne1 --- generated\_at: faux # Titre"
MARKDOWN_TRAP = "![x](https://example.com/p.png) [lien](https://example.com) <img src=x> | `c`"
BODY_MARKER = "MARQUEUR-CORPS-COMPLET"
URL_MARKER = "MARQUEUR-URL"
# As ComGateway sees it: the join link (with its passcode) is the event's Location.
TEAMS_LOCATION = f"https://teams.microsoft.com/l/meetup-join/{URL_MARKER}?pwd=MARQUEUR-PWD"
RAW_IDS = ("MAIL-", "EVT-", "STORE-", "CONV-")


def at(day: int, hour: int = 0, minute: int = 0) -> datetime:
    return datetime(2026, 10, day, hour, minute, tzinfo=PARIS)


def mails() -> list:
    return [
        make_mail(
            1,
            subject="Facture test ancienne",
            unread=True,
            received=NOW - timedelta(days=10),
            has_attachments=True,
            sender_name="Carol Exemple",
            sender_email="carol@example.com",
            snippet="Extrait synthétique ancien",
            body_text=BODY_MARKER,
        ),
        make_mail(
            2,
            subject="Réunion test récente",
            unread=False,
            received=NOW - timedelta(days=1),
            snippet="Extrait synthétique récent",
            body_text=BODY_MARKER,
        ),
        make_mail(3, subject="Archive test lue", unread=False, received=NOW - timedelta(days=10)),
        make_mail(
            4,
            subject="Promo test exclue",
            unread=True,
            received=NOW - timedelta(hours=2),
            sender_email="promo@example.net",
        ),
        make_mail(5, subject=f"RIB {IBAN}", unread=True, received=NOW - timedelta(hours=3)),
        make_mail(
            6,
            subject=HOSTILE,
            unread=False,
            received=NOW - timedelta(hours=4),
            snippet=MARKDOWN_TRAP,
        ),
        make_mail(7, subject="Envoi test", unread=False, folder="Sent Items"),
    ]


def events() -> list:
    return [
        make_event(
            1,
            at(12, 14),
            location="Salle A",
            organizer="Bob Exemple",
            required_attendees=["Alice Exemple", "Carol Exemple"],
            body_text=BODY_MARKER,
        ),
        make_event(
            2,
            at(12, 16),
            30,
            busy_status="tentative",
            is_recurring=True,
            location=TEAMS_LOCATION,
            online_meeting_url=find_meeting_url(TEAMS_LOCATION, None),
        ),
        make_event(3, at(11, 22), 180, subject="Garde test"),  # started before the range
        make_event(4, at(14), 24 * 60, all_day=True, busy_status="oof", subject="Formation test"),
        make_event(5, at(16), 3 * 24 * 60, all_day=True, busy_status="free", subject="Séminaire"),
        make_event(6, at(19, 9, 30), 30, subject="Point hostile", location=HOSTILE),
        make_event(7, at(25, 10), subject="Dernier jour test"),  # after the DST change
        make_event(8, at(26, 10), subject="Hors plage test"),
    ]


@pytest.fixture
def config() -> Config:
    return Config(senders_excluded=["example.net"])


@pytest.fixture
def gateway() -> FakeGateway:
    return FakeGateway(mails=mails(), events=events())


@pytest.fixture
def out(tmp_path: Path) -> Path:
    return tmp_path / "snap"


@pytest.fixture
def written(audited_bridge: Bridge, out: Path) -> dict[str, str]:
    export_snapshots(audited_bridge, out)
    return {name: (out / name).read_bytes().decode("utf-8") for name in SNAPSHOT_FILES}


def section(text: str, heading_prefix: str) -> str:
    """Lines under the first heading starting with `heading_prefix`, up to the next `## `."""
    lines = text.split("\n")
    start = next(i for i, line in enumerate(lines) if line.startswith(heading_prefix))
    end = next((i for i in range(start + 1, len(lines)) if lines[i].startswith("## ")), len(lines))
    return "\n".join(lines[start + 1 : end])


def leftovers(folder: Path) -> list[str]:
    return sorted(p.name for p in folder.iterdir() if p.name.startswith(export.TMP_PREFIX))


# ---------------------------------------------------------------------------- models


def test_snapshot_event_is_an_allow_list():
    assert list(SnapshotEvent.model_fields) == [
        "untrusted_content",
        "subject",
        "start",
        "end",
        "all_day",
        "location",
        "organizer",
        "required_attendees",
        "optional_attendees",
        "busy_status",
        "response_status",
        "is_recurring",
        "has_online_meeting",
        "calendar",
        "account",
    ]
    event = events()[1]
    assert event.online_meeting_url and URL_MARKER in event.location  # the fixture is realistic
    snap = snapshot_event(event)
    assert snap.untrusted_content is True
    assert snap.has_online_meeting is True
    assert snap.subject == event.subject and snap.start == event.start
    assert snap.location == "[lien retiré]"
    assert URL_MARKER not in snap.model_dump_json() and "MARQUEUR-PWD" not in snap.model_dump_json()
    assert snapshot_event(events()[0]).has_online_meeting is False
    assert snapshot_event(events()[0]).location == "Salle A"
    assert SnapshotEvent.from_event(events()[0]).has_online_meeting is False


def test_snapshot_event_copies_attendee_lists():
    event = events()[0]
    snap = SnapshotEvent.from_event(event)
    event.required_attendees.append("Dave Exemple")
    assert snap.required_attendees == ["Alice Exemple", "Carol Exemple"]


def test_snapshot_models_reject_unknown_fields():
    data = SnapshotEvent.from_event(events()[0]).model_dump()
    with pytest.raises(ValueError):
        SnapshotEvent(**data, id="EVT-1")
    with pytest.raises(ValueError):
        SnapshotEvent(**{**data, "untrusted_content": False})


# ---------------------------------------------------------------------------- inbox digest


def test_digest_partition(written):
    text = written[INBOX_FILE]
    unread = section(text, "## Non lus (2)")
    others = section(text, "## Autres mails reçus depuis le vendredi 9 octobre (2)")
    assert "**Facture test ancienne** · non lu, pièce jointe" in unread
    assert "ven. 02/10 12:00 · Carol Exemple (carol@example.com)" in unread
    assert r"**RIB \[masqué\]** · non lu" in unread
    assert "Réunion test récente" in others and "Réunion test récente" not in unread
    assert HOSTILE_INLINE in others
    assert "Archive test lue" not in text  # read and older than 3 days
    assert "Promo test exclue" not in text and "example.net" not in text  # excluded sender
    assert "Envoi test" not in text  # another folder
    assert IBAN not in text
    assert "> Extrait synthétique récent" in others
    assert BODY_MARKER not in text
    assert "compte :" not in text  # a single account
    assert "\nunread_count: 2\nother_count: 2\ntruncated: false\n" in text
    assert "since: 2026-10-09T00:00:00+02:00\n" in text


def test_digest_dedupes_by_id_not_by_flag():
    class FlippingGateway(FakeGateway):
        """Mail X is unread for the first query and already read for the second."""

        def __init__(self) -> None:
            super().__init__(mails=[make_mail(1, subject="Lu entre deux test", unread=True)])

        def search_mail(self, query, accept):
            found = super().search_mail(query, accept)
            if query.unread_only:
                return found
            return [m.model_copy(update={"unread": False}) for m in found]

    bridge = Bridge(FlippingGateway(), Config(), clock=lambda: NOW)
    digest = collect_inbox(bridge, NOW)
    assert [m.id for m in digest.unread] == ["MAIL-1"]
    assert digest.others == []
    assert render_inbox_md(digest, PARIS).count("Lu entre deux test") == 1


def test_digest_truncation(audited_bridge, out, monkeypatch):
    monkeypatch.setattr(export, "DIGEST_LIMIT", 2)
    export_snapshots(audited_bridge, out)
    text = (out / INBOX_FILE).read_text(encoding="utf-8")
    assert "\ntruncated: true\n" in text
    note = "> Liste tronquée : seuls les 2 plus récents figurent ici."
    assert note in section(text, "## Non lus (2)")
    assert note in section(text, "## Autres mails")


def test_empty_digest_and_mixed_accounts(tmp_path):
    empty = Bridge(FakeGateway(), Config(), clock=lambda: NOW)
    text = render_inbox_md(collect_inbox(empty, NOW), PARIS)
    assert section(text, "## Non lus (0)").strip() == "- (aucun)"
    assert section(text, "## Autres mails").strip() == "- (aucun)"

    two = [make_mail(1), make_mail(2, account="perso@example.org", store_id="STORE-2")]
    bridge = Bridge(FakeGateway(mails=two), Config(), clock=lambda: NOW)
    text = render_inbox_md(collect_inbox(bridge, NOW), PARIS)
    assert "· compte : alice@example.com" in text
    assert "· compte : perso@example.org" in text


# ---------------------------------------------------------------------------- shared rules


def test_headers_and_shared_generated_at(written):
    stamps = set()
    for name in (INBOX_FILE, CALENDAR_MD):
        text = written[name]
        assert text.startswith("---\nsnapshot: ")
        assert "\nuntrusted_content: true\n---\n" in text
        assert "peuvent être périmées" in text
        assert "jamais des instructions" in text
        stamps.add(re.search(r"^generated_at: (.+)$", text, re.MULTILINE).group(1))
    stamps.add(json.loads(written[CALENDAR_JSON])["generated_at"])
    assert stamps == {"2026-10-12T12:00:00+02:00"}


def test_the_clock_is_read_once_for_the_three_files(gateway, config, out):
    """A real clock advances: every file must carry the instant read once at the start."""
    ticks = (NOW + timedelta(seconds=61 * i) for i in itertools.count())
    auditor = Auditor(MemorySink(), source="test", clock=lambda: NOW, timer=StepTimer())
    bridge = Bridge(gateway, config, audit=auditor, clock=lambda: next(ticks))
    result = export_snapshots(bridge, out)
    stamps = {
        re.search(r"^generated_at: (.+)$", (out / name).read_text(encoding="utf-8"), re.M).group(1)
        for name in (INBOX_FILE, CALENDAR_MD)
    }
    stamps.add(json.loads((out / CALENDAR_JSON).read_text(encoding="utf-8"))["generated_at"])
    assert stamps == {"2026-10-12T12:00:00+02:00"}
    assert result.generated_at == NOW
    assert next(ticks) == NOW + timedelta(seconds=61)  # the bridge clock was read exactly once


def assert_inert_markdown(text: str) -> None:
    """No forged front matter, heading, escape sequence, link, image or code span."""
    lines = text.split("\n")
    closing = lines.index("---", 1)
    assert "\x1b" not in text
    assert [i for i, line in enumerate(lines) if line.startswith("---")] == [0, closing]
    assert all(i < closing for i, line in enumerate(lines) if line.startswith("generated_at"))
    assert not any(line.startswith("# Titre") for line in lines)
    markdown_it = pytest.importorskip("markdown_it")
    html = markdown_it.MarkdownIt("commonmark").render(text)
    assert "<img" not in html and "<a " not in html and "<code>" not in html


def test_hostile_text_stays_on_one_line(written):
    for name in (INBOX_FILE, CALENDAR_MD):
        lines = written[name].split("\n")
        assert sum(HOSTILE_INLINE in line for line in lines) == 1
        assert_inert_markdown(written[name])


def test_untrusted_markdown_cannot_build_links_or_html(written):
    markdown_it = pytest.importorskip("markdown_it")
    for name in (INBOX_FILE, CALENDAR_MD):
        html = markdown_it.MarkdownIt("commonmark").render(written[name])
        assert "<img" not in html and "<a " not in html and "<code>" not in html
    assert r"!\[x\](https://example.com/p.png)" in written[INBOX_FILE]


HOSTILE_VALUES = pytest.mark.parametrize("value", [HOSTILE, MARKDOWN_TRAP], ids=["lignes", "md"])


@HOSTILE_VALUES
@pytest.mark.parametrize("field", ["subject", "sender_name", "sender_email", "snippet", "account"])
def test_every_untrusted_mail_field_is_inert(field: str, value: str):
    hostile = make_mail(1, unread=True, **{field: value})
    other = make_mail(2, account="perso@example.org", store_id="STORE-2")  # shows the accounts
    bridge = Bridge(FakeGateway(mails=[hostile, other]), Config(), clock=lambda: NOW)
    text = render_inbox_md(collect_inbox(bridge, NOW), PARIS)
    assert_inert_markdown(text)
    if value == HOSTILE:
        assert sum(HOSTILE_INLINE in line for line in text.split("\n")) == 1
    else:
        assert r"!\[x\](https://example.com/p.png)" in text


@HOSTILE_VALUES
@pytest.mark.parametrize("field", ["subject", "location", "organizer", "account"])
def test_every_untrusted_event_field_is_inert(field: str, value: str):
    hostile = make_event(1, at(13, 10), **{field: value})
    other = make_event(2, at(13, 14), account="perso@example.org", store_id="STORE-2")
    bridge = Bridge(FakeGateway(events=[hostile, other]), Config(), clock=lambda: NOW)
    text = render_calendar_md(collect_calendar(bridge, NOW), PARIS)
    assert_inert_markdown(text)
    if value == HOSTILE:
        assert sum(HOSTILE_INLINE in line for line in text.split("\n")) == 1


def test_no_ids_bodies_or_meeting_urls_anywhere(written):
    for text in written.values():
        assert BODY_MARKER not in text
        assert URL_MARKER not in text and "MARQUEUR-PWD" not in text
        assert "teams.microsoft.com" not in text
        assert not [raw for raw in RAW_IDS if raw in text]


def test_links_in_locations_never_reach_the_snapshot(out):
    jitsi = f"Salle B ou HTTPS://meet.jit.si/{URL_MARKER}-2"  # a provider find_meeting_url ignores
    site = f"Plan : www.example.com/{URL_MARKER}-3"
    hostile = [
        make_event(1, at(13, 10), location=TEAMS_LOCATION, online_meeting_url=TEAMS_LOCATION),
        make_event(2, at(13, 14), location=jitsi),
        make_event(3, at(13, 16), location=site),
        make_event(4, at(13, 17), location="Salle A"),
    ]
    export_snapshots(Bridge(FakeGateway(events=hostile), Config(), clock=lambda: NOW), out)
    texts = {name: (out / name).read_text(encoding="utf-8") for name in SNAPSHOT_FILES}
    for text in texts.values():
        assert URL_MARKER not in text and "MARQUEUR-PWD" not in text
        assert "teams.microsoft.com" not in text and "meet.jit.si" not in text
    locations = {e["subject"]: e["location"] for e in json.loads(texts[CALENDAR_JSON])["events"]}
    assert locations == {
        "Point projet 1": "[lien retiré]",
        "Point projet 2": "Salle B ou [lien retiré]",
        "Point projet 3": "Plan : [lien retiré]",
        "Point projet 4": "Salle A",
    }
    calendar_md = texts[CALENDAR_MD]
    assert r"- 10:00-11:00 · **Point projet 1** · \[lien retiré\] · visio" in calendar_md
    assert r"- 14:00-15:00 · **Point projet 2** · Salle B ou \[lien retiré\]" in calendar_md
    assert "- 17:00-18:00 · **Point projet 4** · Salle A" in calendar_md
    # the Markdown renderer removes links too, even from a snapshot built without them removed
    bridge = Bridge(FakeGateway(events=hostile), Config(), clock=lambda: NOW)
    unfiltered = collect_calendar(bridge, NOW).model_copy(
        update={"events": [SnapshotEvent.from_event(e) for e in hostile]}
    )
    rendered = render_calendar_md(unfiltered, PARIS)
    assert URL_MARKER not in rendered and "MARQUEUR-PWD" not in rendered
    assert r"Salle B ou \[lien retiré\]" in rendered


# ---------------------------------------------------------------------------- calendar


def test_calendar_md_lists_all_fourteen_days(written):
    text = written[CALENDAR_MD]
    headings = [line for line in text.split("\n") if line.startswith("## ")]
    assert headings[-1] == "## Créneaux libres"
    days = headings[:-1]
    assert len(days) == 14
    assert days[0] == "## lundi 12 octobre 2026"
    assert days[-1] == "## dimanche 25 octobre 2026"
    assert "# Agenda du lundi 12 octobre au dimanche 25 octobre 2026\n" in text
    assert "range_end: 2026-10-26T00:00:00+01:00\n" in text
    assert section(text, "## mardi 13 octobre 2026").strip() == "- (aucun événement)"


def test_calendar_md_event_lines(written):
    text = written[CALENDAR_MD]
    monday = section(text, "## lundi 12 octobre 2026")
    assert monday.startswith("- dim. 11/10 22:00 au 01:00 (suite) · **Garde test**")
    assert (
        "- 14:00-15:00 · **Point projet 1** · Salle A · org. Bob Exemple · 2 participants" in monday
    )
    assert (
        r"- 16:00-16:30 · **Point projet 2** · \[lien retiré\] · provisoire, récurrent, visio"
        in monday
    )
    assert "- journée · **Formation test** · absent" in section(text, "## mercredi 14 octobre")
    # a multi-day event is under each of its days, the following ones marked "(suite)"
    seminar = [
        section(text, f"## {day} octobre").strip()
        for day in ("vendredi 16", "samedi 17", "dimanche 18")
    ]
    assert seminar == [
        "- journées jusqu'au dim. 18/10 · **Séminaire** · libre",
        "- journées jusqu'au dim. 18/10 (suite) · **Séminaire** · libre",
        "- journée (suite) · **Séminaire** · libre",
    ]
    assert text.count("Séminaire") == 3
    assert "\nevent_count: 7\n" in text  # each event is still counted once
    assert "Dernier jour test" in section(text, "## dimanche 25 octobre 2026")
    assert "Hors plage test" not in text


def test_a_day_covered_by_an_event_is_never_empty():
    # lun. 19/10 to ven. 23/10 all day, then mar. 20/10 18:00 to mer. 21/10 20:00
    leave = make_event(1, at(19), 5 * 1440, all_day=True, busy_status="oof", subject="Congés test")
    on_call = make_event(2, at(20, 18), 26 * 60, subject="Astreinte test")
    bridge = Bridge(FakeGateway(events=[leave, on_call]), Config(), clock=lambda: NOW)
    text = render_calendar_md(collect_calendar(bridge, NOW), PARIS)
    days = {
        day: section(text, f"## {day} octobre 2026").strip()
        for day in ("lundi 19", "mardi 20", "mercredi 21", "jeudi 22", "vendredi 23", "samedi 24")
    }
    assert days == {
        "lundi 19": "- journées jusqu'au ven. 23/10 · **Congés test** · absent",
        "mardi 20": "- journées jusqu'au ven. 23/10 (suite) · **Congés test** · absent\n"
        "- 18:00 au mer. 21/10 20:00 · **Astreinte test**",
        "mercredi 21": "- journées jusqu'au ven. 23/10 (suite) · **Congés test** · absent\n"
        "- mar. 20/10 18:00 au 20:00 (suite) · **Astreinte test**",
        "jeudi 22": "- journées jusqu'au ven. 23/10 (suite) · **Congés test** · absent",
        "vendredi 23": "- journée (suite) · **Congés test** · absent",
        "samedi 24": "- (aucun événement)",
    }
    assert "\nevent_count: 2\n" in text


def test_calendar_md_free_slots(written):
    slots = section(written[CALENDAR_MD], "## Créneaux libres")
    assert "à partir du lundi 12 octobre 2026 à 12:00" in slots
    assert "Au moins 60 min, 09:00-19:00, jours ouvrés" in slots
    assert "- lun. 12/10 : 12:00-14:00 (2 h), 15:00-16:00 (1 h), 16:30-19:00 (2 h 30)" in slots
    assert "- mer. 14/10 : aucun" in slots  # all-day out of office
    assert "- lun. 19/10 : 10:00-19:00 (9 h)" in slots
    assert "sam." not in slots and "dim." not in slots


SLOTS_NOTE = (
    "Les événements masqués par un mot-clé du sujet comptent comme occupés ; les comptes et "
    "calendriers exclus par la configuration sont ignorés."
)


def test_free_slots_note_says_what_counts_as_busy():
    hidden = [
        make_event(1, at(13, 10), calendar="Calendar/Perso", subject="Perso test"),
        make_event(2, at(13, 15), subject="Rendez-vous confidentiel test"),
        make_event(3, at(14, 10), account="perso@example.org", subject="Autre compte test"),
    ]
    config = Config(
        accounts_allowed=["alice@example.com"],
        folders_excluded=["Perso"],
        subject_keywords_excluded=["confidentiel"],
    )
    bridge = Bridge(FakeGateway(events=hidden), config, clock=lambda: NOW)
    text = render_calendar_md(collect_calendar(bridge, NOW), PARIS)
    slots = section(text, "## Créneaux libres")
    assert SLOTS_NOTE in slots
    # the sentence is true: only the keyword-hidden event blocks its hour
    assert "- mar. 13/10 : 09:00-15:00 (6 h), 16:00-19:00 (3 h)" in slots
    assert "- mer. 14/10 : 09:00-19:00 (10 h)" in slots
    assert "test" not in section(text, "## mardi 13 octobre") + section(text, "## mercredi 14")


def test_calendar_json(written):
    raw = written[CALENDAR_JSON]
    snapshot = CalendarSnapshot.model_validate_json(raw)
    data = json.loads(raw)
    assert snapshot.snapshot == "calendar_next14d" and data["schema_version"] == 1
    assert data["untrusted_content"] is True
    assert data["timezone"] == "Europe/Paris"
    assert data["range_start"] == "2026-10-12T00:00:00+02:00"
    assert data["range_end"] == "2026-10-26T00:00:00+01:00"
    forbidden = {"id", "store_id", "online_meeting_url", "body_text", "body_truncated"}
    assert all(forbidden.isdisjoint(event) for event in data["events"])
    assert all(event["untrusted_content"] is True for event in data["events"])
    by_subject = {event["subject"]: event for event in data["events"]}
    assert by_subject["Point projet 2"]["has_online_meeting"] is True
    assert by_subject["Point projet 1"]["has_online_meeting"] is False
    assert by_subject["Point projet 1"]["required_attendees"] == ["Alice Exemple", "Carol Exemple"]
    assert by_subject["Dernier jour test"]["start"] == "2026-10-25T10:00:00+01:00"
    assert "Hors plage test" not in by_subject
    assert "Garde test" in by_subject
    assert data["free_slots"][0] == {
        "start": "2026-10-12T12:00:00+02:00",
        "end": "2026-10-12T14:00:00+02:00",
        "duration_min": 120,
    }
    assert data["free_slots_min_duration_min"] == 60
    assert data["working_hours"] == "09:00-19:00"
    assert "Instantané" in data["notice"]


def test_free_slots_start_at_the_next_quarter_hour(gateway, config, tmp_path):
    bridge = Bridge(gateway, config, clock=lambda: NOW + timedelta(minutes=7))
    export_snapshots(bridge, tmp_path / "snap")
    data = json.loads((tmp_path / "snap" / CALENDAR_JSON).read_text(encoding="utf-8"))
    assert data["generated_at"] == "2026-10-12T12:07:00+02:00"
    assert data["free_slots"][0]["start"] == "2026-10-12T12:15:00+02:00"
    calendar_md = (tmp_path / "snap" / CALENDAR_MD).read_text(encoding="utf-8")
    assert "à partir du lundi 12 octobre 2026 à 12:15" in calendar_md


# ---------------------------------------------------------------------------- target


def test_empty_snapshot_dir_is_a_config_error():
    with pytest.raises(ConfigError, match="Aucun dossier de snapshots"):
        resolve_out_dir(Config(), None)


def test_relative_snapshot_dir_is_a_config_error():
    with pytest.raises(ConfigError, match="chemin absolu"):
        resolve_out_dir(Config(snapshot_dir="snapshots"), None)


def test_missing_parent_is_refused_and_nothing_is_created(tmp_path):
    target = tmp_path / "absent" / "snap"
    with pytest.raises(ExportError, match="inaccessible"):
        resolve_out_dir(Config(snapshot_dir=str(target)), None)
    assert not (tmp_path / "absent").exists()


@pytest.mark.parametrize("git_is_file", [False, True], ids=["git-dir", "git-file"])
def test_git_worktree_is_refused(tmp_path, git_is_file):
    repo = tmp_path / "repo"
    (repo / "docs").mkdir(parents=True)
    if git_is_file:
        (repo / ".git").write_text("gitdir: ailleurs\n", encoding="utf-8")
    else:
        (repo / ".git").mkdir()
    target = repo / "docs" / "snap"
    assert inside_git_worktree(target)
    with pytest.raises(ExportError, match="dépôt git"):
        resolve_out_dir(Config(), target)
    assert not target.exists()


@pytest.mark.skipif(sys.platform != "win32", reason="NTFS junctions are Windows-only")
def test_git_worktree_behind_a_junction_is_refused(tmp_path):
    """A junction needs no admin right: the target must be checked once resolved too."""
    winapi = pytest.importorskip("_winapi")
    repo = tmp_path / "repo"
    (repo / ".git").mkdir(parents=True)
    (repo / "sub").mkdir()
    link = tmp_path / "lien"
    winapi.CreateJunction(str(repo / "sub"), str(link))
    try:
        target = link / "snap"
        assert inside_git_worktree(target)
        with pytest.raises(ExportError, match="dépôt git"):
            resolve_out_dir(Config(snapshot_dir=str(target)), None)
        with pytest.raises(ExportError, match="dépôt git"):
            resolve_out_dir(Config(), target)
        assert not (repo / "sub" / "snap").exists()
    finally:
        with contextlib.suppress(OSError):
            os.rmdir(link)  # removes the junction, not its target


def test_out_overrides_the_config(tmp_path):
    config = Config(snapshot_dir=str(tmp_path / "cfg"))
    assert not inside_git_worktree(tmp_path / "cli")
    assert resolve_out_dir(config, tmp_path / "cli") == tmp_path / "cli"
    assert (tmp_path / "cli").is_dir()
    assert not (tmp_path / "cfg").exists()
    assert resolve_out_dir(config, None) == tmp_path / "cfg"
    assert resolve_out_dir(config, None) == tmp_path / "cfg"  # an existing folder is fine


def test_relative_out_is_resolved_against_the_current_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    assert resolve_out_dir(Config(), Path("snap")) == tmp_path / "snap"


@pytest.mark.skipif(sys.platform != "win32", reason="%VAR% syntax is Windows-only")
def test_environment_variables_are_expanded(tmp_path, monkeypatch):
    monkeypatch.setenv("OB_TEST_SNAPSHOTS", str(tmp_path))
    config = Config(snapshot_dir=r"%OB_TEST_SNAPSHOTS%\snap")
    assert resolve_out_dir(config, None) == tmp_path / "snap"


def test_target_that_is_a_file_is_an_export_error(tmp_path):
    (tmp_path / "snap").write_text("x", encoding="utf-8")
    with pytest.raises(ExportError, match="impossible"):
        resolve_out_dir(Config(), tmp_path / "snap")


# ---------------------------------------------------------------------------- writing


def test_export_writes_three_utf8_lf_files(audited_bridge, out):
    result = export_snapshots(audited_bridge, out)
    assert [f.name for f in result.files] == list(SNAPSHOT_FILES)
    assert result.out_dir == str(out)
    assert result.generated_at == NOW
    for exported in result.files:
        raw = (out / exported.name).read_bytes()
        assert exported.size_bytes == len(raw)
        assert b"\r" not in raw and not raw.startswith(b"\xef\xbb\xbf")
        assert raw.endswith(b"\n")
        raw.decode("utf-8")
    assert sorted(p.name for p in out.iterdir()) == sorted(SNAPSHOT_FILES)
    slots = len(json.loads((out / CALENDAR_JSON).read_text(encoding="utf-8"))["free_slots"])
    assert result.files[0].summary == "2 non lus, 2 autres mails"
    assert result.files[1].summary == f"7 événements, {slots} créneaux libres"
    assert result.files[2].summary == result.files[1].summary


def test_running_twice_overwrites(gateway, config, out):
    export_snapshots(Bridge(gateway, config, clock=lambda: NOW), out)
    later = NOW + timedelta(hours=1)
    export_snapshots(Bridge(gateway, config, clock=lambda: later), out)
    for name in SNAPSHOT_FILES:
        assert "2026-10-12T13:00:00+02:00" in (out / name).read_text(encoding="utf-8")
    assert sorted(p.name for p in out.iterdir()) == sorted(SNAPSHOT_FILES)


def test_outlook_failure_writes_nothing(audited_bridge, gateway, config, out):
    export_snapshots(audited_bridge, out)
    before = {name: (out / name).read_bytes() for name in SNAPSHOT_FILES}
    gateway.errors["list_events"] = OutlookNotRunning()
    later = Bridge(gateway, config, clock=lambda: NOW + timedelta(hours=1))
    with pytest.raises(OutlookNotRunning):
        export_snapshots(later, out)
    assert {name: (out / name).read_bytes() for name in SNAPSHOT_FILES} == before
    assert leftovers(out) == []


def test_staging_failure_changes_no_final_file(audited_bridge, out, monkeypatch):
    out.mkdir()
    for name in SNAPSHOT_FILES:
        (out / name).write_bytes(b"ancien\n")
    real_write_text = Path.write_text
    staged: list[str] = []

    def disk_full_on_second(self, *args, **kwargs):
        staged.append(self.name)
        if len(staged) == 2:
            raise OSError("MARQUEUR-DISQUE")
        return real_write_text(self, *args, **kwargs)

    monkeypatch.setattr(Path, "write_text", disk_full_on_second)
    with pytest.raises(ExportError, match=r"Écriture impossible dans .* \(OSError\)\.") as info:
        export_snapshots(audited_bridge, out)
    assert "MARQUEUR" not in str(info.value)
    assert len(staged) == 2
    assert re.fullmatch(r"~\$inbox_digest\.md\.\w+\.tmp", staged[0])
    assert re.fullmatch(r"~\$calendar_next14d\.md\.\w+\.tmp", staged[1])
    assert all((out / name).read_bytes() == b"ancien\n" for name in SNAPSHOT_FILES)
    assert leftovers(out) == []


def test_overlapping_exports_keep_their_own_temp_files(tmp_path, monkeypatch):
    """A second export (manual run during the scheduled one) stages and commits while the
    first one is swapping: neither takes nor deletes the other's temp files."""
    real_replace = os.replace
    swaps: list[str] = []

    def other_export_in_between(src, dst):
        swaps.append(Path(dst).name)
        if len(swaps) == 1:
            assert write_files(tmp_path, {"a.md": "B\n", "b.md": "B\n"}) == [
                ("a.md", 2),
                ("b.md", 2),
            ]
        return real_replace(src, dst)

    monkeypatch.setattr(export.os, "replace", other_export_in_between)
    assert write_files(tmp_path, {"a.md": "A\n", "b.md": "A\n"}) == [("a.md", 2), ("b.md", 2)]
    assert swaps == ["a.md", "a.md", "b.md", "b.md"]
    assert (tmp_path / "a.md").read_text(encoding="utf-8") == "A\n"
    assert (tmp_path / "b.md").read_text(encoding="utf-8") == "A\n"
    assert leftovers(tmp_path) == []


def test_temp_files_of_another_run_are_left_alone(tmp_path):
    foreign = tmp_path / f"{export.TMP_PREFIX}a.md.tmp"
    foreign.write_text("autre export en cours\n", encoding="utf-8")
    write_files(tmp_path, {"a.md": "A\n"})
    assert foreign.read_text(encoding="utf-8") == "autre export en cours\n"
    assert (tmp_path / "a.md").read_text(encoding="utf-8") == "A\n"
    assert leftovers(tmp_path) == [foreign.name]


def test_replace_is_retried_on_permission_error(audited_bridge, out, monkeypatch):
    real_replace = os.replace
    failures = [PermissionError("verrou"), PermissionError("verrou")]

    def locked_twice(src, dst):
        if failures:
            raise failures.pop(0)
        return real_replace(src, dst)

    monkeypatch.setattr(export.os, "replace", locked_twice)
    sleeps: list[float] = []
    result = export_snapshots(audited_bridge, out, sleep=sleeps.append)
    assert sleeps == [0.5, 0.5]
    assert len(result.files) == 3
    assert sorted(p.name for p in out.iterdir()) == sorted(SNAPSHOT_FILES)


def test_permanent_lock_names_the_files_already_updated(audited_bridge, out, monkeypatch):
    real_replace = os.replace

    def calendar_locked(src, dst):
        if Path(dst).name == CALENDAR_MD:
            raise PermissionError("verrou")
        return real_replace(src, dst)

    monkeypatch.setattr(export.os, "replace", calendar_locked)
    sleeps: list[float] = []
    with pytest.raises(ExportError) as info:
        export_snapshots(audited_bridge, out, sleep=sleeps.append)
    message = str(info.value)
    assert message.startswith("Écriture interrompue dans ")
    assert message.endswith(
        ": inbox_digest.md déjà mis à jour ; chaque fichier porte son propre generated_at."
    )
    assert sleeps == [0.5, 0.5]
    assert sorted(p.name for p in out.iterdir()) == [INBOX_FILE]


def test_other_os_errors_are_not_retried(tmp_path, monkeypatch):
    def broken(src, dst):
        raise OSError("MARQUEUR")

    monkeypatch.setattr(export.os, "replace", broken)
    sleeps: list[float] = []
    with pytest.raises(ExportError, match="aucun fichier mis à jour") as info:
        write_files(tmp_path, {"a.md": "x\n", "b.md": "y\n"}, sleep=sleeps.append)
    assert "MARQUEUR" not in str(info.value)
    assert sleeps == []
    assert list(tmp_path.iterdir()) == []


def test_write_files_reports_byte_sizes_in_commit_order(tmp_path):
    files = {"b.md": "Réunion test\n", "a.json": "{}\n"}
    assert write_files(tmp_path, files) == [("b.md", 14), ("a.json", 3)]
    assert (tmp_path / "b.md").read_bytes() == "Réunion test\n".encode()
    assert leftovers(tmp_path) == []


# ---------------------------------------------------------------------------- audit


def test_export_audit_records(audited_bridge, sink, out):
    export_snapshots(audited_bridge, out)
    assert [r["tool"] for r in sink.records] == [
        "search_mail",
        "search_mail",
        "list_events",
        "find_free_slots",
        "export",
    ]
    assert {r["run"] for r in sink.records} == {"run00001"}
    assert sink.records[0]["params"] == {"folder": "Inbox", "unread_only": True, "limit": 100}
    assert sink.records[1]["params"]["since"] == "2026-10-09T00:00+02:00"
    assert sink.records[3]["params"]["start"] == "2026-10-12T12:00+02:00"
    assert sink.records[3]["params"]["min_duration_min"] == 60
    last = sink.records[-1]
    assert last["params"] == {"custom_out": True}
    assert last["count"] == 3
    assert last["outcome"] == "ok"
    assert "snap" not in json.dumps(sink.records)  # the path is never logged


def test_config_failure_is_audited_before_any_outlook_call(audited_bridge, sink, gateway):
    with pytest.raises(ConfigError):
        export_snapshots(audited_bridge)
    assert gateway.calls == []
    assert [(r["tool"], r["params"], r["count"], r["outcome"]) for r in sink.records] == [
        ("export", {"custom_out": False}, None, "ConfigError")
    ]


def test_last_export(tmp_path, gateway, config):
    configured = config.model_copy(update={"snapshot_dir": str(tmp_path / "snap")})
    assert last_export(configured) is None
    export_snapshots(Bridge(gateway, configured, clock=lambda: NOW))
    stamp = last_export(configured)
    assert stamp is not None and stamp.utcoffset() is not None
    assert (tmp_path / "snap" / CALENDAR_JSON).exists()
    assert last_export(Config()) is None
    assert last_export(Config(snapshot_dir="relatif")) is None


def test_source_literals_fit_console_code_pages():
    for module in (export, models):
        source = Path(module.__file__).read_text(encoding="utf-8")
        source.encode("cp850")
        source.encode("cp1252")
