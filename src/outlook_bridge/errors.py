"""Explicit, actionable errors shared by the CLI and the MCP server."""

from __future__ import annotations


class OutlookBridgeError(Exception):
    """Base error: the message is meant to be shown as is to the user or to Claude."""


class OutlookNotRunning(OutlookBridgeError):
    def __init__(self) -> None:
        super().__init__("Outlook n'est pas lancé : ouvrez l'Outlook classique puis réessayez.")


class NewOutlookDetected(OutlookBridgeError):
    def __init__(self) -> None:
        super().__init__(
            "Nouvel Outlook détecté : COM indisponible. Désactivez « Nouvel Outlook » "
            "pour revenir à l'Outlook classique."
        )


class OutlookUnavailable(OutlookBridgeError):
    """COM call failed (elevation mismatch, Outlook closing, RPC error...)."""


class NotFound(OutlookBridgeError):
    """Item missing *or* excluded by configuration (indistinguishable on purpose)."""

    def __init__(self, what: str = "Élément") -> None:
        super().__init__(f"{what} introuvable.")


class InvalidRequest(OutlookBridgeError):
    """Bad parameters (range too wide, end before start...)."""
