"""Bridge: the single core shared by the CLI, the MCP server and the snapshot export.

The methods that map 1:1 to the MCP tools (here and on `mail` / `calendar`) are `@audited`:
every front end goes through them, so none can skip the audit log. A hand-built
`Bridge(gateway, config)` logs nowhere; `from_environment` wires the real audit file.
"""

from __future__ import annotations

from datetime import datetime
from typing import Literal

from outlook_bridge.audit import Auditor, Clock, JsonlSink, NullSink, audited
from outlook_bridge.config import Config, load_config, log_dir
from outlook_bridge.events import CalendarService
from outlook_bridge.gateway import OutlookGateway
from outlook_bridge.mail import MailService
from outlook_bridge.models import Account, Folder, Status
from outlook_bridge.redact import Policy

SOURCES = ("cli", "mcp")


class Bridge:
    def __init__(
        self,
        gateway: OutlookGateway,
        config: Config,
        *,
        audit: Auditor | None = None,
        clock: Clock | None = None,
    ) -> None:
        self.clock: Clock = clock or (lambda: datetime.now(config.tz))
        # Set first: the audited methods (here and in the services) read it.
        self._audit = (
            audit if audit is not None else Auditor(NullSink(), source="core", clock=self.clock)
        )
        self.config = config
        self.policy = Policy(config)
        self.gateway = gateway
        self.mail = MailService(gateway, config, self.policy, audit=self._audit)
        self.calendar = CalendarService(gateway, config, self.policy, audit=self._audit)

    @classmethod
    def from_environment(cls, *, source: Literal["cli", "mcp"]) -> Bridge:
        """Real Outlook through COM, the user's config.toml and the audit log.

        `source` is required so that no front end can forget to label its records. The COM
        worker thread starts lazily, on the first Outlook call.
        """
        if source not in SOURCES:
            raise ValueError(f"source inconnue : {source!r}")
        from outlook_bridge.com_gateway import ComGateway  # Windows-only import

        config = load_config()

        def clock() -> datetime:
            return datetime.now(config.tz)

        auditor = Auditor(JsonlSink(log_dir()), source=source, clock=clock)
        return cls(ComGateway(tz=config.tz), config, audit=auditor, clock=clock)

    @property
    def audit(self) -> Auditor:
        return self._audit

    def now(self) -> datetime:
        """Current time, aware in config.tz (injected clock in tests)."""
        return self.clock()

    def close(self) -> None:
        self.gateway.close()

    def __enter__(self) -> Bridge:
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    @audited
    def status(self) -> Status:
        status = self.gateway.status()
        return status.model_copy(update={"accounts": self._allowed(status.accounts)})

    @audited
    def list_accounts(self) -> list[Account]:
        return self._allowed(self.gateway.list_accounts())

    @audited
    def list_folders(self, account: str | None = None) -> list[Folder]:
        return [
            f
            for f in self.gateway.list_folders(account)
            if self.policy.account_allowed(f.account) and not self.policy.folder_excluded(f.path)
        ]

    def _allowed(self, accounts: list[Account]) -> list[Account]:
        return [a for a in accounts if self.policy.account_allowed(a.display_name, a.smtp_address)]
