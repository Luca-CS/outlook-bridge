"""Bridge: the single core shared by the CLI, the MCP server and the snapshot export."""

from __future__ import annotations

from outlook_bridge.config import Config, load_config
from outlook_bridge.events import CalendarService
from outlook_bridge.gateway import OutlookGateway
from outlook_bridge.mail import MailService
from outlook_bridge.models import Account, Folder, Status
from outlook_bridge.redact import Policy


class Bridge:
    def __init__(self, gateway: OutlookGateway, config: Config) -> None:
        self.config = config
        self.policy = Policy(config)
        self.gateway = gateway
        self.mail = MailService(gateway, config, self.policy)
        self.calendar = CalendarService(gateway, config, self.policy)

    @classmethod
    def from_environment(cls) -> Bridge:
        """Real Outlook through COM, with the user's config.toml."""
        from outlook_bridge.com_gateway import ComGateway  # Windows-only import

        config = load_config()
        return cls(ComGateway(tz=config.tz), config)

    def status(self) -> Status:
        status = self.gateway.status()
        return status.model_copy(update={"accounts": self._allowed(status.accounts)})

    def list_accounts(self) -> list[Account]:
        return self._allowed(self.gateway.list_accounts())

    def list_folders(self, account: str | None = None) -> list[Folder]:
        return [
            f
            for f in self.gateway.list_folders(account)
            if self.policy.account_allowed(f.account) and not self.policy.folder_excluded(f.path)
        ]

    def _allowed(self, accounts: list[Account]) -> list[Account]:
        return [a for a in accounts if self.policy.account_allowed(a.display_name, a.smtp_address)]
