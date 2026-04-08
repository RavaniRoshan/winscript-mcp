from winscript.tools.com_office import outlook_send_email, outlook_read_inbox
from winscript.tools.app_control import open_app

class OutlookAdapter:
    """Semantic Outlook automation."""

    def send(self, to: str, subject: str, body: str) -> str:
        """Send an email. Example: outlook.send("a@b.com", "Hi", "Body here")"""
        return outlook_send_email(to, subject, body)

    def inbox(self, count: int = 10) -> str:
        """Read recent emails. Example: outlook.inbox(5)"""
        return outlook_read_inbox(count)

    def open(self) -> str:
        return open_app("outlook")

outlook = OutlookAdapter()