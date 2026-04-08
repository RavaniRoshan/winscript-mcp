from winscript.tools.app_control import open_app
from winscript.tools.ui_interaction import click, type_text, press_key, read_text
from winscript.tools.screen import take_screenshot
import time

class ChromeAdapter:
    """Semantic Chrome automation. Navigate, read, interact."""

    def open(self, url: str = "") -> str:
        result = open_app("chrome")
        if url:
            time.sleep(2)
            return self.navigate(url)
        return result

    def navigate(self, url: str) -> str:
        """Navigate to a URL. Example: chrome.navigate("https://github.com")"""
        press_key("ctrl+l", "Google Chrome")
        time.sleep(0.3)
        type_text("Google Chrome", url)
        press_key("enter", "Google Chrome")
        time.sleep(1.5)
        return f"Navigated to {url}"

    def get_url(self) -> str:
        """Read current URL from address bar."""
        press_key("ctrl+l", "Google Chrome")
        import time; time.sleep(0.2)
        press_key("ctrl+c", "Google Chrome")
        import time; time.sleep(0.2)
        from winscript.tools.screen import get_clipboard
        url = get_clipboard()
        press_key("escape", "Google Chrome")
        return url

    def get_title(self) -> str:
        return read_text("Google Chrome", "")

    def new_tab(self) -> str:
        press_key("ctrl+t", "Google Chrome")
        return "New tab opened"

    def close_tab(self) -> str:
        press_key("ctrl+w", "Google Chrome")
        return "Tab closed"

    def screenshot(self) -> str:
        """Screenshot of current Chrome window."""
        return take_screenshot()

    def find_on_page(self, text: str) -> str:
        press_key("ctrl+f", "Google Chrome")
        time.sleep(0.3)
        type_text("Google Chrome", text)
        return f"Searching for '{text}' on page"

chrome = ChromeAdapter()