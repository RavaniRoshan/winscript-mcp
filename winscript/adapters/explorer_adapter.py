from winscript.tools.app_control import open_app
from winscript.tools.ui_interaction import type_text, press_key
from winscript.tools.filesystem import list_dir
import time

class ExplorerAdapter:
    """Semantic File Explorer automation."""

    def open(self, path: str = "") -> str:
        result = open_app("explorer")
        if path:
            time.sleep(1.5)
            return self.navigate(path)
        return result

    def navigate(self, path: str) -> str:
        """Navigate Explorer to a path. Example: explorer.navigate("C:/Users")"""
        press_key("ctrl+l", "File Explorer")
        time.sleep(0.2)
        type_text("File Explorer", path)
        press_key("enter", "File Explorer")
        time.sleep(0.8)
        return f"Navigated to {path}"

    def list(self, path: str) -> str:
        """List files in path (uses filesystem, not UI). Faster."""
        return list_dir(path)

explorer = ExplorerAdapter()