from winscript.tools.app_control import open_app, close_app
from winscript.tools.ui_interaction import type_text, press_key
from winscript.tools.filesystem import read_file_text
import time

class NotepadAdapter:
    """Semantic Notepad automation."""

    def open(self, filepath: str = "") -> str:
        if filepath:
            import subprocess
            subprocess.Popen(["notepad.exe", filepath])
            time.sleep(1.5)
            return f"Opened {filepath} in Notepad"
        return open_app("notepad")

    def type(self, text: str) -> str:
        return type_text("Notepad", text)

    def save_as(self, filepath: str) -> str:
        press_key("ctrl+shift+s", "Notepad")
        time.sleep(0.5)
        type_text("Notepad", filepath)
        press_key("enter", "Notepad")
        return f"Saved as {filepath}"

    def save(self) -> str:
        return press_key("ctrl+s", "Notepad")

    def close(self, save: bool = False) -> str:
        if save:
            press_key("ctrl+s", "Notepad")
            time.sleep(0.3)
        press_key("alt+f4", "Notepad")
        time.sleep(0.3)
        press_key("tab")
        press_key("enter")
        return "Notepad closed"

    def read(self, filepath: str) -> str:
        return read_file_text(filepath)

notepad = NotepadAdapter()