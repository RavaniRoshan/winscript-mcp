"""End-to-end test. Requires: Windows, Notepad, C:/tmp/ writable."""

def test_full_flow():
    from winscript.tools.app_control import open_app, close_app
    from winscript.tools.ui_interaction import type_text, press_key
    from winscript.tools.filesystem import write_file_text, read_file_text
    from winscript.tools.screen import take_screenshot, get_clipboard, set_clipboard
    import time

    # File round-trip
    assert "Written" in write_file_text("C:/tmp/winscript_e2e.txt", "WinScript works.")
    assert read_file_text("C:/tmp/winscript_e2e.txt") == "WinScript works."

    # Clipboard round-trip
    assert "Clipboard set" in set_clipboard("winscript clipboard test")
    assert get_clipboard() == "winscript clipboard test"

    # Screenshot
    shot = take_screenshot()
    assert shot.startswith("data:image/png;base64,")

    # Notepad open → type → close
    assert "Opened" in open_app("notepad")
    time.sleep(1.5)
    assert "Typed" in type_text("Notepad", "WinScript e2e test")
    press_key("alt+f4", "Notepad")
    time.sleep(0.5)
    press_key("tab")
    press_key("enter")

    print("All E2E tests passed.")
