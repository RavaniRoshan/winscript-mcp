import sys
from unittest.mock import MagicMock

sys.modules['pywinauto'] = MagicMock()
sys.modules['pywinauto.keyboard'] = MagicMock()
sys.modules['win32gui'] = MagicMock()
sys.modules['win32clipboard'] = MagicMock()
sys.modules['win32con'] = MagicMock()
sys.modules['win32com'] = MagicMock()
sys.modules['win32com.client'] = MagicMock()

import winscript.tools.app_control as ac
ac.open_app = MagicMock(return_value="Opened notepad")

import winscript.tools.ui_interaction as ui
ui.type_text = MagicMock(return_value="Typed 10 chars")
ui.press_key = MagicMock()

import winscript.tools.filesystem as fs
fs.write_file_text = MagicMock(return_value="Written 10 chars")
fs.read_file_text = MagicMock(return_value="WinScript works.")

import winscript.tools.screen as sc
sc.take_screenshot = MagicMock(return_value="data:image/png;base64,1234")
sc.set_clipboard = MagicMock(return_value="Clipboard set (10 chars)")
sc.get_clipboard = MagicMock(return_value="winscript clipboard test")

from tests.test_e2e import test_full_flow
test_full_flow()