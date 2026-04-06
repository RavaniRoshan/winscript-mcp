import sys
import unittest
from unittest.mock import MagicMock, patch

sys.modules['pywinauto'] = MagicMock()
sys.modules['pywinauto.keyboard'] = MagicMock()

import winscript.tools.ui_interaction as ui
from winscript.core.retry_guard import guard

class MockElement:
    def __init__(self, title, control_type=""):
        self._title = title
        self.control_type = control_type
        self.clicked = False
        
    def window_text(self):
        return self._title
        
    def click_input(self):
        self.clicked = True
        
    def wait(self, state, timeout):
        pass

    def children(self):
        return []

class MockWindow:
    def __init__(self, title):
        self._title = title
        self.focus_called = False
        self.typed_keys = None
        self.control_type = "WindowControl"
        
        self.child_file = MockElement("File", "MenuItemControl")
        self.child_edit = MockElement("Text Editor", "EditControl")
        
    def window_text(self):
        return self._title
        
    def set_focus(self):
        self.focus_called = True
        
    def type_keys(self, keys, with_spaces=True):
        self.typed_keys = keys
        
    def child_window(self, title=None, auto_id=None, title_re=None, found_index=0):
        if title == "File": return self.child_file
        if title == "Text Editor": return self.child_edit
        raise Exception("Element not found")

    def children(self):
        return [self.child_file, self.child_edit]

class TestUIInteraction(unittest.TestCase):
    def setUp(self):
        guard.reset_all()

    @patch('winscript.tools.ui_interaction.find_window')
    def test_click_success(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ui.click("Notepad", "File")
        self.assertEqual(res, "Clicked 'File' in 'Notepad'")
        self.assertTrue(mock_w.child_file.clicked)

    @patch('winscript.tools.ui_interaction.find_window')
    def test_click_missing_element(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ui.click("Notepad", "MissingElement")
        self.assertTrue(res.startswith("ERROR: No element 'MissingElement'"))

    @patch('winscript.tools.ui_interaction.find_window')
    def test_type_text_success(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ui.type_text("Notepad", "Hello World")
        self.assertEqual(res, "Typed 11 chars into 'Notepad'")
        self.assertEqual(mock_w.typed_keys, "Hello World")
        self.assertTrue(mock_w.focus_called)

    @patch('winscript.tools.ui_interaction.find_window')
    def test_read_text_window_title(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ui.read_text("Notepad")
        self.assertEqual(res, "Untitled - Notepad")

    @patch('winscript.tools.ui_interaction.find_window')
    def test_get_ui_tree(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ui.get_ui_tree("Notepad", depth=3)
        self.assertIn("[WindowControl] 'Untitled - Notepad'", res)
        self.assertIn("  [MenuItemControl] 'File'", res)
        self.assertIn("  [EditControl] 'Text Editor'", res)

    @patch('winscript.tools.ui_interaction.find_window')
    @patch('winscript.tools.ui_interaction.pw_send_keys')
    def test_press_key(self, mock_send_keys, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ui.press_key("ctrl+a", "Notepad")
        self.assertEqual(res, "Pressed 'ctrl+a'")
        mock_send_keys.assert_called_once_with("^a")
        self.assertTrue(mock_w.focus_called)

if __name__ == '__main__':
    unittest.main()