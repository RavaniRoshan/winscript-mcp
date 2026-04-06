import sys
import unittest
from unittest.mock import MagicMock, patch

# Need to mock pywinauto before things get imported if we mock on top
sys.modules['pywinauto'] = MagicMock()

import winscript.tools.app_control as ac
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError

class MockWindow:
    def __init__(self, title):
        self._title = title
        self.close_called = False
        self.focus_called = False

    def window_text(self):
        return self._title

    def close(self):
        self.close_called = True

    def set_focus(self):
        self.focus_called = True

class TestAppControl(unittest.TestCase):
    def setUp(self):
        guard.reset_all()

    @patch('winscript.tools.app_control.subprocess.Popen')
    @patch('winscript.tools.app_control.time.sleep')
    def test_open_app_success(self, mock_sleep, mock_popen):
        res = ac.open_app("notepad", wait_seconds=0)
        self.assertEqual(res, "Opened notepad (notepad.exe)")
        mock_popen.assert_called_once_with("notepad.exe")

    @patch('winscript.tools.app_control.get_all_windows')
    def test_get_running_apps(self, mock_get_all):
        mock_get_all.return_value = [
            {"title": "Untitled - Notepad", "pid": 1234}
        ]
        res = ac.get_running_apps()
        self.assertIn("Untitled - Notepad", res)
        self.assertIn("1234", res)

    @patch('winscript.tools.app_control.find_window')
    def test_focus_app_success(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ac.focus_app("Notepad")
        self.assertEqual(res, "Focused: Untitled - Notepad")
        self.assertTrue(mock_w.focus_called)

    @patch('winscript.tools.app_control.find_window')
    def test_close_app_success(self, mock_find):
        mock_w = MockWindow("Untitled - Notepad")
        mock_find.return_value = mock_w
        
        res = ac.close_app("Notepad")
        self.assertEqual(res, "Closed window: Untitled - Notepad")
        self.assertTrue(mock_w.close_called)

    @patch('winscript.tools.app_control.find_window')
    def test_close_app_not_found(self, mock_find):
        mock_find.return_value = None
        res = ac.close_app("xyznonexistent")
        self.assertTrue(res.startswith("ERROR: No window found matching"))

    @patch('winscript.tools.app_control.subprocess.Popen')
    def test_open_app_max_retries(self, mock_popen):
        mock_popen.side_effect = Exception("Simulated launch failure")
        
        for i in range(4):
            res = ac.open_app("badapp", wait_seconds=0)
            self.assertTrue(res.startswith("ERROR opening badapp: Simulated launch failure"))
            
        with self.assertRaises(WinScriptMaxRetriesError):
            ac.open_app("badapp", wait_seconds=0)

if __name__ == '__main__':
    unittest.main()