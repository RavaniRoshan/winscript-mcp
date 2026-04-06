import sys
import unittest
from unittest.mock import MagicMock, patch

# We must mock pywinauto before importing window_resolver
sys.modules['pywinauto'] = MagicMock()
import winscript.core.window_resolver as wr

class MockWindow:
    def __init__(self, title, pid):
        self._title = title
        self._pid = pid
        
    def window_text(self):
        return self._title
        
    def process_id(self):
        return self._pid

class TestWindowResolver(unittest.TestCase):
    def setUp(self):
        self.mock_desktop = MagicMock()
        self.mock_desktop.windows.return_value = [
            MockWindow("Untitled - Notepad", 1234),
            MockWindow("Calculator", 5678),
            MockWindow("Command Prompt - python", 9012)
        ]
        
        # Patch the Desktop class used in window_resolver
        self.patcher = patch('winscript.core.window_resolver.Desktop', return_value=self.mock_desktop)
        self.patcher.start()
        
    def tearDown(self):
        self.patcher.stop()

    def test_find_window_exact_match(self):
        # 1. Exact match
        w = wr.find_window("Calculator")
        self.assertIsNotNone(w)
        self.assertEqual(w.window_text(), "Calculator")

    def test_find_window_case_insensitive_contains(self):
        # 2. Case-insensitive contains
        w = wr.find_window("notepad")
        self.assertIsNotNone(w)
        self.assertEqual(w.window_text(), "Untitled - Notepad")
        
    def test_find_window_regex_match(self):
        # 3. Regex match (e.g. matching prompt that starts with 'Command')
        w = wr.find_window("^Command.*")
        self.assertIsNotNone(w)
        self.assertEqual(w.window_text(), "Command Prompt - python")

    def test_find_window_not_found(self):
        w = wr.find_window("xyznonexistent")
        self.assertIsNone(w)

    def test_get_all_windows(self):
        windows = wr.get_all_windows()
        self.assertEqual(len(windows), 3)
        titles = [w['title'] for w in windows]
        self.assertIn("Untitled - Notepad", titles)

if __name__ == '__main__':
    unittest.main()