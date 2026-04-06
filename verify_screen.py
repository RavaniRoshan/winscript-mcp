import sys
import unittest
from unittest.mock import MagicMock

# Mock out Windows-specific libraries before importing
mock_mss = MagicMock()
sys.modules['mss'] = mock_mss

mock_win32gui = MagicMock()
sys.modules['win32gui'] = mock_win32gui

mock_win32clipboard = MagicMock()
sys.modules['win32clipboard'] = mock_win32clipboard

mock_win32con = MagicMock()
mock_win32con.CF_UNICODETEXT = 13
sys.modules['win32con'] = mock_win32con

import winscript.tools.screen as sc

class MockShot:
    def __init__(self, w, h):
        self.size = (w, h)
        # Mock RGB bytes representing a blank image for BGRX format parsing
        self.bgra = b'\x00' * (w * h * 4)

class MockSct:
    def __init__(self):
        self.monitors = [{}, {"top": 0, "left": 0, "width": 1920, "height": 1080}]
        
    def __enter__(self):
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        pass
        
    def grab(self, monitor):
        w = monitor["width"]
        h = monitor["height"]
        return MockShot(w, h)

mock_mss.mss.return_value = MockSct()

class TestScreen(unittest.TestCase):
    def test_take_screenshot_full(self):
        res = sc.take_screenshot()
        self.assertTrue(res.startswith("data:image/png;base64,"))
        
    def test_take_screenshot_region(self):
        res = sc.take_screenshot({"top":0, "left":0, "width":800, "height":600})
        self.assertTrue(res.startswith("data:image/png;base64,"))
        self.assertTrue(len(res) > 22) # Ensure we actually have data after the prefix

    def test_active_window(self):
        mock_win32gui.GetForegroundWindow.return_value = 1234
        mock_win32gui.GetWindowText.return_value = "Test Window"
        res = sc.get_active_window()
        self.assertEqual(res, "Test Window")
        
    def test_clipboard(self):
        # We need to mock clipboard state functionally
        clipboard_data = ""
        def mock_set(fmt, text):
            nonlocal clipboard_data
            clipboard_data = text
        def mock_get(fmt):
            return clipboard_data
            
        mock_win32clipboard.SetClipboardData.side_effect = mock_set
        mock_win32clipboard.GetClipboardData.side_effect = mock_get
        
        set_res = sc.set_clipboard("winscript test")
        self.assertTrue(set_res.startswith("Clipboard set"))
        
        get_res = sc.get_clipboard()
        self.assertEqual(get_res, "winscript test")

if __name__ == '__main__':
    unittest.main()