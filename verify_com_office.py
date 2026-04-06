import sys
import unittest
from unittest.mock import MagicMock, patch

# Mock win32com before import
mock_win32com = MagicMock()
sys.modules['win32com'] = mock_win32com
sys.modules['win32com.client'] = mock_win32com.client

import winscript.tools.com_office as co
from winscript.core.retry_guard import guard

class MockExcelRange:
    def __init__(self, value):
        self.Value = value

class MockExcelSheet:
    def __init__(self, data):
        self._data = data
        
    def Range(self, rng):
        return MockExcelRange(self._data.get(rng, None))

class MockExcelWorkbook:
    def __init__(self, sheet_data):
        self._sheets = {"Sheet1": MockExcelSheet(sheet_data)}
        self.saved = False
        self.closed = False
        
    def Sheets(self, name):
        return self._sheets[name]
        
    def Save(self):
        self.saved = True
        
    def Close(self, SaveChanges=False):
        self.closed = True

class MockExcelApp:
    def __init__(self, sheet_data):
        self.Workbooks = MagicMock()
        self.Workbooks.Open.return_value = MockExcelWorkbook(sheet_data)
        self.Quit_called = False
        
    def Quit(self):
        self.Quit_called = True

class MockMailItem:
    def __init__(self):
        self.To = ""
        self.Subject = ""
        self.Body = ""
        self.sent = False
        
    def Send(self):
        self.sent = True

class MockOutlookItem:
    def __init__(self, sender, subject, time):
        self.SenderName = sender
        self.Subject = subject
        self.ReceivedTime = time

class MockOutlookInbox:
    def __init__(self):
        self.Items = MagicMock()
        self.Items.__iter__.return_value = [
            MockOutlookItem("Alice", "Hello", "2023-10-01"),
            MockOutlookItem("Bob", "World", "2023-10-02")
        ]

class MockOutlookApp:
    def __init__(self):
        pass
        
    def CreateItem(self, type):
        return MockMailItem()
        
    def GetNamespace(self, ns):
        ns_mock = MagicMock()
        ns_mock.GetDefaultFolder.return_value = MockOutlookInbox()
        return ns_mock

class TestComOffice(unittest.TestCase):
    def setUp(self):
        guard.reset_all()

    @patch('winscript.tools.com_office.win32com.client.Dispatch')
    def test_excel_read_cell(self, mock_dispatch):
        mock_excel = MockExcelApp({"A1": "Hello"})
        mock_dispatch.return_value = mock_excel
        
        res = co.excel_read_cell("test.xlsx", "Sheet1", "A1")
        self.assertEqual(res, "Hello")
        self.assertTrue(mock_excel.Quit_called)

    @patch('winscript.tools.com_office.win32com.client.Dispatch')
    def test_excel_write_cell(self, mock_dispatch):
        mock_excel = MockExcelApp({})
        mock_dispatch.return_value = mock_excel
        
        res = co.excel_write_cell("test.xlsx", "Sheet1", "A1", "World")
        self.assertEqual(res, "Written 'World' to Sheet1!A1")
        self.assertTrue(mock_excel.Quit_called)
        
    @patch('winscript.tools.com_office.win32com.client.Dispatch')
    def test_excel_read_range(self, mock_dispatch):
        mock_excel = MockExcelApp({"A1:B2": [["A1", "B1"], ["A2", "B2"]]})
        mock_dispatch.return_value = mock_excel
        
        res = co.excel_read_range("test.xlsx", "Sheet1", "A1", "B2")
        self.assertEqual(res, "A1, B1\nA2, B2")
        self.assertTrue(mock_excel.Quit_called)

    @patch('winscript.tools.com_office.win32com.client.Dispatch')
    def test_outlook_send_email(self, mock_dispatch):
        mock_outlook = MockOutlookApp()
        mock_dispatch.return_value = mock_outlook
        
        res = co.outlook_send_email("test@test.com", "Subj", "Body")
        self.assertEqual(res, "Email sent to test@test.com | Subject: Subj")

    @patch('winscript.tools.com_office.win32com.client.Dispatch')
    def test_outlook_read_inbox(self, mock_dispatch):
        mock_outlook = MockOutlookApp()
        mock_dispatch.return_value = mock_outlook
        
        res = co.outlook_read_inbox(2)
        self.assertIn("Alice | Hello", res)
        self.assertIn("Bob | World", res)

if __name__ == '__main__':
    unittest.main()