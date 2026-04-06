try:
    import win32com.client
except ImportError:
    win32com = None
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError

def excel_read_cell(filepath: str, sheet: str, cell: str) -> str:
    """Read a cell from Excel. Example: excel_read_cell("C:/data.xlsx","Sheet1","B3")"""
    args = {"filepath": filepath, "sheet": sheet, "cell": cell}
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(filepath)
        value = wb.Sheets(sheet).Range(cell).Value
        wb.Close(SaveChanges=False)
        guard.record_success("excel_read_cell", args)
        return str(value) if value is not None else "(empty)"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("excel_read_cell", args)
        return f"ERROR: {str(e)}"
    finally:
        try:
            if excel: excel.Quit()
        except Exception:
            pass

def excel_write_cell(filepath: str, sheet: str, cell: str, value: str) -> str:
    """Write a value to an Excel cell and save.
    Example: excel_write_cell("C:/data.xlsx","Sheet1","B3","Done")"""
    args = {"filepath": filepath, "sheet": sheet, "cell": cell, "value": value}
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(filepath)
        wb.Sheets(sheet).Range(cell).Value = value
        wb.Save()
        wb.Close(SaveChanges=True)
        guard.record_success("excel_write_cell", args)
        return f"Written '{value}' to {sheet}!{cell}"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("excel_write_cell", args)
        return f"ERROR: {str(e)}"
    finally:
        try:
            if excel: excel.Quit()
        except Exception:
            pass

def excel_read_range(filepath: str, sheet: str, start_cell: str, end_cell: str) -> str:
    """Read a range from Excel as CSV-style string.
    Example: excel_read_range("C:/data.xlsx","Sheet1","A1","C5")"""
    args = {"filepath": filepath, "sheet": sheet, "start": start_cell, "end": end_cell}
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(filepath)
        rows = wb.Sheets(sheet).Range(f"{start_cell}:{end_cell}").Value
        wb.Close(SaveChanges=False)
        if rows is None:
            return "(empty range)"
        lines = [", ".join(str(v) if v is not None else "" for v in row) for row in rows]
        guard.record_success("excel_read_range", args)
        return "\n".join(lines)
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("excel_read_range", args)
        return f"ERROR: {str(e)}"
    finally:
        try:
            if excel: excel.Quit()
        except Exception:
            pass

def outlook_send_email(to: str, subject: str, body: str) -> str:
    """Send an email via Outlook. Requires Outlook installed + account configured.
    Example: outlook_send_email("team@co.com","Report","See attached.")"""
    args = {"to": to, "subject": subject}
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        mail.Send()
        guard.record_success("outlook_send_email", args)
        return f"Email sent to {to} | Subject: {subject}"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("outlook_send_email", args)
        return f"ERROR: {str(e)}"

def outlook_read_inbox(count: int = 10) -> str:
    """Read N most recent emails from Outlook inbox.
    Example: outlook_read_inbox(5)"""
    args = {"count": count}
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)
        results = []
        for i, msg in enumerate(messages):
            if i >= count: break
            try:
                results.append(f"[{msg.ReceivedTime}] {msg.SenderName} | {msg.Subject}")
            except Exception:
                continue
        guard.record_success("outlook_read_inbox", args)
        return "\n".join(results) if results else "(no emails)"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("outlook_read_inbox", args)
        return f"ERROR: {str(e)}"