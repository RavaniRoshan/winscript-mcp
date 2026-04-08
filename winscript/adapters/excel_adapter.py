from winscript.tools.com_office import excel_read_cell, excel_write_cell, excel_read_range
from winscript.tools.ui_interaction import click, press_key, type_text
from winscript.tools.app_control import open_app, focus_app
from winscript.tools.filesystem import file_exists

class ExcelAdapter:
    """Semantic Excel automation. Knows Excel's structure — not just its pixels."""

    def open(self, filepath: str) -> str:
        """Open an Excel file. Example: excel.open("C:/reports/q1.xlsx")"""
        if file_exists(filepath).startswith("NOT FOUND"):
            return f"ERROR: File not found: {filepath}"
        result = open_app("excel")
        import time; time.sleep(2)
        # Open file via keyboard
        press_key("ctrl+o", "Microsoft Excel")
        import time; time.sleep(0.5)
        type_text("Microsoft Excel", filepath)
        press_key("enter")
        import time; time.sleep(1.5)
        return f"Opened {filepath} in Excel"

    def read_cell(self, filepath: str, sheet: str, cell: str) -> str:
        return excel_read_cell(filepath, sheet, cell)

    def write_cell(self, filepath: str, sheet: str, cell: str, value: str) -> str:
        return excel_write_cell(filepath, sheet, cell, value)

    def read_range(self, filepath: str, sheet: str, start: str, end: str) -> str:
        return excel_read_range(filepath, sheet, start, end)

    def save(self, app_title: str = "Microsoft Excel") -> str:
        return press_key("ctrl+s", app_title)

    def close(self, save: bool = True) -> str:
        key = "ctrl+s" if save else ""
        if key:
            press_key(key, "Microsoft Excel")
            import time; time.sleep(0.3)
        return press_key("alt+f4", "Microsoft Excel")

excel = ExcelAdapter()