from fastmcp import FastMCP

from winscript.tools.app_control import open_app, close_app, focus_app, get_running_apps
from winscript.tools.ui_interaction import click, type_text, read_text, press_key, get_ui_tree
from winscript.tools.com_office import (
    excel_read_cell, excel_write_cell, excel_read_range,
    outlook_send_email, outlook_read_inbox
)
from winscript.tools.filesystem import (
    read_file_text, write_file_text, list_dir,
    move_file, copy_file, delete_file, file_exists
)
from winscript.tools.screen import take_screenshot, get_active_window, get_clipboard, set_clipboard

mcp = FastMCP(
    name="winscript",
    version="0.1.0"
)

# App Control
mcp.tool()(open_app)
mcp.tool()(close_app)
mcp.tool()(focus_app)
mcp.tool()(get_running_apps)

# UI Interaction
mcp.tool()(click)
mcp.tool()(type_text)
mcp.tool()(read_text)
mcp.tool()(press_key)
mcp.tool()(get_ui_tree)

# COM Office
mcp.tool()(excel_read_cell)
mcp.tool()(excel_write_cell)
mcp.tool()(excel_read_range)
mcp.tool()(outlook_send_email)
mcp.tool()(outlook_read_inbox)

# File System
mcp.tool()(read_file_text)
mcp.tool()(write_file_text)
mcp.tool()(list_dir)
mcp.tool()(move_file)
mcp.tool()(copy_file)
mcp.tool()(delete_file)
mcp.tool()(file_exists)

# Screen
mcp.tool()(take_screenshot)
mcp.tool()(get_active_window)
mcp.tool()(get_clipboard)
mcp.tool()(set_clipboard)

if __name__ == "__main__":
    mcp.run()