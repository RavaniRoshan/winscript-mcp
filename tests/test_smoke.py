from winscript import mcp

def test_server_exists():
    assert mcp is not None
    assert mcp.name == "winscript"