import sys
from winscript.core.window_resolver import find_window
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError
from winscript.core.selector import find_element
from winscript.core.state_diff import StateCapture

try:
    import pywinauto
    from pywinauto.keyboard import send_keys as pw_send_keys
except ImportError:
    pywinauto = None
    pw_send_keys = None

KEY_MAP = {
    "enter": "{ENTER}", "tab": "{TAB}", "escape": "{ESC}", "esc": "{ESC}",
    "backspace": "{BACKSPACE}", "delete": "{DELETE}",
    "home": "{HOME}", "end": "{END}",
    "up": "{UP}", "down": "{DOWN}", "left": "{LEFT}", "right": "{RIGHT}",
    "ctrl+c": "^c", "ctrl+v": "^v", "ctrl+x": "^x",
    "ctrl+z": "^z", "ctrl+s": "^s", "ctrl+a": "^a", "ctrl+w": "^w",
    "alt+f4": "%{F4}",
    "f1": "{F1}", "f2": "{F2}", "f3": "{F3}", "f4": "{F4}",
    "f5": "{F5}", "f6": "{F6}",
}

def click(app_title: str, element_name: str) -> str:
    """Click a UI element by name in a window.
    Example: click("Notepad", "File"), click("Chrome", "Address bar")"""
    args = {"app_title": app_title, "element_name": element_name}
    with StateCapture("click", args) as capture:
        try:
            window = find_window(app_title)
            if window is None:
                guard.record_failure("click", args)
                result = f"ERROR: No window '{app_title}'"
            else:
                element_result = find_element(window, element_name)
                if element_result is None:
                    guard.record_failure("click", args)
                    result = f"ERROR: No element '{element_name}' in '{app_title}'"
                else:
                    element_result.element.click_input()
                    guard.record_success("click", args)
                    result = f"Clicked '{element_name}' in '{app_title}' [via {element_result.layer_used}, confidence {element_result.confidence:.2f}]"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("click", args)
            result = f"ERROR clicking: {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"

def type_text(app_title: str, text: str) -> str:
    """Type text into the focused element of a window.
    Example: type_text("Notepad", "Hello world")"""
    args = {"app_title": app_title, "text": text[:50]}
    with StateCapture("type_text", args) as capture:
        try:
            window = find_window(app_title)
            if window is None:
                guard.record_failure("type_text", args)
                result = f"ERROR: No window '{app_title}'"
            else:
                window.set_focus()
                window.type_keys(text, with_spaces=True)
                guard.record_success("type_text", args)
                result = f"Typed {len(text)} chars into '{app_title}'"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("type_text", args)
            result = f"ERROR typing: {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"

def read_text(app_title: str, element_name: str = "") -> str:
    """Read text from a UI element. If element_name empty, reads window title.
    Example: read_text("Chrome", "Address bar")"""
    args = {"app_title": app_title, "element_name": element_name}
    try:
        window = find_window(app_title)
        if window is None:
            guard.record_failure("read_text", args)
            return f"ERROR: No window '{app_title}'"
        if not element_name:
            result_text = window.window_text()
            layer_info = ""
        else:
            result = find_element(window, element_name)
            if result is None:
                guard.record_failure("read_text", args)
                return f"ERROR: No element '{element_name}'"
            result_text = result.element.window_text()
            layer_info = f" [via {result.layer_used}, confidence {result.confidence:.2f}]"
        guard.record_success("read_text", args)
        return str(result_text) + layer_info
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("read_text", args)
        return f"ERROR reading: {str(e)}"

def press_key(key: str, app_title: str = "") -> str:
    """Press a key or combo. Focuses app first if given.
    Supported: enter, tab, ctrl+c, ctrl+s, alt+f4, f1-f6, etc.
    Example: press_key("ctrl+s", "Notepad")"""
    args = {"key": key, "app_title": app_title}
    with StateCapture("press_key", args) as capture:
        try:
            if app_title:
                window = find_window(app_title)
                if window:
                    window.set_focus()
            mapped = KEY_MAP.get(key.lower(), key)
            if pw_send_keys is None:
                 raise ImportError("pywinauto is not available")
            pw_send_keys(mapped)
            guard.record_success("press_key", args)
            result = f"Pressed '{key}'"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("press_key", args)
            result = f"ERROR pressing '{key}': {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"

def get_ui_tree(app_title: str, depth: int = 3) -> str:
    """Get accessible UI element tree of a window.
    Use this to discover what elements exist before clicking.
    depth: 1-5, default 3. Example: get_ui_tree("Notepad")"""
    args = {"app_title": app_title, "depth": depth}
    try:
        window = find_window(app_title)
        if window is None:
            guard.record_failure("get_ui_tree", args)
            return f"ERROR: No window '{app_title}'"
        lines = []
        def walk(elem, level):
            if level > depth:
                return
            try:
                title = elem.window_text() or ""
                # Handle mocked objects or actual pywinauto objects safely
                ctrl_type = ""
                if hasattr(elem, 'element_info') and elem.element_info:
                    ctrl_type = getattr(elem.element_info, 'control_type', "")
                elif hasattr(elem, 'control_type'):
                     ctrl_type = elem.control_type
                lines.append(f"{'  ' * level}[{ctrl_type}] {title!r}")
                if hasattr(elem, 'children'):
                    for child in elem.children():
                        walk(child, level + 1)
            except Exception:
                pass
        walk(window, 0)
        guard.record_success("get_ui_tree", args)
        return "\n".join(lines) if lines else "(empty tree)"
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("get_ui_tree", args)
        return f"ERROR getting UI tree: {str(e)}"

def coordinate_click(x: int, y: int) -> str:
    """Click exactly at (x, y) coordinates. Useful for Electron/UWP apps with poor accessibility trees.
    Example: coordinate_click(500, 300)"""
    args = {"x": x, "y": y}
    with StateCapture("coordinate_click", args) as capture:
        try:
            if pywinauto is None:
                raise ImportError("pywinauto is not available")
            pywinauto.mouse.click(button='left', coords=(x, y))
            guard.record_success("coordinate_click", args)
            result = f"Clicked at ({x}, {y})"
        except WinScriptMaxRetriesError:
            raise
        except Exception as e:
            guard.record_failure("coordinate_click", args)
            result = f"ERROR clicking at ({x}, {y}): {str(e)}"
    delta_summary = capture.delta.to_summary() if capture.delta else ""
    return f"{result} | {delta_summary}"
