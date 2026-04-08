from dataclasses import dataclass, field
from typing import Optional
import time

@dataclass
class DesktopSnapshot:
    """Point-in-time snapshot of the desktop state."""
    timestamp: float
    active_window: str
    open_windows: list[str]
    clipboard: str
    focused_element: str = ""

    @classmethod
    def capture(cls) -> "DesktopSnapshot":
        try:
            import win32gui
            import win32clipboard
            import win32con
        except ImportError:
            win32gui = None
            win32clipboard = None
            win32con = None
        from winscript.core.window_resolver import get_all_windows

        active = ""
        try:
            if win32gui:
                hwnd = win32gui.GetForegroundWindow()
                active = win32gui.GetWindowText(hwnd)
        except Exception:
            pass

        windows = []
        try:
            windows = [w["title"] for w in get_all_windows()]
        except Exception:
            pass

        clipboard = ""
        try:
            if win32clipboard:
                win32clipboard.OpenClipboard()
                clipboard = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                win32clipboard.CloseClipboard()
        except Exception:
            pass

        return cls(
            timestamp=time.time(),
            active_window=active,
            open_windows=windows,
            clipboard=clipboard,
        )

@dataclass
class StateDelta:
    """What changed between two snapshots."""
    action_tool: str
    action_args: dict
    duration_ms: float
    before: DesktopSnapshot
    after: DesktopSnapshot

    # Computed deltas
    active_window_changed: bool = False
    new_windows: list[str] = field(default_factory=list)
    closed_windows: list[str] = field(default_factory=list)
    clipboard_changed: bool = False
    clipboard_new_value: str = ""

    def compute(self) -> "StateDelta":
        self.active_window_changed = (
            self.before.active_window != self.after.active_window
        )
        before_set = set(self.before.open_windows)
        after_set = set(self.after.open_windows)
        self.new_windows = list(after_set - before_set)
        self.closed_windows = list(before_set - after_set)
        self.clipboard_changed = self.before.clipboard != self.after.clipboard
        self.clipboard_new_value = self.after.clipboard if self.clipboard_changed else ""
        return self

    def to_summary(self) -> str:
        """Human-readable delta summary for agent consumption."""
        parts = []
        if self.active_window_changed:
            parts.append(
                f"Active window: '{self.before.active_window}' → '{self.after.active_window}'"
            )
        if self.new_windows:
            parts.append(f"Opened: {self.new_windows}")
        if self.closed_windows:
            parts.append(f"Closed: {self.closed_windows}")
        if self.clipboard_changed:
            preview = self.clipboard_new_value[:50]
            parts.append(f"Clipboard changed → '{preview}{'...' if len(self.clipboard_new_value) > 50 else ''}'")
        if not parts:
            parts.append("No detectable state change")
        parts.append(f"Duration: {self.duration_ms:.0f}ms")
        return " | ".join(parts)

    def to_dict(self) -> dict:
        return {
            "tool": self.action_tool,
            "duration_ms": round(self.duration_ms, 2),
            "active_window_changed": self.active_window_changed,
            "new_windows": self.new_windows,
            "closed_windows": self.closed_windows,
            "clipboard_changed": self.clipboard_changed,
            "clipboard_preview": self.clipboard_new_value[:100],
        }

# Context manager for wrapping any tool call with state diffing
class StateCapture:
    def __init__(self, tool_name: str, args: dict):
        self.tool_name = tool_name
        self.args = args
        self.before: Optional[DesktopSnapshot] = None
        self.start_time: float = 0
        self.delta: Optional[StateDelta] = None

    def __enter__(self) -> "StateCapture":
        self.before = DesktopSnapshot.capture()
        self.start_time = time.time()
        return self

    def __exit__(self, *_):
        from winscript.core.workflow import recorder
        
        after = DesktopSnapshot.capture()
        duration_ms = (time.time() - self.start_time) * 1000
        self.delta = StateDelta(
            action_tool=self.tool_name,
            action_args=self.args,
            duration_ms=duration_ms,
            before=self.before,
            after=after,
        ).compute()

        # Try to hook into the existing result passing back up the stack.
        # For simplicity, we just pass an empty result string to the recorder
        # since the tools themselves handle their return strings, but we DO
        # record that the tool was called. We'll patch tools to update the result
        # if needed, but for now, logging the action + state delta is the core.
        if recorder.is_recording:
            recorder.record_step(
                tool_name=self.tool_name,
                args=self.args,
                result="Action completed", # This gets replaced if the tool overrides it, but the tools don't currently know about the recorder.
                duration_ms=duration_ms,
                state_delta=self.delta.to_dict()
            )