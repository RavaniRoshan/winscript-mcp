"""
USAGE PATTERN:
    from winscript.core.retry_guard import guard
    from winscript.core.errors import WinScriptMaxRetriesError

    def some_tool(app_title: str) -> str:
        args = {"app_title": app_title}
        try:
            result = do_the_thing(app_title)
            guard.record_success("some_tool", args)
            return result
        except WinScriptMaxRetriesError:
            raise  # always propagate
        except Exception as e:
            guard.record_failure("some_tool", args)
            return f"ERROR: {str(e)}"
"""

import hashlib
import json
from collections import defaultdict
from .errors import WinScriptMaxRetriesError

MAX_CONSECUTIVE_FAILURES = 5

class RetryGuard:
    def __init__(self):
        self._failures: dict[str, int] = defaultdict(int)

    def _key(self, tool_name: str, args: dict) -> str:
        args_str = json.dumps(args, sort_keys=True, default=str)
        args_hash = hashlib.md5(args_str.encode()).hexdigest()[:8]
        return f"{tool_name}:{args_hash}"

    def record_failure(self, tool_name: str, args: dict) -> None:
        key = self._key(tool_name, args)
        self._failures[key] += 1
        if self._failures[key] >= MAX_CONSECUTIVE_FAILURES:
            _, args_hash = key.split(":", 1)
            raise WinScriptMaxRetriesError(
                tool=tool_name,
                args_hash=args_hash,
                attempts=self._failures[key]
            )

    def record_success(self, tool_name: str, args: dict) -> None:
        self._failures.pop(self._key(tool_name, args), None)

    def reset_all(self) -> None:
        self._failures.clear()

# Singleton — shared across all tool modules
guard = RetryGuard()