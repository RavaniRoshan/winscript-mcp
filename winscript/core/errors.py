class WinScriptError(Exception):
    def __init__(self, tool: str, detail: str):
        self.tool = tool
        self.detail = detail
        super().__init__(f"WinScript [{tool}]: {detail}")

class WinScriptMaxRetriesError(Exception):
    def __init__(self, tool: str, args_hash: str, attempts: int):
        self.tool = tool
        self.args_hash = args_hash
        self.attempts = attempts
        super().__init__(
            f"WinScript HARD STOP: [{tool}] failed {attempts} consecutive times "
            f"with identical arguments. Try different args. "
            f"Args fingerprint: {args_hash}"
        )