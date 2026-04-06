import subprocess
from winscript.core.retry_guard import guard
from winscript.core.errors import WinScriptMaxRetriesError

def run_powershell(command: str) -> str:
    """Run a PowerShell command as a fallback mechanism for administration or advanced queries.
    Use sparingly when specific tools are unavailable.
    Example: run_powershell('Get-Process | Where-Object {$_.Name -eq "Code"}')"""
    args = {"command": command}
    try:
        # Run powershell command, timeout after 30 seconds
        result = subprocess.run(
            ["powershell", "-Command", command],
            capture_output=True,
            text=True,
            timeout=30
        )
        if result.returncode == 0:
            guard.record_success("run_powershell", args)
            return result.stdout.strip() if result.stdout else "Command executed successfully (no output)."
        else:
            guard.record_failure("run_powershell", args)
            return f"ERROR: PowerShell returned code {result.returncode}\nSTDOUT: {result.stdout.strip()}\nSTDERR: {result.stderr.strip()}"
    except subprocess.TimeoutExpired:
        guard.record_failure("run_powershell", args)
        return "ERROR: PowerShell command timed out after 30 seconds."
    except WinScriptMaxRetriesError:
        raise
    except Exception as e:
        guard.record_failure("run_powershell", args)
        return f"ERROR running PowerShell: {str(e)}"