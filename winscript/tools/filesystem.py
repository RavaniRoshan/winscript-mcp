import shutil
from pathlib import Path

def read_file_text(path: str, max_chars: int = 10000) -> str:
    """Read text from file. Truncates at max_chars.
    Example: read_file_text("C:/docs/report.txt")"""
    try:
        from winscript.core.memory import remember_file
        remember_file(path)
        text = Path(path).read_text(encoding="utf-8", errors="replace")
        if len(text) > max_chars:
            text = text[:max_chars] + f"\n... [truncated at {max_chars} chars]"
        return text
    except Exception as e:
        return f"ERROR reading {path}: {str(e)}"

def write_file_text(path: str, content: str) -> str:
    """Write text to file. Creates file and parent dirs if needed.
    Example: write_file_text("C:/docs/output.txt", "Hello world")"""
    try:
        from winscript.core.memory import remember_file
        remember_file(path)
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        Path(path).write_text(content, encoding="utf-8")
        return f"Written {len(content)} chars to {path}"
    except Exception as e:
        return f"ERROR writing {path}: {str(e)}"

def list_dir(path: str) -> str:
    """List files and folders in a directory.
    Example: list_dir("C:/Users/Roshan/Documents")"""
    try:
        entries = sorted(Path(path).iterdir())
        lines = []
        for e in entries:
            tag = "[DIR]" if e.is_dir() else "[FILE]"
            size = f" ({e.stat().st_size} bytes)" if e.is_file() else ""
            lines.append(f"{tag} {e.name}{size}")
        return "\n".join(lines) if lines else "(empty directory)"
    except Exception as e:
        return f"ERROR listing {path}: {str(e)}"

def move_file(src: str, dst: str) -> str:
    """Move a file from src to dst. Example: move_file("C:/a.txt","C:/b/a.txt")"""
    try:
        shutil.move(src, dst)
        return f"Moved {src} → {dst}"
    except Exception as e:
        return f"ERROR moving: {str(e)}"

def copy_file(src: str, dst: str) -> str:
    """Copy a file. Example: copy_file("C:/a.txt","C:/backup/a.txt")"""
    try:
        shutil.copy2(src, dst)
        return f"Copied {src} → {dst}"
    except Exception as e:
        return f"ERROR copying: {str(e)}"

def delete_file(path: str) -> str:
    """Delete a file. Does NOT delete directories.
    Example: delete_file("C:/tmp/old.txt")"""
    try:
        p = Path(path)
        if not p.exists():
            return f"ERROR: {path} does not exist"
        if p.is_dir():
            return f"ERROR: {path} is a directory — refusing to delete"
        p.unlink()
        return f"Deleted {path}"
    except Exception as e:
        return f"ERROR deleting: {str(e)}"

def file_exists(path: str) -> str:
    """Check if file or directory exists. Example: file_exists("C:/docs/report.txt")"""
    p = Path(path)
    if p.exists():
        kind = "directory" if p.is_dir() else "file"
        return f"EXISTS ({kind}): {path}"
    return f"NOT FOUND: {path}"