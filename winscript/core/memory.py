import sqlite3, json, time
from pathlib import Path

MEMORY_DB = Path.home() / ".winscript" / "memory.db"
MEMORY_DB.parent.mkdir(parents=True, exist_ok=True)

def _conn():
    c = sqlite3.connect(str(MEMORY_DB))
    c.row_factory = sqlite3.Row
    return c

def init_memory():
    with _conn() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS window_memory (
                title TEXT PRIMARY KEY,
                last_seen REAL,
                seen_count INTEGER DEFAULT 1
            );
            CREATE TABLE IF NOT EXISTS file_memory (
                path TEXT PRIMARY KEY,
                last_opened REAL,
                open_count INTEGER DEFAULT 1,
                extension TEXT
            );
            CREATE TABLE IF NOT EXISTS action_memory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tool TEXT,
                args TEXT,
                result TEXT,
                timestamp REAL
            );
        """)

def remember_window(title: str):
    with _conn() as c:
        c.execute("""
            INSERT INTO window_memory (title, last_seen, seen_count)
            VALUES (?, ?, 1)
            ON CONFLICT(title) DO UPDATE SET
                last_seen=excluded.last_seen,
                seen_count=seen_count+1
        """, (title, time.time()))

def remember_file(path: str):
    ext = Path(path).suffix.lstrip(".")
    with _conn() as c:
        c.execute("""
            INSERT INTO file_memory (path, last_opened, open_count, extension)
            VALUES (?, ?, 1, ?)
            ON CONFLICT(path) DO UPDATE SET
                last_opened=excluded.last_opened,
                open_count=open_count+1
        """, (path, time.time(), ext))

def remember_action(tool: str, args: dict, result: str):
    with _conn() as c:
        c.execute(
            "INSERT INTO action_memory (tool, args, result, timestamp) VALUES (?,?,?,?)",
            (tool, json.dumps(args, default=str), result[:200], time.time())
        )
        # Keep only last 100 actions
        c.execute("""
            DELETE FROM action_memory WHERE id NOT IN (
                SELECT id FROM action_memory ORDER BY timestamp DESC LIMIT 100
            )
        """)

def recall_windows(limit: int = 10) -> list[dict]:
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM window_memory ORDER BY last_seen DESC LIMIT ?", (limit,)
        ).fetchall()
        return [dict(r) for r in rows]

def recall_files(limit: int = 10, extension: str = "") -> list[dict]:
    with _conn() as c:
        if extension:
            rows = c.execute(
                "SELECT * FROM file_memory WHERE extension=? ORDER BY last_opened DESC LIMIT ?",
                (extension, limit)
            ).fetchall()
        else:
            rows = c.execute(
                "SELECT * FROM file_memory ORDER BY last_opened DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(r) for r in rows]

def recall_actions(limit: int = 20) -> list[dict]:
    with _conn() as c:
        rows = c.execute(
            "SELECT * FROM action_memory ORDER BY timestamp DESC LIMIT ?", (limit,)
        ).fetchall()
        return [dict(r) for r in rows]

init_memory()