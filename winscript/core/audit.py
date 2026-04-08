import sqlite3, json, time
from pathlib import Path
from contextlib import contextmanager

AUDIT_DB = Path.home() / ".winscript" / "audit.db"
AUDIT_DB.parent.mkdir(parents=True, exist_ok=True)

def _get_conn():
    conn = sqlite3.connect(str(AUDIT_DB))
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tool TEXT NOT NULL,
                args TEXT NOT NULL,
                result TEXT,
                error TEXT,
                duration_ms REAL,
                state_delta TEXT,
                selector_layer TEXT,
                timestamp REAL NOT NULL,
                session_id TEXT
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_tool ON audit_log(tool)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ts ON audit_log(timestamp)")

def log_action(
    tool: str,
    args: dict,
    result: str = "",
    error: str = "",
    duration_ms: float = 0,
    state_delta: dict = None,
    selector_layer: str = "",
    session_id: str = "",
):
    with _get_conn() as conn:
        conn.execute("""
            INSERT INTO audit_log
            (tool, args, result, error, duration_ms, state_delta, selector_layer, timestamp, session_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            tool,
            json.dumps(args, default=str),
            result[:500] if result else "",
            error[:500] if error else "",
            round(duration_ms, 2),
            json.dumps(state_delta or {}),
            selector_layer,
            time.time(),
            session_id,
        ))

def query_log(limit: int = 50, tool_filter: str = "") -> list[dict]:
    with _get_conn() as conn:
        if tool_filter:
            rows = conn.execute(
                "SELECT * FROM audit_log WHERE tool = ? ORDER BY timestamp DESC LIMIT ?",
                (tool_filter, limit)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM audit_log ORDER BY timestamp DESC LIMIT ?",
                (limit,)
            ).fetchall()
        return [dict(r) for r in rows]

def purge_old_logs(days: int = 30):
    cutoff = time.time() - (days * 86400)
    with _get_conn() as conn:
        result = conn.execute(
            "DELETE FROM audit_log WHERE timestamp < ?", (cutoff,)
        )
        return result.rowcount

def get_failure_summary() -> list[dict]:
    """Return tools with highest failure rates."""
    with _get_conn() as conn:
        rows = conn.execute("""
            SELECT tool,
                   COUNT(*) as total,
                   SUM(CASE WHEN error != '' THEN 1 ELSE 0 END) as failures,
                   AVG(duration_ms) as avg_ms
            FROM audit_log
            GROUP BY tool
            ORDER BY failures DESC
            LIMIT 20
        """).fetchall()
        return [dict(r) for r in rows]

# Run on startup
init_db()
purge_old_logs(30)
