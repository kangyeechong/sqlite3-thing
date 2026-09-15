"""
Small helpers for opening a connection to the SQLite ledger database
and creating its tables from schema.sql. Deliberately thin - no ORM,
so every query elsewhere in the app is plain, readable SQL.
"""

import sqlite3
from pathlib import Path

_SCHEMA_PATH = Path(__file__).parent / "schema.sql"


def get_connection(db_path: str) -> sqlite3.Connection:
    """
    Opens a connection with foreign keys enforced and rows returned as
    dict-like objects (so code can do row["po_no"] instead of row[0]).
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db(db_path: str) -> None:
    """Creates every table if it doesn't already exist."""
    schema_sql = _SCHEMA_PATH.read_text()
    conn = get_connection(db_path)
    try:
        conn.executescript(schema_sql)
        conn.commit()
    finally:
        conn.close()
