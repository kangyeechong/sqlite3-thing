"""
Small helpers for opening a connection to the SQLite ledger database
and creating its tables from schema.sql. Deliberately thin - no ORM,
so every query elsewhere in the app is plain, readable SQL.
"""

import sqlite3
from pathlib import Path

_SCHEMA_PATH = Path(__file__).parent / "schema.sql"

# Lightweight, additive-only migrations - no framework, since every
# change so far has just been "add a new column with a default" and a
# full migration tool would be overkill at this scale. Each entry is a
# fixed SQL literal (never built from a variable) checked against
# PRAGMA table_info before running, so it's safe to attempt on every
# connection: a column that already exists is simply skipped, and a
# table that doesn't exist yet (a brand-new, not-yet-initialized db)
# is skipped too - schema.sql will create it with every column already
# present, so there's nothing to migrate there.
#
# IMPORTANT: this is what keeps an existing ledger.db (one a user
# already has real data in) from hard-crashing the moment a new column
# gets added elsewhere in the code - init_db() only ever runs once, on
# a brand-new file, so anything relying solely on schema.sql would
# never reach a database that already existed before the change.
_MIGRATIONS = [
    ("agencies", "commission_split_type",
     ["ALTER TABLE agencies ADD COLUMN commission_split_type TEXT NOT NULL DEFAULT 'flat'"]),
    ("contracts", "fb_lead_referred",
     ["ALTER TABLE contracts ADD COLUMN fb_lead_referred INTEGER NOT NULL DEFAULT 0"]),
    ("commission_events", "agency_amount",
     ["ALTER TABLE commission_events ADD COLUMN agency_amount NUMERIC"]),
    ("commission_events", "agent_amount",
     ["ALTER TABLE commission_events ADD COLUMN agent_amount NUMERIC"]),
    ("agencies", "agency_group",
     ["ALTER TABLE agencies ADD COLUMN agency_group TEXT"]),
    ("contracts", "full_commission_paid_date",
     ["ALTER TABLE contracts ADD COLUMN full_commission_paid_date TEXT"]),
    ("contracts", "installment_1_commission_paid_date",
     ["ALTER TABLE contracts ADD COLUMN installment_1_commission_paid_date TEXT"]),
    ("contracts", "installment_6_commission_paid_date",
     ["ALTER TABLE contracts ADD COLUMN installment_6_commission_paid_date TEXT"]),
    ("commission_events", "status",
     [
         "ALTER TABLE commission_events ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'",
         # Every row that already existed the instant this column got
         # added was created under the old fully-automatic system,
         # where detection alone meant "this is due" - possibly already
         # downloaded and sent to Accounts. Backfill exactly those rows
         # (a fresh table has none, so this is a no-op there) as
         # confirmed, so turning on the review-and-confirm workflow
         # never makes an already-final commission silently disappear
         # from a report.
         "UPDATE commission_events SET status = 'confirmed'",
     ]),
    ("commission_events", "confirmed_at",
     ["ALTER TABLE commission_events ADD COLUMN confirmed_at TEXT"]),
    ("commission_events", "confirmed_by_user",
     ["ALTER TABLE commission_events ADD COLUMN confirmed_by_user TEXT"]),
]


def _run_migrations(conn: sqlite3.Connection) -> None:
    existing_tables = {row["name"] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    for table, column, statements in _MIGRATIONS:
        if table not in existing_tables:
            continue
        # PRAGMA doesn't support "?" parameter substitution, so this is
        # an f-string by necessity - safe here because `table` only
        # ever comes from the fixed _MIGRATIONS list above, never from
        # user input.
        existing_columns = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}
        if column not in existing_columns:
            for statement in statements:
                conn.execute(statement)
    conn.commit()


def get_connection(db_path: str) -> sqlite3.Connection:
    """
    Opens a connection with foreign keys enforced and rows returned as
    dict-like objects (so code can do row["po_no"] instead of row[0]).
    Also brings the database's schema up to date first (see
    _run_migrations) - every caller goes through this function, so
    this is the one place that guarantees an older database never gets
    left behind by a newer version of the code.
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    _run_migrations(conn)
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
