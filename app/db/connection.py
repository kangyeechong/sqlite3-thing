"""
Small helpers for opening a connection to the SQLite ledger database
and creating its tables from schema.sql. Deliberately thin - no ORM,
so every query elsewhere in the app is plain, readable SQL.
"""

import datetime
import sqlite3
from pathlib import Path

_SCHEMA_PATH = Path(__file__).parent / "schema.sql"

# Lightweight, additive-only migrations - no framework, since almost
# every change so far has just been "add a new column with a default"
# (_MIGRATIONS) or, occasionally, "add a whole new table"
# (_NEW_TABLE_MIGRATIONS); a full migration tool would be overkill at
# this scale. Each entry is a fixed SQL literal (never built from a
# variable) checked against PRAGMA table_info / sqlite_master before
# running, so it's safe to attempt on every connection: a column or
# table that already exists is simply skipped. A brand-new,
# not-yet-initialized db has nothing to migrate either way - schema.sql
# (via init_db) creates every table with every column already present.
#
# IMPORTANT: this is what keeps an existing ledger.db (one a user
# already has real data in) from hard-crashing the moment a new column
# gets added elsewhere in the code - init_db() only ever runs once, on
# a brand-new file, so anything relying solely on schema.sql would
# never reach a database that already existed before the change.
#
# Each entry is (table, column, add_column_sql, backfill_statements).
# add_column_sql runs only when the column doesn't exist yet (SQLite
# commits an ALTER TABLE immediately - it can't be rolled back, so
# column-existence is a safe, natural "already done" signal for it).
# backfill_statements is None for a plain new column, or a list of DML
# statements for a migration that also has to fix up existing rows
# (e.g. the 'status' column below). Those are NOT safe to gate on
# column-existence alone: a crash after the ALTER but before the
# backfill would otherwise skip the backfill forever, since the column
# would already appear to exist on every later connection. They're
# gated on the schema_migrations marker table instead - see
# _run_migrations.
_MIGRATIONS = [
    ("agencies", "commission_split_type",
     "ALTER TABLE agencies ADD COLUMN commission_split_type TEXT NOT NULL DEFAULT 'flat'", None),
    ("contracts", "fb_lead_referred",
     "ALTER TABLE contracts ADD COLUMN fb_lead_referred INTEGER NOT NULL DEFAULT 0", None),
    ("commission_events", "agency_amount",
     "ALTER TABLE commission_events ADD COLUMN agency_amount NUMERIC", None),
    ("commission_events", "agent_amount",
     "ALTER TABLE commission_events ADD COLUMN agent_amount NUMERIC", None),
    ("agencies", "agency_group",
     "ALTER TABLE agencies ADD COLUMN agency_group TEXT", None),
    ("contracts", "full_commission_paid_date",
     "ALTER TABLE contracts ADD COLUMN full_commission_paid_date TEXT", None),
    ("contracts", "installment_1_commission_paid_date",
     "ALTER TABLE contracts ADD COLUMN installment_1_commission_paid_date TEXT", None),
    ("contracts", "installment_6_commission_paid_date",
     "ALTER TABLE contracts ADD COLUMN installment_6_commission_paid_date TEXT", None),
    ("commission_events", "status",
     "ALTER TABLE commission_events ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'",
     [
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
     "ALTER TABLE commission_events ADD COLUMN confirmed_at TEXT", None),
    ("commission_events", "confirmed_by_user",
     "ALTER TABLE commission_events ADD COLUMN confirmed_by_user TEXT", None),
]

# For a brand-new table (not a new column on an existing one) - same
# additive spirit as _MIGRATIONS above, just CREATE TABLE instead of
# ALTER TABLE ADD COLUMN. Every statement here is a fixed literal.
# schema_migrations itself is listed here too, so an existing database
# upgrading straight into this version gets the marker table created
# before _MIGRATIONS below ever needs to read or write it.
_NEW_TABLE_MIGRATIONS = [
    ("historical_summary_rows", """
        CREATE TABLE historical_summary_rows (
            id                      INTEGER PRIMARY KEY AUTOINCREMENT,
            date_record             TEXT NOT NULL UNIQUE,
            full_commission         NUMERIC NOT NULL DEFAULT 0,
            first_half_commission   NUMERIC NOT NULL DEFAULT 0,
            second_half_commission  NUMERIC NOT NULL DEFAULT 0,
            remarks                 TEXT,
            imported_at             TEXT NOT NULL,
            imported_by_user        TEXT,
            source_filename         TEXT
        )
    """),
    ("schema_migrations", """
        CREATE TABLE schema_migrations (
            key        TEXT PRIMARY KEY,
            applied_at TEXT NOT NULL
        )
    """),
    ("aor_receipts", """
        CREATE TABLE aor_receipts (
            id                         INTEGER PRIMARY KEY AUTOINCREMENT,
            acknowledgment_receipt_no  TEXT NOT NULL UNIQUE,
            po_no                      INTEGER,
            imported_at                TEXT NOT NULL,
            imported_by_user           TEXT,
            source_filename            TEXT
        )
    """),
]


def _run_migrations(conn: sqlite3.Connection) -> None:
    existing_tables = {row["name"] for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}

    for table, create_sql in _NEW_TABLE_MIGRATIONS:
        if table not in existing_tables:
            # A CREATE TABLE (DDL) commits itself immediately in
            # SQLite regardless of any explicit commit() call, so two
            # connections racing this same check on a brand-new db
            # could both see the table missing and both try to create
            # it. Harmless either way - the loser just hits "table
            # already exists" - but it should never surface as an
            # unhandled crash for what is, in the end, a no-op.
            try:
                conn.execute(create_sql)
            except sqlite3.OperationalError as exc:
                if "already exists" not in str(exc):
                    raise
            existing_tables.add(table)
    conn.commit()

    applied = {row["key"] for row in conn.execute("SELECT key FROM schema_migrations")}

    for table, column, add_column_sql, backfill_statements in _MIGRATIONS:
        if table not in existing_tables:
            continue
        # PRAGMA doesn't support "?" parameter substitution, so this is
        # an f-string by necessity - safe here because `table` only
        # ever comes from the fixed _MIGRATIONS list above, never from
        # user input.
        existing_columns = {row["name"] for row in conn.execute(f"PRAGMA table_info({table})")}
        if column not in existing_columns:
            try:
                conn.execute(add_column_sql)
            except sqlite3.OperationalError as exc:
                # Same race as above, for a concurrent ALTER TABLE.
                if "duplicate column name" not in str(exc):
                    raise

        if backfill_statements is None:
            continue

        key = f"{table}.{column}"
        if key in applied:
            continue
        for statement in backfill_statements:
            conn.execute(statement)
        conn.execute(
            "INSERT OR IGNORE INTO schema_migrations (key, applied_at) VALUES (?, ?)",
            (key, datetime.datetime.now().isoformat()),
        )
        # Committed immediately (not batched with the rest of this
        # function) so the backfill and its completion marker land
        # together - if the process dies before this line, the marker
        # is simply absent and the backfill safely reruns next time;
        # once this line completes, it can never be skipped again.
        conn.commit()
        applied.add(key)


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
