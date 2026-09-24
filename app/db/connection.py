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
    # detected_by_user was in schema.sql from early on but, unlike
    # every other commission_events column, never got its own
    # _MIGRATIONS entry - a gap that only surfaced once
    # _widen_commission_events_status_check (below) started assuming
    # every column it copies already exists on the old table.
    ("commission_events", "detected_by_user",
     "ALTER TABLE commission_events ADD COLUMN detected_by_user TEXT", None),
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
    ("aor_receipts", "aor_upload_id",
     "ALTER TABLE aor_receipts ADD COLUMN aor_upload_id INTEGER REFERENCES aor_uploads(id)", None),
    ("aor_receipts", "receipt_date",
     "ALTER TABLE aor_receipts ADD COLUMN receipt_date TEXT", None),
    ("aor_receipts", "reference_text",
     "ALTER TABLE aor_receipts ADD COLUMN reference_text TEXT", None),
    ("aor_receipts", "payment_received",
     "ALTER TABLE aor_receipts ADD COLUMN payment_received NUMERIC", None),
    ("aor_receipts", "trigger_type",
     "ALTER TABLE aor_receipts ADD COLUMN trigger_type TEXT", None),
    ("commission_events", "voided_at",
     "ALTER TABLE commission_events ADD COLUMN voided_at TEXT", None),
    ("commission_events", "voided_by_user",
     "ALTER TABLE commission_events ADD COLUMN voided_by_user TEXT", None),
    ("commission_events", "void_reason",
     "ALTER TABLE commission_events ADD COLUMN void_reason TEXT", None),
    ("aor_uploads", "period_start",
     "ALTER TABLE aor_uploads ADD COLUMN period_start TEXT", None),
    ("aor_uploads", "period_end",
     "ALTER TABLE aor_uploads ADD COLUMN period_end TEXT", None),
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
    ("aor_uploads", """
        CREATE TABLE aor_uploads (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            commission_run_id  INTEGER REFERENCES commission_runs(id),
            filename           TEXT NOT NULL,
            file_bytes         BLOB NOT NULL,
            uploaded_at        TEXT NOT NULL
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

    # Must run last - it copies every column of commission_events
    # wholesale, so every ADD COLUMN above needs to have already run
    # against the old table first. See its own docstring for why this
    # can't just be another _MIGRATIONS entry.
    _widen_commission_events_status_check(conn)


def _widen_commission_events_status_check(conn: sqlite3.Connection) -> None:
    """
    SQLite has no ALTER TABLE for changing a CHECK constraint, so a
    database created by an older schema.sql (back when 'voided' wasn't
    yet a valid commission_events.status - see
    app.commission.void_commission_event) still has that column
    permanently locked to CHECK (status IN ('pending', 'confirmed')),
    even after the plain ADD COLUMN entries in _MIGRATIONS above add
    every new *column* it needs. Voiding an event on such a database
    would hit a CHECK constraint violation before void_commission_event's
    own logic ever runs - found by hand-testing against a simulated
    pre-existing database, not theoretical.

    The standard SQLite fix for a CHECK (or any constraint) that needs
    to change: rebuild the table under a temp name, copy every row
    across, drop the old one, rename the new one into place - wrapped
    in one explicit transaction so a crash mid-rebuild leaves the
    original table untouched rather than half-migrated.

    Detected by checking the table's own recorded CREATE TABLE text for
    the literal 'voided' - naturally idempotent and self-skipping on a
    brand-new database (init_db's schema.sql already allows it from the
    start, so this never runs there) without needing a separate marker
    the way the 'status' backfill does.

    Must run after every _MIGRATIONS entry above has already added its
    column to the OLD table (see _run_migrations' call order) - the
    INSERT below names every column explicitly and expects all of them
    to already exist on the table it's copying from.
    """
    row = conn.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'commission_events'"
    ).fetchone()
    if row is None or "'voided'" in row["sql"]:
        return  # table doesn't exist yet, or already allows 'voided'

    conn.execute("BEGIN")
    try:
        conn.execute(
            """
            CREATE TABLE commission_events_new (
                id                 INTEGER PRIMARY KEY AUTOINCREMENT,
                po_no              INTEGER NOT NULL REFERENCES contracts(po_no),
                trigger_type       TEXT NOT NULL CHECK (trigger_type IN ('full_payment', 'installment_1', 'installment_6')),
                trigger_date       TEXT NOT NULL,
                amount             NUMERIC NOT NULL,
                agency_amount      NUMERIC,
                agent_amount       NUMERIC,
                detected_at        TEXT NOT NULL,
                detected_by_user   TEXT,
                commission_run_id  INTEGER REFERENCES commission_runs(id),
                status             TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'confirmed', 'voided')),
                confirmed_at       TEXT,
                confirmed_by_user  TEXT,
                voided_at          TEXT,
                voided_by_user     TEXT,
                void_reason        TEXT
            )
            """
        )
        conn.execute(
            """
            INSERT INTO commission_events_new (
                id, po_no, trigger_type, trigger_date, amount, agency_amount, agent_amount,
                detected_at, detected_by_user, commission_run_id, status,
                confirmed_at, confirmed_by_user, voided_at, voided_by_user, void_reason
            )
            SELECT
                id, po_no, trigger_type, trigger_date, amount, agency_amount, agent_amount,
                detected_at, detected_by_user, commission_run_id, status,
                confirmed_at, confirmed_by_user, voided_at, voided_by_user, void_reason
            FROM commission_events
            """
        )
        conn.execute("DROP TABLE commission_events")
        conn.execute("ALTER TABLE commission_events_new RENAME TO commission_events")
        conn.commit()
    except Exception:
        conn.rollback()
        raise


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
