"""
Tests for app.db.connection's additive-only migration runner, focused
on the crash-safety of migrations that pair a column add with a
one-time data backfill.

Run with: pytest tests/test_migrations.py -v
"""

import sqlite3

from app.db.connection import get_connection, init_db


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def test_backfill_still_runs_if_a_previous_process_crashed_right_after_the_alter(tmp_path):
    """
    Regression test: SQLite commits an ALTER TABLE immediately and
    irreversibly, but the backfill UPDATE that has to follow it (for
    the 'status' column - see _MIGRATIONS in connection.py) is
    separate DML that only becomes durable on a later commit. The old
    migration runner inferred "already migrated" purely from column
    existence, so a crash between the two statements would leave the
    column present but every pre-existing commission_events row
    permanently stuck at the default 'pending' - silently dropping
    already-paid-out commissions out of every future report. This
    reproduces exactly that crash by hand-crafting a database in the
    "ALTER ran, backfill didn't" state and checking that the next
    get_connection() call still finishes the backfill.
    """
    db_path = _db_path(tmp_path)

    # A minimal stand-in for a real production ledger.db from before
    # the 'status' column existed at all: one contract, one
    # commission_event that was already raised (and, in the old
    # fully-automatic system, implicitly "final") under the old schema.
    raw = sqlite3.connect(db_path)
    raw.executescript(
        """
        CREATE TABLE contracts (po_no INTEGER PRIMARY KEY);
        CREATE TABLE commission_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            po_no INTEGER NOT NULL REFERENCES contracts(po_no),
            trigger_type TEXT NOT NULL,
            trigger_date TEXT NOT NULL,
            amount NUMERIC NOT NULL,
            detected_at TEXT NOT NULL,
            commission_run_id INTEGER
        );
        INSERT INTO contracts (po_no) VALUES (1);
        INSERT INTO commission_events
            (po_no, trigger_type, trigger_date, amount, detected_at)
            VALUES (1, 'full_payment', '2026-01-01', 1500.0, '2026-01-01T00:00:00');
        """
    )
    # Simulate the crash: the ALTER TABLE from the 'status' migration
    # has already run (and, being DDL, is already permanently
    # committed) but the backfill UPDATE that must follow it never got
    # to execute.
    raw.execute(
        "ALTER TABLE commission_events ADD COLUMN status TEXT NOT NULL DEFAULT 'pending'"
    )
    raw.commit()
    raw.close()

    # A restarted app reconnecting is exactly what should finish the
    # interrupted backfill.
    conn = get_connection(db_path)
    status = conn.execute(
        "SELECT status FROM commission_events WHERE po_no = 1"
    ).fetchone()["status"]
    conn.close()

    assert status == "confirmed"


def test_backfill_does_not_rerun_once_its_marker_is_recorded(tmp_path):
    """A second connection after the backfill already completed must
    not touch status again - a genuinely new 'pending' event created
    after upgrade must stay pending, not get swept up by a rerun."""
    db_path = _db_path(tmp_path)
    init_db(db_path)

    conn1 = get_connection(db_path)
    conn1.execute("INSERT INTO contracts (po_no, net_price) VALUES (1, 10000)")
    conn1.execute(
        "INSERT INTO commission_events "
        "(po_no, trigger_type, trigger_date, amount, detected_at, status) "
        "VALUES (1, 'full_payment', '2026-01-01', 1500.0, '2026-01-01T00:00:00', 'pending')"
    )
    conn1.commit()
    conn1.close()

    # A fresh connection re-runs migration checks, but the 'status'
    # migration's marker is already recorded from init_db's own first
    # connection - it must not re-run the blanket "SET status =
    # 'confirmed'" backfill and flip this genuinely-new pending row.
    conn2 = get_connection(db_path)
    status = conn2.execute(
        "SELECT status FROM commission_events WHERE po_no = 1"
    ).fetchone()["status"]
    conn2.close()

    assert status == "pending"
