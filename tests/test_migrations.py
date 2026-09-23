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
        CREATE TABLE commission_runs (id INTEGER PRIMARY KEY AUTOINCREMENT, run_date TEXT);
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


def test_an_old_database_with_the_original_status_check_can_still_be_voided(tmp_path):
    """
    Regression test: any database created by init_db() before the
    'voided' status existed has commission_events.status permanently
    locked to CHECK (status IN ('pending', 'confirmed')) - SQLite has
    no ALTER TABLE for widening a CHECK constraint, so without
    _widen_commission_events_status_check rebuilding the table, the
    very first attempt to void a confirmed event on a real, already-
    deployed database would fail with a CHECK constraint violation
    instead of ever reaching void_commission_event's own logic. Found
    by hand-testing against a simulated pre-existing database, not
    theoretical - the plain ADD COLUMN migrations alone don't touch
    the CHECK constraint at all.
    """
    db_path = _db_path(tmp_path)

    # A stand-in for a real production ledger.db from before 'voided'
    # existed - created the same way any real deployment's database
    # actually was, via a CREATE TABLE that still has the old CHECK.
    raw = sqlite3.connect(db_path)
    raw.executescript(
        """
        CREATE TABLE commission_runs (id INTEGER PRIMARY KEY AUTOINCREMENT, run_date TEXT);
        CREATE TABLE contracts (
            po_no INTEGER PRIMARY KEY, status TEXT, net_price NUMERIC,
            installment_1_commission_flagged INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE commission_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            po_no INTEGER NOT NULL REFERENCES contracts(po_no),
            trigger_type TEXT NOT NULL CHECK (trigger_type IN ('full_payment', 'installment_1', 'installment_6')),
            trigger_date TEXT NOT NULL,
            amount NUMERIC NOT NULL,
            detected_at TEXT NOT NULL,
            commission_run_id INTEGER REFERENCES commission_runs(id),
            status TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'confirmed')),
            confirmed_at TEXT,
            confirmed_by_user TEXT
        );
        INSERT INTO contracts (po_no, status, net_price, installment_1_commission_flagged)
            VALUES (1, 'active', 10000, 1);
        INSERT INTO commission_events (po_no, trigger_type, trigger_date, amount, detected_at, status)
            VALUES (1, 'installment_1', '2026-08-01', 750.0, '2026-08-01T00:00:00', 'confirmed');
        """
    )
    raw.commit()
    raw.close()

    conn = get_connection(db_path)
    event_id = conn.execute("SELECT id FROM commission_events").fetchone()["id"]

    from app.commission import void_commission_event
    voided = void_commission_event(conn, event_id, "boss@xekl.example", "wrong price")
    conn.commit()

    assert voided is True
    row = conn.execute(
        "SELECT status, void_reason FROM commission_events WHERE id = ?", (event_id,)
    ).fetchone()
    assert row["status"] == "voided"
    assert row["void_reason"] == "wrong price"
    conn.close()

    # Idempotent: a second connection against the now-rebuilt table
    # must not error or touch the data again.
    conn2 = get_connection(db_path)
    still_one_row = conn2.execute("SELECT COUNT(*) AS n FROM commission_events").fetchone()["n"]
    conn2.close()
    assert still_one_row == 1
