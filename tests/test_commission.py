"""
Direct tests of app.commission's confirm/query helpers, as opposed to
the full upload/report pipeline covered elsewhere.

Run with: pytest tests/test_commission.py -v
"""

import datetime

from app.commission import confirm_commission_events
from app.db.connection import get_connection, init_db


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def test_confirming_a_very_large_batch_does_not_exceed_sqlite_variable_limit(tmp_path):
    """
    Regression test: confirm_commission_events used to build one SQL
    statement with one bound parameter per event id and no batching.
    A "confirm all" selection larger than SQLite's default
    SQLITE_MAX_VARIABLE_NUMBER (999) would raise sqlite3.OperationalError
    and confirm nothing at all. This exercises a batch well past that
    limit and expects every row to still get confirmed.
    """
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)

    conn.execute(
        "INSERT INTO contracts (po_no, niche_price, promotion, discount, net_price) "
        "VALUES (1, 10000, 0, 0, 10000)"
    )
    event_count = 1500
    now_iso = datetime.datetime.now().isoformat()
    event_ids = []
    for _ in range(event_count):
        cursor = conn.execute(
            "INSERT INTO commission_events "
            "(po_no, trigger_type, trigger_date, amount, detected_at, status) "
            "VALUES (1, 'full_payment', '2026-01-01', 1500.0, ?, 'pending')",
            (now_iso,),
        )
        event_ids.append(cursor.lastrowid)
    conn.commit()

    confirmed_count = confirm_commission_events(conn, event_ids, "staff@xekl.example")
    conn.commit()

    assert confirmed_count == event_count
    still_pending = conn.execute(
        "SELECT COUNT(*) AS n FROM commission_events WHERE status = 'pending'"
    ).fetchone()["n"]
    assert still_pending == 0
    conn.close()


def test_confirming_an_empty_selection_does_nothing_and_does_not_error(tmp_path):
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    assert confirm_commission_events(conn, [], "staff@xekl.example") == 0
    conn.close()
