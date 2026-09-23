"""
Direct tests of app.commission's confirm/query helpers, as opposed to
the full upload/report pipeline covered elsewhere.

Run with: pytest tests/test_commission.py -v
"""

import datetime

from app.commission import confirm_commission_events, process_commission_run, void_commission_event
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


def _insert_contract(conn, po_no=1, first_installment_paid_date="2026-08-01"):
    conn.execute(
        "INSERT INTO contracts (po_no, status, niche_price, promotion, discount, net_price, "
        "first_installment_paid_date) VALUES (?, 'active', 10000, 0, 0, 10000, ?)",
        (po_no, first_installment_paid_date),
    )


def test_voiding_a_confirmed_event_marks_it_voided_and_clears_the_flag(tmp_path):
    """
    The whole point of voiding: it must be excluded from reports (any
    status other than 'confirmed' already achieves that - see
    app.report's `status = 'confirmed'` joins) AND clear the PO's own
    flag, so a corrected re-upload can raise this trigger again instead
    of it staying permanently marked "already handled".
    """
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    _insert_contract(conn)
    conn.commit()

    run_id, raised = process_commission_run(
        conn, as_of=datetime.date(2026, 8, 5), run_date=datetime.date(2026, 8, 5),
        source_filename="test.xlsx", created_by_user="staff@xekl.example",
    )
    conn.commit()
    event_id = raised[0]["id"]
    confirm_commission_events(conn, [event_id], "staff@xekl.example")
    conn.commit()

    assert void_commission_event(conn, event_id, "boss@xekl.example", "wrong price, re-checking") is True
    conn.commit()

    row = conn.execute(
        "SELECT status, voided_by_user, void_reason, voided_at FROM commission_events WHERE id = ?",
        (event_id,),
    ).fetchone()
    assert row["status"] == "voided"
    assert row["voided_by_user"] == "boss@xekl.example"
    assert row["void_reason"] == "wrong price, re-checking"
    assert row["voided_at"] is not None

    flag = conn.execute(
        "SELECT installment_1_commission_flagged FROM contracts WHERE po_no = 1"
    ).fetchone()["installment_1_commission_flagged"]
    assert not flag
    conn.close()


def test_voiding_a_pending_event_does_nothing(tmp_path):
    """A still-pending event was never confirmed in the first place -
    there's nothing to reverse, so voiding it is a no-op."""
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    _insert_contract(conn)
    conn.commit()
    _run_id, raised = process_commission_run(
        conn, as_of=datetime.date(2026, 8, 5), run_date=datetime.date(2026, 8, 5),
        source_filename="test.xlsx", created_by_user="staff@xekl.example",
    )
    conn.commit()
    event_id = raised[0]["id"]

    assert void_commission_event(conn, event_id, "boss@xekl.example", "mistake") is False
    status = conn.execute("SELECT status FROM commission_events WHERE id = ?", (event_id,)).fetchone()["status"]
    assert status == "pending"
    conn.close()


def test_voiding_an_already_voided_event_does_nothing(tmp_path):
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    _insert_contract(conn)
    conn.commit()
    _run_id, raised = process_commission_run(
        conn, as_of=datetime.date(2026, 8, 5), run_date=datetime.date(2026, 8, 5),
        source_filename="test.xlsx", created_by_user="staff@xekl.example",
    )
    conn.commit()
    event_id = raised[0]["id"]
    confirm_commission_events(conn, [event_id], "staff@xekl.example")
    conn.commit()
    assert void_commission_event(conn, event_id, "boss@xekl.example", "first void") is True
    conn.commit()

    assert void_commission_event(conn, event_id, "boss@xekl.example", "second void attempt") is False
    reason = conn.execute("SELECT void_reason FROM commission_events WHERE id = ?", (event_id,)).fetchone()["void_reason"]
    assert reason == "first void"  # unchanged by the second, rejected attempt
    conn.close()


def test_voiding_lets_the_po_be_detected_again(tmp_path):
    """
    The actual point of the whole feature: void a wrongly-confirmed
    event, and the next processing run must be able to raise this PO's
    trigger again as a fresh pending item - not silently do nothing
    because the flag is still set.
    """
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    _insert_contract(conn)
    conn.commit()
    _run_id, raised = process_commission_run(
        conn, as_of=datetime.date(2026, 8, 5), run_date=datetime.date(2026, 8, 5),
        source_filename="test.xlsx", created_by_user="staff@xekl.example",
    )
    conn.commit()
    event_id = raised[0]["id"]
    confirm_commission_events(conn, [event_id], "staff@xekl.example")
    conn.commit()
    void_commission_event(conn, event_id, "boss@xekl.example", "wrong price")
    conn.commit()

    # Nothing about the contract changed (still the same "mistake") -
    # a real re-upload with corrected data would look the same to
    # process_commission_run: the flag being clear is what matters.
    run_id2, raised2 = process_commission_run(
        conn, as_of=datetime.date(2026, 8, 6), run_date=datetime.date(2026, 8, 6),
        source_filename="corrected.xlsx", created_by_user="staff@xekl.example",
    )
    conn.commit()

    assert run_id2 is not None
    assert len(raised2) == 1
    assert raised2[0]["po_no"] == 1
    new_event = conn.execute(
        "SELECT status FROM commission_events WHERE id = ?", (raised2[0]["id"],)
    ).fetchone()
    assert new_event["status"] == "pending"
    conn.close()
