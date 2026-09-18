"""
remove_contracts.py: the one-time cleanup tool for data that should
never have been imported (e.g. a sample/test file uploaded into a real
database by mistake) - not how a real cancelled PO gets handled (those
stay forever, see docs/data_model.md).

Run with: pytest tests/test_remove_contracts.py -v
"""

import subprocess
import sys

from app.db.connection import get_connection, init_db


def _seed(db_path):
    init_db(db_path)
    conn = get_connection(db_path)
    conn.execute("INSERT INTO customers (customer_id, name) VALUES ('TEST101', 'Test A')")
    conn.execute("INSERT INTO customers (customer_id, name) VALUES ('XEKL000401', 'Real Customer')")
    conn.execute(
        "INSERT INTO contracts (po_no, customer_id, agent_name, net_price) "
        "VALUES (90101, 'TEST101', 'Test Agent', 1000)"
    )
    conn.execute(
        "INSERT INTO contracts (po_no, customer_id, agent_name, net_price) "
        "VALUES (20240170, 'XEKL000401', 'agent 1', 5000)"
    )
    conn.execute(
        "INSERT INTO commission_events (po_no, trigger_type, trigger_date, amount, detected_at) "
        "VALUES (90101, 'full_payment', '2026-01-01', 150, '2026-01-01T00:00:00')"
    )
    conn.commit()
    conn.close()


def _run(db_path, args, confirm=True):
    return subprocess.run(
        [sys.executable, "remove_contracts.py", "--db", db_path, *args],
        input="yes\n" if confirm else "no\n",
        capture_output=True, text=True,
    )


def test_removes_only_the_specified_po_and_its_events(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    _seed(db_path)

    result = _run(db_path, ["--po", "90101"])
    assert result.returncode == 0
    assert "Removed 1 contract(s)" in result.stdout

    conn = get_connection(db_path)
    remaining = {r["po_no"] for r in conn.execute("SELECT po_no FROM contracts")}
    remaining_events = conn.execute("SELECT COUNT(*) AS n FROM commission_events").fetchone()["n"]
    conn.close()
    assert remaining == {20240170}  # the real PO is untouched
    assert remaining_events == 0


def test_declining_confirmation_deletes_nothing(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    _seed(db_path)

    result = _run(db_path, ["--po", "90101"], confirm=False)
    assert "Cancelled - nothing was deleted" in result.stdout

    conn = get_connection(db_path)
    remaining = {r["po_no"] for r in conn.execute("SELECT po_no FROM contracts")}
    conn.close()
    assert remaining == {90101, 20240170}


def test_customer_id_prefix_matches_without_naming_every_po(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    _seed(db_path)

    result = _run(db_path, ["--customer-id-prefix", "TEST"])
    assert result.returncode == 0
    assert "PO 90101" in result.stdout

    conn = get_connection(db_path)
    remaining = {r["po_no"] for r in conn.execute("SELECT po_no FROM contracts")}
    conn.close()
    assert remaining == {20240170}


def test_no_matches_deletes_nothing_and_does_not_crash(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    _seed(db_path)

    result = _run(db_path, ["--po", "99999"])
    assert result.returncode == 0
    assert "nothing to remove" in result.stdout.lower()

    conn = get_connection(db_path)
    remaining = {r["po_no"] for r in conn.execute("SELECT po_no FROM contracts")}
    conn.close()
    assert remaining == {90101, 20240170}
