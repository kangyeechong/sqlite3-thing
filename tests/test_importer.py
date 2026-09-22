"""
Tests for app.importer's sheet-shape detection - specifically that it
can tell a real Commission Base Report apart from the AOR
(Acknowledgment of Receipt) export.

Run with: pytest tests/test_importer.py -v
"""

import datetime

import openpyxl
import pytest

from app.db.connection import get_connection, init_db
from app.importer import import_master_report
from tests.helpers import build_master_report

# A real AOR export's header row, exactly as Kenjin produces it - it
# happens to share "No", "PO No", and "Customer ID" with a genuine
# Master report, which used to be all the importer checked for.
_AOR_HEADERS = [
    "No", "Acknowledgment Receipt No", "OR Receipt", "Acknowledgment Receipt Date",
    "Purchase Statement No", "Purchase Statement Date", "PO No", "Lot No",
    "Customer ID", "Customer Name", "Payor Name", "Payment Mode",
    "Instalment Plan Status", "Reference No", "Payment Received (RM)",
    "Created By", "Date Created",
]


def _build_aor_file(path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "fin_dmy_collect_report-9"
    sheet.cell(row=5, column=1, value="List of Acknowledgment of Receipt Report")
    for col, header in enumerate(_AOR_HEADERS, start=1):
        sheet.cell(row=22, column=col, value=header)
    sheet.cell(row=23, column=1, value=1)
    sheet.cell(row=23, column=2, value="RC-2026-003036")
    sheet.cell(row=23, column=7, value=90001)  # PO No
    sheet.cell(row=23, column=9, value="XEKL000522")  # Customer ID
    sheet.cell(row=23, column=10, value="customer 1")  # Customer Name
    workbook.save(path)


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def test_uploading_an_aor_file_is_rejected_with_a_specific_error_not_silently_imported(tmp_path):
    """
    Regression test: an AOR export shares "No", "PO No", and
    "Customer ID" with a genuine Master report, which used to be all
    _is_master_shaped checked for - uploading the wrong file here
    silently imported one bogus "contract" per AOR receipt row (blank
    price, blank dates, wrong Customer Name casing) instead of being
    rejected. Now it must fail loudly and specifically, and the ledger
    must be left untouched.
    """
    db_path = _db_path(tmp_path)
    init_db(db_path)
    aor_path = tmp_path / "aor_export.xlsx"
    _build_aor_file(aor_path)

    conn = get_connection(db_path)
    with pytest.raises(ValueError, match="Acknowledgment of Receipt"):
        import_master_report(conn, str(aor_path))

    still_empty = conn.execute("SELECT COUNT(*) AS n FROM contracts").fetchone()["n"]
    conn.close()
    assert still_empty == 0  # nothing was silently imported from it


def test_a_genuine_master_report_still_imports_fine(tmp_path):
    """Sanity check: the stricter shape check doesn't reject real files."""
    today = datetime.date.today()
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 80001, "Customer ID": "CUSTM1", "Customer Name": "Customer M1",
        "Niche/Tablet Price (RM)": 10000,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result = import_master_report(conn, str(xlsx_path))
    conn.commit()
    conn.close()

    assert result.contracts_new == 1


def test_a_gap_in_po_no_sequence_gets_filled_as_a_cancelled_po(tmp_path):
    """
    A missing PO No between two real POs in the same upload always
    means that PO was cancelled before it was ever finalized -
    confirmed with the business. Each missing number gets its own
    synthetic contract: status 'cancelled', Remarks "Cancelled PO",
    excluded from contracts_seen/contracts_new (those describe what
    was literally in the file, not something inferred from it).
    """
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": 80010, "Customer ID": "CUSTG1", "Customer Name": "Customer G1", "Agency Code": "AC001"},
        # 80011, 80012 missing here
        {"No": 2, "PO No": 80013, "Customer ID": "CUSTG2", "Customer Name": "Customer G2", "Agency Code": "AC001"},
    ])

    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result = import_master_report(conn, str(xlsx_path))
    conn.commit()

    assert result.contracts_seen == 2  # only the 2 real rows in the file
    assert result.contracts_new == 2
    assert result.cancelled_po_gaps_detected == 2
    gap_flags = [f for f in result.review_flags if f.check == "inferred_cancelled_po"]
    assert {f.po_no for f in gap_flags} == {80011, 80012}

    for po_no in (80011, 80012):
        row = conn.execute("SELECT status, remarks, net_price FROM contracts WHERE po_no = ?", (po_no,)).fetchone()
        assert row["status"] == "cancelled"
        assert row["remarks"] == "Cancelled PO"
        assert row["net_price"] == 0.0

    conn.close()


def test_a_gap_already_filled_is_not_reflagged_on_a_later_upload(tmp_path):
    """Re-uploading (or extending) the same range must not re-create or
    re-flag a gap that was already filled in on an earlier upload."""
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [
        {"No": 1, "PO No": 80020, "Customer ID": "CUSTG3", "Customer Name": "Customer G3", "Agency Code": "AC001"},
        {"No": 2, "PO No": 80022, "Customer ID": "CUSTG4", "Customer Name": "Customer G4", "Agency Code": "AC001"},
    ])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result1 = import_master_report(conn, str(xlsx1))
    conn.commit()
    assert result1.cancelled_po_gaps_detected == 1  # PO 80021

    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [
        {"No": 1, "PO No": 80020, "Customer ID": "CUSTG3", "Customer Name": "Customer G3", "Agency Code": "AC001"},
        {"No": 2, "PO No": 80022, "Customer ID": "CUSTG4", "Customer Name": "Customer G4", "Agency Code": "AC001"},
        {"No": 3, "PO No": 80023, "Customer ID": "CUSTG5", "Customer Name": "Customer G5", "Agency Code": "AC001"},
    ])
    result2 = import_master_report(conn, str(xlsx2))
    conn.commit()

    assert result2.cancelled_po_gaps_detected == 0  # 80021 already covered, no new gaps between 80020-80023
    row = conn.execute("SELECT status FROM contracts WHERE po_no = 80021").fetchone()
    assert row["status"] == "cancelled"  # untouched, still exactly what it was
    conn.close()


def test_a_real_po_arriving_later_overwrites_its_inferred_cancelled_placeholder(tmp_path):
    """If the business un-cancels a PO (or it was a mistake), a real
    row for that same number on a later upload must overwrite the
    placeholder with real data - the normal upsert path, nothing
    special needed."""
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [
        {"No": 1, "PO No": 80030, "Customer ID": "CUSTG6", "Customer Name": "Customer G6", "Agency Code": "AC001"},
        {"No": 2, "PO No": 80032, "Customer ID": "CUSTG7", "Customer Name": "Customer G7", "Agency Code": "AC001"},
    ])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    import_master_report(conn, str(xlsx1))
    conn.commit()
    assert conn.execute("SELECT status FROM contracts WHERE po_no = 80031").fetchone()["status"] == "cancelled"

    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [{
        "No": 1, "PO No": 80031, "Customer ID": "CUSTG8", "Customer Name": "Customer G8",
        "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001",
    }])
    result2 = import_master_report(conn, str(xlsx2))
    conn.commit()

    row = conn.execute("SELECT status, net_price, customer_id FROM contracts WHERE po_no = 80031").fetchone()
    assert row["status"] == "active"
    assert row["net_price"] == 10000.0
    assert row["customer_id"] == "CUSTG8"
    assert result2.contracts_seen == 1
    assert result2.cancelled_po_gaps_detected == 0
    conn.close()


def test_a_paid_date_confirmed_by_aor_survives_a_later_master_report_reupload(tmp_path):
    """
    Regression test: the real Master report never actually carries the
    Full Settlement / First Instalment / Sixth Instalment Paid Date
    columns (confirmed against real data - every row is blank there),
    Accounts or an AOR upload fills them in later instead. A later
    Master report re-upload for the same PO used to blindly overwrite
    those three columns with whatever the new file had (blank, in
    practice) - silently wiping a paid-date an AOR upload had already
    confirmed back to NULL. Found via reproduction against real data,
    not theoretical - this only ever fills a blank now, exactly like
    the AOR importer already does.
    """
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [{
        "No": 1, "PO No": 80050, "Customer ID": "CUSTG11", "Customer Name": "Customer G11", "Agency Code": "AC001",
    }])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    import_master_report(conn, str(xlsx1))
    conn.commit()

    # An AOR upload (or Accounts, by hand) fills in the paid-date later.
    conn.execute(
        "UPDATE contracts SET first_installment_paid_date = '2026-06-03' WHERE po_no = 80050"
    )
    conn.commit()

    # A later Master report re-upload for the same PO - e.g. a
    # cumulative monthly export that includes every PO again, not just
    # new ones - must not wipe that paid-date back to blank.
    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [{
        "No": 1, "PO No": 80050, "Customer ID": "CUSTG11", "Customer Name": "Customer G11", "Agency Code": "AC001",
    }])
    import_master_report(conn, str(xlsx2))
    conn.commit()

    row = conn.execute("SELECT first_installment_paid_date FROM contracts WHERE po_no = 80050").fetchone()
    assert row["first_installment_paid_date"] == "2026-06-03"
    conn.close()


def test_an_implausibly_wide_po_gap_is_flagged_not_auto_filled(tmp_path):
    """
    A gap in the tens of thousands is far more likely to be a typo in
    one PO No (an extra digit) than that many real cancellations -
    flagged for a human instead of trying to insert that many
    synthetic rows.
    """
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": 80040, "Customer ID": "CUSTG9", "Customer Name": "Customer G9", "Agency Code": "AC001"},
        {"No": 2, "PO No": 8004099, "Customer ID": "CUSTG10", "Customer Name": "Customer G10", "Agency Code": "AC001"},
    ])

    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result = import_master_report(conn, str(xlsx_path))
    conn.commit()

    assert result.cancelled_po_gaps_detected == 0
    assert any(f.check == "po_range_too_wide_to_scan" for f in result.review_flags)
    assert not any(f.check == "inferred_cancelled_po" for f in result.review_flags)
    # Only the 2 real rows exist - nothing synthetic got created.
    count = conn.execute("SELECT COUNT(*) AS n FROM contracts").fetchone()["n"]
    assert count == 2
    conn.close()
