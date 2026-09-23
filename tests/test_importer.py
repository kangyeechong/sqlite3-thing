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


def test_a_cancelled_po_gap_infers_a_po_date_from_its_nearest_real_neighbor(tmp_path):
    """
    A synthetic placeholder has no PO Date of its own - Kenjin never
    assigned one, since the PO was never finalized. Without SOME date
    it would never show up on any period report at all (period reports
    are scoped by po_date), so it borrows the nearest real PO No's own
    date. 80071 is 1 away from 80070 (2026-06-05) and 3 away from
    80074 (2026-06-20) - the closer neighbor wins.
    """
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": 80070, "Customer ID": "CUSTH1", "Customer Name": "Customer H1",
         "Agency Code": "AC001", "PO Date": datetime.date(2026, 6, 5)},
        # 80071, 80072, 80073 missing here
        {"No": 2, "PO No": 80074, "Customer ID": "CUSTH2", "Customer Name": "Customer H2",
         "Agency Code": "AC001", "PO Date": datetime.date(2026, 6, 20)},
    ])

    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    import_master_report(conn, str(xlsx_path))
    conn.commit()

    assert conn.execute("SELECT po_date FROM contracts WHERE po_no = 80071").fetchone()["po_date"] == "2026-06-05"
    assert conn.execute("SELECT po_date FROM contracts WHERE po_no = 80073").fetchone()["po_date"] == "2026-06-20"
    conn.close()


def test_an_older_placeholder_missing_a_po_date_is_backfilled_on_a_later_upload(tmp_path):
    """
    A placeholder created before this inference existed (or simply
    surrounded by no dated neighbor at the time) still has po_date=NULL
    sitting in the ledger - a later upload that adds a dated real PO
    nearby is a fresh chance to fill it in, even though 80081 itself
    isn't a "new gap" on this second upload (it was already filled in
    on the first).
    """
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [
        {"No": 1, "PO No": 80080, "Customer ID": "CUSTH3", "Customer Name": "Customer H3", "Agency Code": "AC001"},
        # 80081 missing, no PO Date anywhere nearby yet - placeholder gets po_date=NULL
        {"No": 2, "PO No": 80082, "Customer ID": "CUSTH4", "Customer Name": "Customer H4", "Agency Code": "AC001"},
    ])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    import_master_report(conn, str(xlsx1))
    conn.commit()
    assert conn.execute("SELECT po_date FROM contracts WHERE po_no = 80081").fetchone()["po_date"] is None

    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [
        {"No": 1, "PO No": 80080, "Customer ID": "CUSTH3", "Customer Name": "Customer H3",
         "Agency Code": "AC001", "PO Date": datetime.date(2026, 7, 10)},
        {"No": 2, "PO No": 80082, "Customer ID": "CUSTH4", "Customer Name": "Customer H4", "Agency Code": "AC001"},
    ])
    import_master_report(conn, str(xlsx2))
    conn.commit()

    assert conn.execute("SELECT po_date FROM contracts WHERE po_no = 80081").fetchone()["po_date"] == "2026-07-10"
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


def test_a_cumulative_reupload_only_reports_the_genuinely_new_pos(tmp_path):
    """
    Staff process month by month, but the real Kenjin export is
    cumulative (an "August" export re-lists every PO back to whenever
    records began, not just August's new ones) - confirmed against a
    real August export that re-included all of June's rows. Re-
    uploading that file must only report the genuinely new POs in
    contracts_seen/contracts_new (and not re-run the per-row review
    checks against June's already-known, unchanged rows), even though
    every row - June's included - is still safely upserted underneath.
    """
    xlsx_june = tmp_path / "june.xlsx"
    june_rows = [
        {"No": 1, "PO No": 80060, "Customer ID": "CUSTM1", "Customer Name": "Customer M1", "Agency Code": "AC001"},
        {"No": 2, "PO No": 80061, "Customer ID": "CUSTM2", "Customer Name": "Customer M2", "Agency Code": "AC001"},
    ]
    build_master_report(xlsx_june, june_rows)
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result_june = import_master_report(conn, str(xlsx_june))
    conn.commit()
    assert result_june.contracts_seen == 2
    assert result_june.contracts_new == 2

    # August's export is cumulative - it repeats June's two rows
    # unchanged AND adds one genuinely new PO of its own.
    xlsx_august = tmp_path / "august.xlsx"
    august_rows = june_rows + [
        {"No": 3, "PO No": 80062, "Customer ID": "CUSTM3", "Customer Name": "Customer M3", "Agency Code": "AC001"},
    ]
    build_master_report(xlsx_august, august_rows)
    result_august = import_master_report(conn, str(xlsx_august))
    conn.commit()

    assert result_august.contracts_seen == 1  # only PO 80062
    assert result_august.contracts_new == 1
    assert result_august.contracts_updated == 0
    assert result_august.review_flags == []  # June's rows not re-checked

    # All three POs still genuinely exist in the ledger - nothing about
    # this scoping skipped writing June's data, only reporting it.
    count = conn.execute("SELECT COUNT(*) AS n FROM contracts").fetchone()["n"]
    assert count == 3
    conn.close()


def test_a_previously_inferred_cancelled_po_turning_real_is_still_reported(tmp_path):
    """
    A synthetic cancelled-PO placeholder (see _detect_cancelled_po_gaps)
    is already "known" to the ledger by PO No, but has no real data -
    it must NOT be treated as an already-known PO for reporting
    purposes once real data for it finally arrives (the business un-
    cancelled it, or the gap-fill was wrong): that's genuine news for a
    human to see, not an old month's PO repeating itself.
    """
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [
        {"No": 1, "PO No": 80070, "Customer ID": "CUSTM4", "Customer Name": "Customer M4", "Agency Code": "AC001"},
        # 80071 missing here - auto-filled as a cancelled placeholder
        {"No": 2, "PO No": 80072, "Customer ID": "CUSTM5", "Customer Name": "Customer M5", "Agency Code": "AC001"},
    ])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    import_master_report(conn, str(xlsx1))
    conn.commit()
    assert conn.execute("SELECT status FROM contracts WHERE po_no = 80071").fetchone()["status"] == "cancelled"

    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [{
        "No": 1, "PO No": 80071, "Customer ID": "CUSTM6", "Customer Name": "Customer M6",
        "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001",
    }])
    result2 = import_master_report(conn, str(xlsx2))
    conn.commit()

    # An UPDATE, not an INSERT - the placeholder row already existed -
    # but still counted and reported, not silently absorbed into
    # "already known": a real customer arriving for a PO that used to
    # be a no-customer placeholder is exactly the kind of change this
    # scoping exists to surface.
    assert result2.contracts_seen == 1
    assert result2.contracts_new == 0
    assert result2.contracts_updated == 1
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


def test_a_commission_paid_date_hand_typed_by_accounts_survives_a_later_master_report_reupload(tmp_path):
    """
    Same bug, same fix, for the three *_commission_paid_date columns:
    Accounts hand-types Full Commission Paid Date / 1st Half Commission
    Paid Date / Balance Half Commission Paid Date directly onto the real
    Master Report once they've actually sent the money (see the
    schema.sql comment on these columns) - Kenjin's own export doesn't
    retain that on a fresh re-generation, exactly like the three paid-
    date columns above. A later Master report re-upload for the same PO
    must not blindly overwrite an already-recorded commission-paid-date
    back to NULL.
    """
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [{
        "No": 1, "PO No": 80060, "Customer ID": "CUSTG12", "Customer Name": "Customer G12", "Agency Code": "AC001",
    }])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    import_master_report(conn, str(xlsx1))
    conn.commit()

    # Accounts hand-types this onto the real file once they've paid it.
    conn.execute(
        "UPDATE contracts SET full_commission_paid_date = '2026-07-01' WHERE po_no = 80060"
    )
    conn.commit()

    # A later Master report re-upload for the same PO - a fresh Kenjin
    # export that never carried Accounts' hand-typed edit in the first
    # place - must not wipe that commission-paid-date back to blank.
    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [{
        "No": 1, "PO No": 80060, "Customer ID": "CUSTG12", "Customer Name": "Customer G12", "Agency Code": "AC001",
    }])
    import_master_report(conn, str(xlsx2))
    conn.commit()

    row = conn.execute("SELECT full_commission_paid_date FROM contracts WHERE po_no = 80060").fetchone()
    assert row["full_commission_paid_date"] == "2026-07-01"
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


def test_new_po_date_range_covers_only_the_genuinely_new_rows(tmp_path):
    """
    Powers the results page's one-click link straight to that month's
    Overall Commission report (see app.report.generate_period_report) -
    computed from the genuinely-new rows only (same scoping as
    contracts_seen/contracts_new), not the whole file, so a cumulative
    re-upload doesn't drag an old month's dates into this upload's
    suggested range.
    """
    xlsx1 = tmp_path / "upload1.xlsx"
    build_master_report(xlsx1, [
        {"No": 1, "PO No": 80080, "Customer ID": "CUSTN1", "Customer Name": "Customer N1",
         "PO Date": datetime.date(2026, 6, 5), "Agency Code": "AC001"},
    ])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result1 = import_master_report(conn, str(xlsx1))
    conn.commit()
    assert result1.new_po_date_min == "2026-06-05"
    assert result1.new_po_date_max == "2026-06-05"

    # A cumulative August upload that repeats June's row and adds two
    # genuinely new August ones - the range must reflect only August.
    xlsx2 = tmp_path / "upload2.xlsx"
    build_master_report(xlsx2, [
        {"No": 1, "PO No": 80080, "Customer ID": "CUSTN1", "Customer Name": "Customer N1",
         "PO Date": datetime.date(2026, 6, 5), "Agency Code": "AC001"},
        {"No": 2, "PO No": 80081, "Customer ID": "CUSTN2", "Customer Name": "Customer N2",
         "PO Date": datetime.date(2026, 8, 3), "Agency Code": "AC001"},
        {"No": 3, "PO No": 80082, "Customer ID": "CUSTN3", "Customer Name": "Customer N3",
         "PO Date": datetime.date(2026, 8, 20), "Agency Code": "AC001"},
    ])
    result2 = import_master_report(conn, str(xlsx2))
    conn.commit()
    assert result2.new_po_date_min == "2026-08-03"
    assert result2.new_po_date_max == "2026-08-20"
    conn.close()


def test_new_po_date_range_is_none_when_nothing_has_a_po_date(tmp_path):
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": 80083, "Customer ID": "CUSTN4", "Customer Name": "Customer N4", "Agency Code": "AC001"},
    ])
    db_path = _db_path(tmp_path)
    init_db(db_path)
    conn = get_connection(db_path)
    result = import_master_report(conn, str(xlsx_path))
    conn.commit()
    conn.close()
    assert result.new_po_date_min is None
    assert result.new_po_date_max is None
