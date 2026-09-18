"""
The uploaded sheet carries its own trailing "Date Record" summary
table - the permanent history of every processing cycle that happened
before this tool existed. These tests cover reading it in, never
duplicating it on re-upload, and never double-counting money it
already accounts for against this tool's own fresh detection.

Run with: pytest tests/test_historical_import.py -v
"""

import datetime

from app.db.connection import get_connection
from app.pipeline import process_upload
from app.report import _load_summary_rows
from tests.helpers import build_master_report, confirm_all_pending


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def test_historical_rows_are_imported_and_shown_with_no_double_count(tmp_path):
    """
    A PO whose paid-date already falls on or before the sheet's own
    historical cutoff must not be redetected as newly due - that money
    is already counted in the historical total. Confirmed against the
    real file: an onboarding upload with a full trailing history
    detects nothing new at all, since everything is already accounted
    for.
    """
    cutoff = datetime.date(2026, 6, 5)
    # Genuine historical data always has the settlement date safely
    # before the "As at" date it was recognized on - the 5-day
    # cooling-off gate has to have already cleared for the old manual
    # process to have recognized it as due in the first place.
    settlement_date = cutoff - datetime.timedelta(days=6)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(
        xlsx_path,
        [{
            "No": 1, "PO No": 90001, "Customer ID": "CUSTH1", "Customer Name": "Customer H1",
            "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        }],
        historical_summary_rows=[{
            "date_record": cutoff,
            "full_commission": 1500.0,
            "first_half_commission": 0.0,
            "second_half_commission": 0.0,
        }],
    )

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=datetime.date.today())

    assert result["import_result"].historical_rows_imported == 1
    assert result["raised_events"] == []  # already accounted for historically, not newly due
    assert result["commission_run_id"] is None

    conn = get_connection(db_path)
    summary = _load_summary_rows(conn)
    conn.close()
    assert len(summary) == 1
    assert summary[0]["date_record"] == "As at 05/06/2026"
    assert summary[0]["full_commission"] == 1500.0
    assert summary[0]["running_total"] == 1500.0


def test_reuploading_the_same_file_does_not_duplicate_historical_rows(tmp_path):
    cutoff = datetime.date(2026, 6, 5)
    settlement_date = cutoff - datetime.timedelta(days=6)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(
        xlsx_path,
        [{
            "No": 1, "PO No": 90002, "Customer ID": "CUSTH2", "Customer Name": "Customer H2",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        }],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 1500.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )

    db_path = _db_path(tmp_path)
    result1 = process_upload(db_path, str(xlsx_path), run_date=datetime.date.today())
    assert result1["import_result"].historical_rows_imported == 1

    result2 = process_upload(db_path, str(xlsx_path), run_date=datetime.date.today())
    assert result2["import_result"].historical_rows_imported == 0  # already there, not duplicated
    assert result2["raised_events"] == []

    conn = get_connection(db_path)
    summary = _load_summary_rows(conn)
    conn.close()
    assert len(summary) == 1  # still just the one historical row, not two


def test_a_payment_after_the_historical_cutoff_is_still_detected_fresh(tmp_path):
    """
    The historical cutoff only suppresses what it actually covers - a
    PO whose paid-date falls on or before the sheet's last "As at"
    date is already accounted for, but one dated after it is a
    genuinely new payment (e.g. Accounts confirmed it after the old
    spreadsheet was last updated) and must still go through normal
    detection, not get silently swallowed by the historical import.
    """
    cutoff = datetime.date(2026, 6, 5)
    settlement_date = cutoff - datetime.timedelta(days=6)
    after_cutoff = datetime.date(2026, 9, 1)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(
        xlsx_path,
        [
            {
                "No": 1, "PO No": 90003, "Customer ID": "CUSTH3", "Customer Name": "Customer H3",
                "Niche/Tablet Price (RM)": 10000,  # already in the historical total
                "Full Settlement Paid Date": settlement_date,
                "Agency Code": "AC001",
            },
            {
                "No": 2, "PO No": 90004, "Customer ID": "CUSTH4", "Customer Name": "Customer H4",
                "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00, genuinely new
                "Full Settlement Paid Date": after_cutoff,
                "Agency Code": "AC001",
            },
        ],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 1500.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=datetime.date(2026, 9, 17))

    assert result["import_result"].historical_rows_imported == 1
    raised_pos = {e["po_no"] for e in result["raised_events"]}
    assert raised_pos == {90004}  # only the genuinely new one
    confirm_all_pending(db_path, result["commission_run_id"])

    conn = get_connection(db_path)
    summary = _load_summary_rows(conn)
    conn.close()
    assert len(summary) == 2
    assert summary[-1]["running_total"] == 3000.0  # 1500 historical + 1500 new


def test_bad_net_price_at_onboarding_does_not_permanently_lose_the_commission(tmp_path):
    """
    Regression test: a PO with a non-positive Net Price (bad source
    data - e.g. Discount larger than Niche/Tablet Price) must NOT get
    flagged by the historical import just because its paid-date falls
    on or before the cutoff. A flag is permanent and never gets unset
    anywhere in this codebase, so flagging it here - before staff have
    even had a chance to fix the data - would silently and permanently
    lose that commission even after the data is corrected and
    re-uploaded. Mirrors the exact same guarantee
    full_payment_is_due/installment_1_is_due/installment_6_is_due
    already make on their own (see commission.py).
    """
    cutoff = datetime.date(2026, 6, 5)
    settlement_date = cutoff - datetime.timedelta(days=6)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(
        xlsx_path,
        [{
            "No": 1, "PO No": 90006, "Customer ID": "CUSTH6", "Customer Name": "Customer H6",
            "Niche/Tablet Price (RM)": 10000,
            "Discount (RM)": 15000,  # net_price = -5000, bad data
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        }],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 1500.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )

    db_path = _db_path(tmp_path)
    process_upload(db_path, str(xlsx_path), run_date=datetime.date.today())

    conn = get_connection(db_path)
    flagged = conn.execute(
        "SELECT full_commission_flagged FROM contracts WHERE po_no = 90006"
    ).fetchone()["full_commission_flagged"]
    conn.close()
    assert flagged == 0  # not permanently locked out just because the data was bad

    # Now staff fix the Discount and re-upload - the commission must
    # still be raised correctly, not silently lost forever.
    xlsx_fixed = tmp_path / "upload_fixed.xlsx"
    build_master_report(xlsx_fixed, [{
        "No": 1, "PO No": 90006, "Customer ID": "CUSTH6", "Customer Name": "Customer H6",
        "Niche/Tablet Price (RM)": 10000,
        "Discount (RM)": 0,
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])
    result = process_upload(db_path, str(xlsx_fixed), run_date=datetime.date.today())
    raised_pos = {e["po_no"] for e in result["raised_events"]}
    assert raised_pos == {90006}


def test_at_need_payment_after_cutoff_is_still_detected_not_lost(tmp_path):
    """
    Regression test: full_payment_is_due's At-Need branch deliberately
    ignores `as_of` (no cooling-off wait for At-Need), so reusing it
    alone for the historical cutoff check would flag an At-Need PO as
    "already accounted for" regardless of whether its settlement date
    is actually before or after the cutoff. An At-Need payment dated
    AFTER the historical cutoff - a genuinely new payment - must still
    be detected normally, not silently and permanently lost.
    """
    cutoff = datetime.date(2026, 6, 5)
    after_cutoff = datetime.date(2026, 9, 1)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(
        xlsx_path,
        [{
            "No": 1, "PO No": 90007, "Customer ID": "CUSTH7", "Customer Name": "Customer H7",
            "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
            "Full Settlement Paid Date": after_cutoff,
            "Agency Code": "AC001",
            "Remarks": "At need case, Inurnment on 10/09/2026",
        }],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 0.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=datetime.date(2026, 9, 17))

    raised_pos = {e["po_no"] for e in result["raised_events"]}
    assert raised_pos == {90007}  # not silently swallowed by the historical cutoff


def test_a_new_po_added_after_an_already_established_cutoff_is_not_silently_lost(tmp_path):
    """
    Regression test: _flag_historically_accounted_commissions used to
    run its blanket "already accounted for" check against every active
    contract on every upload that had any historical rows at all - with
    no way to tell a PO that already existed when the cutoff was fixed
    apart from one that is brand new to the ledger this very upload. A
    PO that didn't exist yet back when a historical total was computed
    cannot possibly be money that total already covers, no matter what
    its own paid-date says - silently flagging it (with no
    commission_event ever created, and a flag that's never unset
    anywhere) would permanently lose that commission with nothing to
    review. This reproduces exactly that: upload #1 establishes the
    cutoff with one PO; upload #2 (same historical rows, nothing new
    there) introduces a brand-new PO whose paid-date predates the
    cutoff - it must still surface for review, not vanish.
    """
    cutoff = datetime.date(2026, 6, 5)
    settlement_date = cutoff - datetime.timedelta(days=6)
    xlsx_1 = tmp_path / "upload1.xlsx"
    build_master_report(
        xlsx_1,
        [{
            "No": 1, "PO No": 90008, "Customer ID": "CUSTH8", "Customer Name": "Customer H8",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        }],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 1500.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )
    db_path = _db_path(tmp_path)
    result1 = process_upload(db_path, str(xlsx_1), run_date=datetime.date.today())
    assert result1["import_result"].historical_rows_imported == 1
    assert result1["raised_events"] == []

    # Same historical rows (nothing new there), but a PO that has never
    # appeared before - with a paid-date that predates the cutoff, as
    # if it were a late Kenjin entry for an older sale.
    xlsx_2 = tmp_path / "upload2.xlsx"
    build_master_report(
        xlsx_2,
        [
            {
                "No": 1, "PO No": 90008, "Customer ID": "CUSTH8", "Customer Name": "Customer H8",
                "Niche/Tablet Price (RM)": 10000,
                "Full Settlement Paid Date": settlement_date,
                "Agency Code": "AC001",
            },
            {
                "No": 2, "PO No": 90009, "Customer ID": "CUSTH9", "Customer Name": "Customer H9",
                "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00, brand new, must not be lost
                "Full Settlement Paid Date": settlement_date,
                "Agency Code": "AC001",
            },
        ],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 1500.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )
    result2 = process_upload(db_path, str(xlsx_2), run_date=datetime.date.today())
    assert result2["import_result"].historical_rows_imported == 0  # already on file
    raised_pos = {e["po_no"] for e in result2["raised_events"]}
    assert raised_pos == {90009}  # the new PO must surface for review, not be silently lost


def test_historical_rows_only_appear_on_the_unscoped_all_view(tmp_path):
    """There's no per-agency breakdown of the sheet's own pre-existing
    history to read - only the unscoped "All" summary includes it."""
    cutoff = datetime.date(2026, 6, 5)
    settlement_date = cutoff - datetime.timedelta(days=6)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(
        xlsx_path,
        [{
            "No": 1, "PO No": 90005, "Customer ID": "CUSTH5", "Customer Name": "Customer H5",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        }],
        historical_summary_rows=[{
            "date_record": cutoff, "full_commission": 1500.0,
            "first_half_commission": 0.0, "second_half_commission": 0.0,
        }],
    )

    db_path = _db_path(tmp_path)
    process_upload(db_path, str(xlsx_path), run_date=datetime.date.today())

    conn = get_connection(db_path)
    assert len(_load_summary_rows(conn)) == 1  # unscoped: includes the historical row
    assert len(_load_summary_rows(conn, agency_group="AC001")) == 0  # scoped: no historical data to show
    conn.close()
