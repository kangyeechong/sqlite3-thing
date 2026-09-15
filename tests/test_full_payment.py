"""
Step 1 verification: full-payment flagging.

These are the scenarios from the incremental build plan - fake data,
not real customer data, so correctness can be checked without reading
implementation code. Run with: pytest tests/test_full_payment.py -v
"""

import datetime

from app.pipeline import process_upload
from app.db.connection import get_connection
from tests.helpers import build_master_report


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def test_full_payment_flagged_once_cooling_off_gate_clears(tmp_path):
    """
    A Pre-Need full-payment contract, settled 6 days ago (past the
    5-day gate), should be flagged as newly due for 15% of Net Price.
    """
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90001, "Customer ID": "CUST001",
        "Customer Name": "Customer 1",
        "Niche/Tablet Price (RM)": 20000, "Promotion (RM)": 0, "Discount (RM)": 500,
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today, created_by_user="tester@test.com")

    assert result["import_result"].contracts_new == 1
    assert len(result["raised_events"]) == 1
    event = result["raised_events"][0]
    assert event["po_no"] == 90001
    assert event["trigger_type"] == "full_payment"
    # Net Price = 20000 - 0 - 500 = 19500; 15% = 2925.00
    assert event["amount"] == 2925.0


def test_full_payment_not_flagged_inside_cooling_off_window(tmp_path):
    """A contract settled only 2 days ago must not be flagged yet."""
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=2)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90002, "Customer ID": "CUST002",
        "Customer Name": "Customer 2",
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert result["raised_events"] == []


def test_full_payment_not_flagged_twice_across_uploads(tmp_path):
    """
    The same contract appearing in two consecutive uploads (the normal
    case - each new export includes everything, not just what's new)
    must only ever be flagged once.
    """
    db_path = _db_path(tmp_path)
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    row = {
        "No": 1, "PO No": 90003, "Customer ID": "CUST003",
        "Customer Name": "Customer 3",
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }

    first_upload = tmp_path / "upload1.xlsx"
    build_master_report(first_upload, [row])
    first_result = process_upload(db_path, str(first_upload), run_date=today)
    assert len(first_result["raised_events"]) == 1

    # Second cycle: same PO, same (still-filled-in) settlement date -
    # exactly what a re-exported report looks like.
    second_upload = tmp_path / "upload2.xlsx"
    build_master_report(second_upload, [row])
    later = today + datetime.timedelta(days=15)
    second_result = process_upload(db_path, str(second_upload), run_date=later)

    assert second_result["raised_events"] == []


def test_at_need_flagged_immediately_with_inurnment_date(tmp_path):
    """
    At-Need contracts skip the cooling-off wait entirely, but require
    an inurnment date to be readable from Remarks.
    """
    today = datetime.date(2026, 8, 20)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90004, "Customer ID": "CUST004",
        "Customer Name": "Customer 4",
        "Full Settlement Paid Date": today,  # settled today - zero days elapsed
        "Agency Code": "AC001",
        "Remarks": "At need case\nInurnment on 20/08/2026",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert len(result["raised_events"]) == 1
    assert result["raised_events"][0]["po_no"] == 90004


def test_at_need_without_readable_inurnment_date_is_not_flagged_and_is_reviewed(tmp_path):
    """
    An At-Need case where the inurnment date couldn't be parsed must
    NOT be flagged (fails toward the safe direction: delay, don't
    release early) - and must show up in the review panel so a human
    fills it in.
    """
    today = datetime.date(2026, 8, 20)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90005, "Customer ID": "CUST005",
        "Customer Name": "Customer 5",
        "Full Settlement Paid Date": today,
        "Agency Code": "AC001",
        "Remarks": "At need case - inurnment date TBC",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert result["raised_events"] == []
    flags = [f.check for f in result["import_result"].review_flags]
    assert "at_need_missing_inurnment_date" in flags


def test_cancelled_po_excluded_even_if_settled(tmp_path):
    """A PO whose Remarks say it's cancelled must never be flagged."""
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=30)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90006, "Customer ID": "CUST006",
        "Customer Name": "Customer 6",
        "Full Settlement Paid Date": settlement_date,
        "Remarks": "Cancelled PO",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert result["raised_events"] == []
    conn = get_connection(_db_path(tmp_path))
    row = conn.execute("SELECT status FROM contracts WHERE po_no = 90006").fetchone()
    assert row["status"] == "cancelled"


def test_reinstated_po_becomes_eligible_again(tmp_path):
    """
    Status is re-derived fresh on every import, not a one-way lock -
    so a PO whose Remarks no longer say "cancelled" naturally becomes
    active again and can be flagged, without any special-case code.
    """
    db_path = _db_path(tmp_path)
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    cancelled_upload = tmp_path / "upload1.xlsx"
    build_master_report(cancelled_upload, [{
        "No": 1, "PO No": 90007, "Customer ID": "CUST007",
        "Customer Name": "Customer 7",
        "Full Settlement Paid Date": settlement_date,
        "Remarks": "Cancelled PO",
    }])
    process_upload(db_path, str(cancelled_upload), run_date=today)

    reinstated_upload = tmp_path / "upload2.xlsx"
    build_master_report(reinstated_upload, [{
        "No": 1, "PO No": 90007, "Customer ID": "CUST007",
        "Customer Name": "Customer 7",
        "Full Settlement Paid Date": settlement_date,
        "Remarks": None,  # remarks cleared - PO put back to active
    }])
    result = process_upload(db_path, str(reinstated_upload), run_date=today)

    assert len(result["raised_events"]) == 1
    assert result["raised_events"][0]["po_no"] == 90007
    flags = [f.check for f in result["import_result"].review_flags]
    assert "status_changed" in flags  # surfaced for a human to confirm, not silent


def test_ac001_seeded_as_no_agent_split_new_agency_seeded_as_split(tmp_path):
    """
    AC001 (XEMP, in-house staff) must default to splits_by_agent=False
    on first sight; any other new agency code defaults to True.
    """
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": 90020, "Customer ID": "CUST020", "Customer Name": "Customer 20", "Agency Code": "AC001"},
        {"No": 2, "PO No": 90021, "Customer ID": "CUST021", "Customer Name": "Customer 21", "Agency Code": "AC999"},
    ])

    db_path = _db_path(tmp_path)
    process_upload(db_path, str(xlsx_path))

    conn = get_connection(db_path)
    ac001 = conn.execute("SELECT splits_by_agent FROM agencies WHERE agency_code = 'AC001'").fetchone()
    ac999 = conn.execute("SELECT splits_by_agent FROM agencies WHERE agency_code = 'AC999'").fetchone()
    assert ac001["splits_by_agent"] == 0
    assert ac999["splits_by_agent"] == 1


def test_review_panel_flags_duplicate_po_in_same_upload(tmp_path):
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": 90008, "Customer ID": "CUST008", "Customer Name": "Customer 8"},
        {"No": 2, "PO No": 90008, "Customer ID": "CUST008", "Customer Name": "Customer 8"},
    ])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    checks = [f.check for f in result["import_result"].review_flags]
    assert "duplicate_po" in checks


def test_review_panel_flags_non_positive_net_price(tmp_path):
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90009, "Customer ID": "CUST009", "Customer Name": "Customer 9",
        "Niche/Tablet Price (RM)": 1000, "Discount (RM)": 1500,  # discount exceeds price
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    checks = [f.check for f in result["import_result"].review_flags]
    assert "non_positive_net_price" in checks


def test_review_panel_flags_blank_agency_code_on_active_po(tmp_path):
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90010, "Customer ID": "CUST010", "Customer Name": "Customer 10",
        "Agency Code": None,
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    checks = [f.check for f in result["import_result"].review_flags]
    assert "blank_agency_code" in checks


def test_bogus_inurnment_date_is_rejected(tmp_path):
    """
    A placeholder like "99/99/9999" matches the date pattern's shape
    but isn't a real date - must be treated the same as no date at
    all (no early release, flagged for review), not accepted at face
    value.
    """
    today = datetime.date(2026, 8, 20)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90012, "Customer ID": "CUST012",
        "Customer Name": "Customer 12",
        "Full Settlement Paid Date": today,
        "Remarks": "At need case\nInurnment on 99/99/9999",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert result["raised_events"] == []
    flags = [f.check for f in result["import_result"].review_flags]
    assert "at_need_missing_inurnment_date" in flags


def test_negative_net_price_is_never_commissioned_and_stays_correctable(tmp_path):
    """
    A discount larger than the price (a data-entry mistake) must never
    produce a commission event, and must NOT burn the flag - once the
    source data is fixed and re-uploaded, this PO must still be
    eligible to be correctly flagged.
    """
    db_path = _db_path(tmp_path)
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    bad_upload = tmp_path / "upload1.xlsx"
    build_master_report(bad_upload, [{
        "No": 1, "PO No": 90013, "Customer ID": "CUST013", "Customer Name": "Customer 13",
        "Niche/Tablet Price (RM)": 1000, "Discount (RM)": 1500,  # net_price = -500
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])
    bad_result = process_upload(db_path, str(bad_upload), run_date=today)

    assert bad_result["raised_events"] == []

    fixed_upload = tmp_path / "upload2.xlsx"
    build_master_report(fixed_upload, [{
        "No": 1, "PO No": 90013, "Customer ID": "CUST013", "Customer Name": "Customer 13",
        "Niche/Tablet Price (RM)": 10000, "Discount (RM)": 500,  # corrected
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])
    fixed_result = process_upload(db_path, str(fixed_upload), run_date=today)

    assert len(fixed_result["raised_events"]) == 1
    assert fixed_result["raised_events"][0]["po_no"] == 90013


def test_row_number_as_float_does_not_truncate_the_import(tmp_path):
    """
    openpyxl can hand back a General-formatted numeric cell as a float
    (1.0) rather than an int (1) - the row-continuation check must
    still recognize it as valid, or the whole import silently stops
    after zero rows.
    """
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1.0, "PO No": 90014, "Customer ID": "CUST014", "Customer Name": "Customer 14"},
        {"No": 2.0, "PO No": 90015, "Customer ID": "CUST015", "Customer Name": "Customer 15"},
    ])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    assert result["import_result"].contracts_seen == 2


def test_row_with_missing_po_no_is_skipped_not_fatal(tmp_path):
    """One row with no PO No must not take down the rest of the upload."""
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {"No": 1, "PO No": None, "Customer ID": "CUST016", "Customer Name": "Customer 16"},
        {"No": 2, "PO No": 90017, "Customer ID": "CUST017", "Customer Name": "Customer 17"},
    ])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    assert result["import_result"].contracts_seen == 1
    assert result["import_result"].contracts_new == 1
    checks = [f.check for f in result["import_result"].review_flags]
    assert "missing_po_no" in checks


def test_commission_rounds_half_up_at_the_half_cent_boundary():
    from app.commission import calculate_full_payment_commission
    # 333.3 * 15% = 49.995 exactly - half-up must round to 50.00, not
    # 49.99 (which is what binary-float round() can produce).
    assert calculate_full_payment_commission(333.3) == 50.0


def test_review_panel_does_not_flag_blank_agency_code_on_cancelled_po(tmp_path):
    """Per the business: a blank agency on a cancelled PO is expected, not a data gap."""
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 90011, "Customer ID": "CUST011", "Customer Name": "Customer 11",
        "Agency Code": None,
        "Remarks": "Cancelled PO",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    checks = [f.check for f in result["import_result"].review_flags]
    assert "blank_agency_code" not in checks
