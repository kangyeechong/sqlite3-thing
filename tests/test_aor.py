"""
The AOR (Acknowledgment of Receipt) export - Kenjin's list of actual
payment receipts - fills in a contract's paid-date columns so the
existing detection pipeline picks them up exactly as if the Master
report itself had carried them. See app/aor.py for the confirmed
Reference No classification rules this exercises.

Run with: pytest tests/test_aor.py -v
"""

import datetime

import pytest

import openpyxl

from app.aor import _GREEN_FILL, _YELLOW_FILL, _classify_reference, annotate_aor_file
from app.db.connection import get_connection
from app.pipeline import process_aor_upload, process_upload
from tests.helpers import AOR_HEADERS, build_aor_report, build_master_report, confirm_all_pending


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


# --- _classify_reference: the confirmed rules, in isolation -----------

@pytest.mark.parametrize("reference,expected", [
    ("JPAY D84E4S9N (INST 15/24)", ("installments", [15])),
    ("PBL CYB0807 (INST 06/24)", ("installments", [6])),
    ("PBL CYB0807 (INST 01/24)", ("installments", [1])),
    ("TRF 14/08/2026 (INST 22/24, 23/24 & 24/24)", ("installments", [22, 23, 24])),
    # The (INST X/Y) tag wins regardless of the word in front of it.
    ("ADVANCE PARTIAL PAYMENT (INST 01/24)", ("installments", [1])),
    ("TRF 13/08/2026 ADVANCE BALANCE PAYMENT (INST 01/24)", ("installments", [1])),
    # No tag at all: FULL/BALANCE PAYMENT completes the full price.
    ("HLB 712873 FULL PAYMENT", ("full_payment", None)),
    ("G M0301 BALANCE PAYMENT", ("full_payment", None)),
    # No tag, not the completing payment: recognized, non-triggering.
    ("HLB 712873 STAMP DUTY", ("skip", None)),
    ("G M4176 DEPOSIT", ("skip", None)),
    ("G V9095 PARTIAL PAYMENT", ("skip", None)),
    # "PATRIAL PAYMENT" is a real, recurring typo in the actual export -
    # same meaning as "PARTIAL PAYMENT", not a different, unrecognized case.
    ("C M6243 PATRIAL PAYMENT", ("skip", None)),
    # Doesn't match anything known.
    ("C M6243 MYSTERY PAYMENT", ("unrecognized", None)),
    (None, ("unrecognized", None)),
    ("", ("unrecognized", None)),
])
def test_classify_reference(reference, expected):
    assert _classify_reference(reference) == expected


# --- End-to-end through the real pipeline ------------------------------

def test_aor_fills_a_blank_paid_date_and_triggers_detection(tmp_path):
    """
    The core flow: a PO exists (from a Master report upload) with no
    First Instalment Paid Date yet. An AOR receipt for that PO, tagged
    installment 1, fills it in - and since that's now a real paid-date
    on file, the exact same detection process_upload uses picks it up
    as newly due, without any AOR-specific commission logic at all.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80001, "Customer ID": "CUSTA1", "Customer Name": "Customer A1",
        "Niche/Tablet Price (RM)": 10000,  # 7.5% = 750.00
        "Agency Code": "AC001",
        # no First Instalment Paid Date yet
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0001",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80001, "Customer ID": "CUSTA1", "Customer Name": "Customer A1",
        "Reference No": "TRF 10/08/2026 (INST 01/24)",
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))

    assert result["import_result"].receipts_imported == 1
    assert result["import_result"].paid_dates_written == 1
    assert result["import_result"].review_flags == []
    raised_pos = {e["po_no"] for e in result["raised_events"]}
    assert raised_pos == {80001}
    assert result["raised_events"][0]["trigger_type"] == "installment_1"
    assert result["raised_events"][0]["amount"] == 750.0

    conn = get_connection(db_path)
    paid_date = conn.execute(
        "SELECT first_installment_paid_date FROM contracts WHERE po_no = 80001"
    ).fetchone()["first_installment_paid_date"]
    conn.close()
    assert paid_date == "2026-08-10"


def test_aor_never_overwrites_an_existing_paid_date(tmp_path):
    """
    A paid-date already on file (whether from the Master report or an
    earlier AOR upload) must never be overwritten - an AOR receipt only
    ever fills in a blank.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80002, "Customer ID": "CUSTA2", "Customer Name": "Customer A2",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": datetime.date(2026, 7, 1),  # already on file
        "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0002",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),  # different date
        "PO No": 80002, "Customer ID": "CUSTA2", "Customer Name": "Customer A2",
        "Reference No": "TRF 10/08/2026 (INST 01/24)",
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))

    assert result["import_result"].paid_dates_written == 0  # nothing to fill in, already there

    conn = get_connection(db_path)
    paid_date = conn.execute(
        "SELECT first_installment_paid_date FROM contracts WHERE po_no = 80002"
    ).fetchone()["first_installment_paid_date"]
    conn.close()
    assert paid_date == "2026-07-01"  # untouched


def test_reuploading_an_overlapping_aor_export_does_not_reapply_receipts(tmp_path):
    """
    Real AOR exports routinely overlap (the same report, exported again
    later with an extended date range) - a receipt already applied must
    never be re-processed, matching the same idempotency
    historical_summary_rows already has for Date Record rows.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80003, "Customer ID": "CUSTA3", "Customer Name": "Customer A3",
        "Niche/Tablet Price (RM)": 10000,
        "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0003",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80003, "Customer ID": "CUSTA3", "Customer Name": "Customer A3",
        "Reference No": "TRF 10/08/2026 (INST 01/24)",
    }])
    result1 = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))
    assert result1["import_result"].receipts_imported == 1
    assert result1["import_result"].paid_dates_written == 1

    result2 = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 24))
    assert result2["import_result"].receipts_imported == 0  # already applied
    assert result2["import_result"].paid_dates_written == 0
    assert result2["raised_events"] == []  # nothing new - already flagged from the first upload


def test_annotate_aor_file_scoped_to_upload_excludes_an_earlier_uploads_receipts(tmp_path):
    """
    Staff process month by month, but the real Kenjin export is
    cumulative (an "August" export re-lists every receipt back to
    whenever records began, not just August's) - when annotate_aor_file
    is given the upload it's regenerating the Filtered sheet for (see
    /download-aor-annotated/<id>), an old month's already-processed
    receipt must not show up again just because a later, wider export
    happens to repeat it.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [
        {"No": 1, "PO No": 80005, "Customer ID": "CUSTA5", "Customer Name": "Customer A5",
         "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001"},
        {"No": 2, "PO No": 80006, "Customer ID": "CUSTA6", "Customer Name": "Customer A6",
         "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001"},
    ])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    june_row = {
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-JUN01",
        "Acknowledgment Receipt Date": datetime.date(2026, 6, 10),
        "PO No": 80005, "Customer ID": "CUSTA5", "Customer Name": "Customer A5",
        "Reference No": "TRF 10/06/2026 (INST 01/24)",
    }
    xlsx_june = tmp_path / "june.xlsx"
    build_aor_report(xlsx_june, [june_row])
    process_aor_upload(db_path, str(xlsx_june), run_date=datetime.date(2026, 6, 17))

    # August's export is cumulative - it repeats June's receipt
    # unchanged AND adds a genuinely new one of its own.
    august_row = {
        "No": 2, "Acknowledgment Receipt No": "RC-TEST-AUG01",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80006, "Customer ID": "CUSTA6", "Customer Name": "Customer A6",
        "Reference No": "TRF 10/08/2026 (INST 01/24)",
    }
    xlsx_august = tmp_path / "august.xlsx"
    build_aor_report(xlsx_august, [june_row, august_row])
    result_august = process_aor_upload(db_path, str(xlsx_august), run_date=datetime.date(2026, 8, 17))
    assert result_august["import_result"].receipts_imported == 1  # only August's own new receipt

    conn = get_connection(db_path)
    output_path = tmp_path / "annotated_august.xlsx"
    annotate_aor_file(
        str(xlsx_august), str(output_path), conn=conn, aor_upload_id=result_august["aor_upload_id"],
    )
    conn.close()

    rows = _filtered_rows(openpyxl.load_workbook(output_path))
    assert [r[0] for r in rows] == ["TRF 10/08/2026 (INST 01/24)"]  # not June's repeated row


def test_annotate_aor_file_falls_back_to_unscoped_for_an_upload_predating_the_link(tmp_path):
    """
    Regression test: an aor_uploads row created before aor_upload_id
    existed on aor_receipts has no receipts linked to it at all (every
    one of them is orphaned, not genuinely empty) - scoping to that
    empty set would turn re-downloading an old upload's annotated copy
    into a blank Filtered sheet the moment this feature ships. Falls
    back to the old unscoped behavior instead.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80007, "Customer ID": "CUSTA7", "Customer Name": "Customer A7",
        "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-OLD01",
        "Acknowledgment Receipt Date": datetime.date(2026, 6, 10),
        "PO No": 80007, "Customer ID": "CUSTA7", "Customer Name": "Customer A7",
        "Reference No": "TRF 10/06/2026 (INST 01/24)",
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 6, 17))

    conn = get_connection(db_path)
    # Simulates a pre-migration upload: its receipt exists, but with no
    # link to the upload that (really did) introduce it.
    conn.execute("UPDATE aor_receipts SET aor_upload_id = NULL")
    conn.commit()

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path), conn=conn, aor_upload_id=result["aor_upload_id"])
    conn.close()

    rows = _filtered_rows(openpyxl.load_workbook(output_path))
    assert [r[0] for r in rows] == ["TRF 10/06/2026 (INST 01/24)"]  # shown, not hidden


def test_annotate_aor_file_with_no_upload_id_still_shows_everything_in_the_file(tmp_path):
    """The conn/aor_upload_id scoping is opt-in - a caller that doesn't
    pass them (matching every use of this function before this
    feature) keeps seeing exactly what's in the file, unscoped."""
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-NOSCOPE",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 90099, "Customer ID": "CUSTX", "Customer Name": "Customer X",
        "Reference No": "TRF 10/08/2026 (INST 01/24)",
    }])

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path))

    rows = _filtered_rows(openpyxl.load_workbook(output_path))
    assert [r[0] for r in rows] == ["TRF 10/08/2026 (INST 01/24)"]


def test_a_receipt_outside_the_chosen_period_is_left_for_a_later_upload(tmp_path):
    """
    The real Kenjin export is never cut on clean calendar-month
    boundaries (one real export covered 1 June - 17 Aug, the next
    18 Aug - 23 Sep), but staff process month by month regardless - a
    receipt outside the chosen period must be left completely
    untouched (not recorded into aor_receipts), so a LATER upload
    scoped to its actual period still picks it up. This is deliberately
    different from the already-imported skip: that one is permanent,
    this one is "not yet, try again with the right period."
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80008, "Customer ID": "CUSTA8", "Customer Name": "Customer A8",
        "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-PERIOD1",
        "Acknowledgment Receipt Date": datetime.date(2026, 6, 15),  # June, not August
        "PO No": 80008, "Customer ID": "CUSTA8", "Customer Name": "Customer A8",
        "Reference No": "TRF 15/06/2026 (INST 01/24)",
    }])

    # Processing "August" - this June receipt must not be touched.
    result_august = process_aor_upload(
        db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17),
        period_start="2026-08-01", period_end="2026-08-31",
    )
    assert result_august["import_result"].receipts_imported == 0
    assert result_august["import_result"].receipts_outside_period == 1
    assert result_august["import_result"].paid_dates_written == 0
    assert result_august["raised_events"] == []

    conn = get_connection(db_path)
    assert conn.execute(
        "SELECT first_installment_paid_date FROM contracts WHERE po_no = 80008"
    ).fetchone()["first_installment_paid_date"] is None
    conn.close()

    # Re-uploading the SAME file, now scoped to June (its real period) -
    # the receipt is still there, untouched, ready to be picked up.
    result_june = process_aor_upload(
        db_path, str(xlsx_aor), run_date=datetime.date(2026, 6, 17),
        period_start="2026-06-01", period_end="2026-06-30",
    )
    assert result_june["import_result"].receipts_imported == 1
    assert result_june["import_result"].receipts_outside_period == 0
    assert result_june["import_result"].paid_dates_written == 1
    assert len(result_june["raised_events"]) == 1


def test_a_june_po_paid_in_august_is_processed_when_uploading_augusts_period(tmp_path):
    """
    The whole point of period scoping being receipt-date-based, not
    PO-date-based: a June-purchased PO whose installment is actually
    paid in August must still be picked up when processing August -
    it's the payment's own date that matters, not when the PO was
    originally bought.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80009, "Customer ID": "CUSTA9", "Customer Name": "Customer A9",
        "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date(2026, 6, 30))

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-PERIOD2",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),  # paid in August
        "PO No": 80009, "Customer ID": "CUSTA9", "Customer Name": "Customer A9",
        "Reference No": "TRF 10/08/2026 (INST 01/24)",
    }])

    result_august = process_aor_upload(
        db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17),
        period_start="2026-08-01", period_end="2026-08-31",
    )
    assert result_august["import_result"].receipts_imported == 1
    assert result_august["import_result"].paid_dates_written == 1
    assert len(result_august["raised_events"]) == 1
    assert result_august["raised_events"][0]["po_no"] == 80009


def test_a_receipt_with_no_date_is_left_untouched_when_a_period_is_given(tmp_path):
    """Can't check a missing date against a period, so it's treated the
    same as outside-period rather than flagged - a period-less upload
    still flags it as a genuine data problem (see the sibling test in
    this file), but under period scoping it's just "not yet"."""
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80011, "Customer ID": "CUSTA11", "Customer Name": "Customer A11",
        "Niche/Tablet Price (RM)": 10000, "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-PERIOD3",
        "Acknowledgment Receipt Date": None,
        "PO No": 80011, "Customer ID": "CUSTA11", "Customer Name": "Customer A11",
        "Reference No": "TRF (INST 01/24)",
    }])

    result = process_aor_upload(
        db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17),
        period_start="2026-08-01", period_end="2026-08-31",
    )
    assert result["import_result"].receipts_outside_period == 1
    assert result["import_result"].review_flags == []


def test_a_split_receipt_pair_for_the_same_installment_uses_the_later_date(tmp_path):
    """
    A single logical payment can be split across two receipts on the
    same day or close together (the real file has this: "ADVANCE
    PARTIAL PAYMENT" + "ADVANCE BALANCE PAYMENT", both tagged the same
    installment number) - both just confirm "this installment
    happened," and the later date is the one that actually completed
    it.
    """
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80004, "Customer ID": "CUSTA4", "Customer Name": "Customer A4",
        "Niche/Tablet Price (RM)": 10000,
        "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [
        {
            "No": 1, "Acknowledgment Receipt No": "RC-TEST-0004A",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 13),
            "PO No": 80004, "Customer ID": "CUSTA4", "Customer Name": "Customer A4",
            "Reference No": "C V5956 ADVANCE PARTIAL PAYMENT (INST 01/24)",
        },
        {
            "No": 2, "Acknowledgment Receipt No": "RC-TEST-0004B",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 15),  # later
            "PO No": 80004, "Customer ID": "CUSTA4", "Customer Name": "Customer A4",
            "Reference No": "TRF 15/08/2026 ADVANCE BALANCE PAYMENT (INST 01/24)",
        },
    ])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))

    assert result["import_result"].receipts_imported == 2
    assert result["import_result"].paid_dates_written == 1  # one PO, one trigger

    conn = get_connection(db_path)
    paid_date = conn.execute(
        "SELECT first_installment_paid_date FROM contracts WHERE po_no = 80004"
    ).fetchone()["first_installment_paid_date"]
    conn.close()
    assert paid_date == "2026-08-15"  # the later of the two receipts


def test_unrecognized_reference_is_flagged_not_silently_applied(tmp_path):
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80005, "Customer ID": "CUSTA5", "Customer Name": "Customer A5",
        "Niche/Tablet Price (RM)": 10000,
        "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0005",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80005, "Customer ID": "CUSTA5", "Customer Name": "Customer A5",
        "Reference No": "C M6243 MYSTERY PAYMENT",  # doesn't match any known pattern
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))

    assert result["import_result"].paid_dates_written == 0
    assert len(result["import_result"].review_flags) == 1
    assert "doesn't match any known pattern" in result["import_result"].review_flags[0].message
    assert result["raised_events"] == []


def test_a_receipt_for_a_po_not_in_the_ledger_yet_is_flagged(tmp_path):
    """An AOR export can reference a PO the Master report hasn't
    brought in yet (it covers a wider company-wide window) - flagged,
    not silently dropped, so staff know to upload the Master report
    for it too."""
    db_path = _db_path(tmp_path)
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0006",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80006, "Customer ID": "CUSTA6", "Customer Name": "Customer A6",
        "Reference No": "TRF 10/08/2026 FULL PAYMENT",
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))

    assert result["import_result"].receipts_imported == 1  # still recorded, never reprocessed
    assert result["import_result"].paid_dates_written == 0
    assert len(result["import_result"].review_flags) == 1
    assert "doesn't exist in the ledger yet" in result["import_result"].review_flags[0].message


def test_a_non_1_or_6_installment_number_is_recorded_but_triggers_nothing(tmp_path):
    """Only installment 1 and 6 are fixed commission-release points -
    a receipt for any other installment number is a real, recognized
    payment (no review flag) but doesn't touch any paid-date column."""
    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80007, "Customer ID": "CUSTA7", "Customer Name": "Customer A7",
        "Niche/Tablet Price (RM)": 10000,
        "Agency Code": "AC001",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0007",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80007, "Customer ID": "CUSTA7", "Customer Name": "Customer A7",
        "Reference No": "JPAY D84E4S9N (INST 15/24)",
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))

    assert result["import_result"].receipts_imported == 1
    assert result["import_result"].paid_dates_written == 0
    assert result["import_result"].review_flags == []  # recognized, just not commission-relevant

    conn = get_connection(db_path)
    row = conn.execute(
        "SELECT first_installment_paid_date, sixth_installment_paid_date FROM contracts WHERE po_no = 80007"
    ).fetchone()
    conn.close()
    assert row["first_installment_paid_date"] is None
    assert row["sixth_installment_paid_date"] is None


def test_full_pipeline_aor_paid_date_flows_through_to_confirmed_report(tmp_path):
    """
    The full loop end to end: AOR fills in a paid-date, detection
    raises it, a human confirms it on the review page, and it shows up
    correctly in the downloaded report - agency/agent splitting and the
    Date Record summary both need zero AOR-specific code, since by the
    time the report is generated this is just a normal confirmed
    commission_event like any other.
    """
    from app.pipeline import confirm_events, generate_report, load_review

    db_path = _db_path(tmp_path)
    xlsx_master = tmp_path / "master.xlsx"
    build_master_report(xlsx_master, [{
        "No": 1, "PO No": 80008, "Customer ID": "CUSTA8", "Customer Name": "Customer A8",
        "Niche/Tablet Price (RM)": 20000,  # 7.5% = 1500.00
        "Agency Code": "AC108-02", "FCC/Agent": "AOR Test Agent",
    }])
    process_upload(db_path, str(xlsx_master), run_date=datetime.date.today())

    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-TEST-0008",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 10),
        "PO No": 80008, "Customer ID": "CUSTA8", "Customer Name": "Customer A8",
        "Reference No": "TRF 10/08/2026 (INST 06/24)",
    }])
    result = process_aor_upload(db_path, str(xlsx_aor), run_date=datetime.date(2026, 8, 17))
    run_id = result["commission_run_id"]
    assert run_id is not None

    events = load_review(db_path, run_id)
    confirm_events(db_path, run_id, {e["id"] for e in events}, "test-user")

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, run_id, str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert "AW Consultancy" in workbook.sheetnames
    sheet = workbook["AW Consultancy"]
    # Agency 3.5% = 700.00, Agent 4% = 800.00, total 1500.00 - the same
    # split math a Master-report-detected commission would get.
    values = [c.value for row in sheet.iter_rows() for c in row]
    assert 700.0 in values
    assert 800.0 in values


# --- annotate_aor_file: the downloadable annotated copy -----------------
#
# The original sheet(s) must come through completely untouched - the
# new "Filtered" sheet is the only place any coloring or row-picking
# happens.

def _filtered_rows(workbook):
    """Returns the "Filtered" sheet's data rows as (reference_no, fill)."""
    sheet = workbook["Filtered"]
    ref_col = AOR_HEADERS.index("Reference No") + 1
    rows = []
    for row_num in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row_num, column=ref_col)
        if cell.value is None:
            continue
        rows.append((cell.value, cell.fill))
    return rows


def test_annotate_aor_file_leaves_the_original_sheet_untouched(tmp_path):
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [{
        "No": 1, "Acknowledgment Receipt No": "RC-Z-0001", "PO No": 90000,
        "Customer ID": "CUSTB0", "Customer Name": "Customer B0",
        "Acknowledgment Receipt Date": datetime.date(2026, 8, 7),
        "Reference No": "TRF 07/08/2026 (INST 01/24)",
    }])

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path))

    original = openpyxl.load_workbook(xlsx_aor)
    result = openpyxl.load_workbook(output_path)

    assert result.sheetnames[0] == original.sheetnames[0]
    original_sheet = original[original.sheetnames[0]]
    result_sheet = result[original.sheetnames[0]]
    ref_col = AOR_HEADERS.index("Reference No") + 1
    # The installment-1 row would be yellow in the filtered sheet, but
    # the original sheet's own cell must be left plain - unmodified.
    assert result_sheet.cell(row=23, column=ref_col).fill.fill_type is None
    assert result_sheet.cell(row=23, column=ref_col).value == original_sheet.cell(row=23, column=ref_col).value
    assert "Filtered" in result.sheetnames
    assert "Filtered" not in original.sheetnames


def test_annotate_aor_file_colors_a_full_payment_group_green(tmp_path):
    """
    The exact real-file scenario the business described: a PARTIAL (or
    the real "PATRIAL" typo) payment followed by a BALANCE PAYMENT for
    the same PO is one completed sale - both rows land in the filtered
    sheet colored green, not just the completing row.
    """
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [
        {
            "No": 1, "Acknowledgment Receipt No": "RC-A-0091", "PO No": 90001,
            "Customer ID": "CUSTB1", "Customer Name": "Customer B1",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 7),
            "Reference No": "C M6243 PATRIAL PAYMENT",
        },
        {
            "No": 2, "Acknowledgment Receipt No": "RC-A-0092", "PO No": 90001,
            "Customer ID": "CUSTB1", "Customer Name": "Customer B1",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 8),
            "Reference No": "C V5135 BALANCE PAYMENT",
        },
    ])

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path))

    workbook = openpyxl.load_workbook(output_path)
    rows = _filtered_rows(workbook)
    assert [r[0] for r in rows] == ["C M6243 PATRIAL PAYMENT", "C V5135 BALANCE PAYMENT"]
    assert rows[0][1].fgColor.rgb == _GREEN_FILL.fgColor.rgb
    assert rows[1][1].fgColor.rgb == _GREEN_FILL.fgColor.rgb


def test_annotate_aor_file_colors_installment_1_and_6_yellow(tmp_path):
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [
        {
            "No": 1, "Acknowledgment Receipt No": "RC-B-0001", "PO No": 90002,
            "Customer ID": "CUSTB2", "Customer Name": "Customer B2",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 7),
            "Reference No": "TRF 07/08/2026 (INST 01/24)",
        },
        {
            "No": 2, "Acknowledgment Receipt No": "RC-B-0002", "PO No": 90003,
            "Customer ID": "CUSTB3", "Customer Name": "Customer B3",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 8),
            "Reference No": "TRF 08/08/2026 (INST 06/24)",
        },
    ])

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path))

    workbook = openpyxl.load_workbook(output_path)
    rows = _filtered_rows(workbook)
    assert [r[0] for r in rows] == ["TRF 07/08/2026 (INST 01/24)", "TRF 08/08/2026 (INST 06/24)"]
    assert rows[0][1].fgColor.rgb == _YELLOW_FILL.fgColor.rgb
    assert rows[1][1].fgColor.rgb == _YELLOW_FILL.fgColor.rgb


def test_annotate_aor_file_excludes_non_matching_rows_from_the_filtered_sheet(tmp_path):
    """A plain DEPOSIT that never leads to a full payment in this file,
    and an installment number that isn't 1 or 6, are real recognized
    rows - but neither is what a staff member filters for, so neither
    shows up in the filtered sheet at all."""
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [
        {
            "No": 1, "Acknowledgment Receipt No": "RC-C-0001", "PO No": 90004,
            "Customer ID": "CUSTB4", "Customer Name": "Customer B4",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 7),
            "Reference No": "G M4176 DEPOSIT",
        },
        {
            "No": 2, "Acknowledgment Receipt No": "RC-C-0002", "PO No": 90005,
            "Customer ID": "CUSTB5", "Customer Name": "Customer B5",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 8),
            "Reference No": "JPAY D84E4S9N (INST 15/24)",
        },
    ])

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path))

    workbook = openpyxl.load_workbook(output_path)
    assert _filtered_rows(workbook) == []


def test_annotate_aor_file_sorts_by_po_no_ascending(tmp_path):
    """The filtered sheet reads by PO No, not by the file's own row
    order - 20260299 before 20260300, etc - with each PO's own rows
    (e.g. a PARTIAL/BALANCE pair) still kept together."""
    xlsx_aor = tmp_path / "aor.xlsx"
    build_aor_report(xlsx_aor, [
        # Deliberately out of PO order in the source file.
        {
            "No": 1, "Acknowledgment Receipt No": "RC-D-0001", "PO No": 20260300,
            "Customer ID": "CUSTD1", "Customer Name": "Customer D1",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 7),
            "Reference No": "TRF 07/08/2026 (INST 01/24)",
        },
        {
            "No": 2, "Acknowledgment Receipt No": "RC-D-0002", "PO No": 20260299,
            "Customer ID": "CUSTD2", "Customer Name": "Customer D2",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 7),
            "Reference No": "C M0001 PATRIAL PAYMENT",
        },
        {
            "No": 3, "Acknowledgment Receipt No": "RC-D-0003", "PO No": 20260299,
            "Customer ID": "CUSTD2", "Customer Name": "Customer D2",
            "Acknowledgment Receipt Date": datetime.date(2026, 8, 8),
            "Reference No": "C M0002 BALANCE PAYMENT",
        },
    ])

    output_path = tmp_path / "annotated.xlsx"
    annotate_aor_file(str(xlsx_aor), str(output_path))

    workbook = openpyxl.load_workbook(output_path)
    filtered_sheet = workbook["Filtered"]
    po_col = AOR_HEADERS.index("PO No") + 1
    po_nos = [filtered_sheet.cell(row=r, column=po_col).value for r in range(2, filtered_sheet.max_row + 1)]
    assert po_nos == [20260299, 20260299, 20260300]
