"""
Step 3 verification: the Excel report handed to Accounts, in the real
Master Report column layout.

Run with: pytest tests/test_report.py -v
"""

import datetime

import openpyxl
import pytest

from app.pipeline import process_upload, generate_report
from tests.helpers import build_master_report, confirm_all_pending


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def _find_table_rows(sheet, header_marker="PO No"):
    """
    Returns (headers, data_rows) for the first PO-level table on a
    sheet, as a list of dicts keyed by column header - so tests don't
    have to hardcode column positions.
    """
    header_row_num = None
    for row in sheet.iter_rows():
        if any(cell.value == header_marker for cell in row):
            header_row_num = row[0].row
            break
    assert header_row_num is not None, f"No header row found on sheet {sheet.title}"

    headers = [cell.value for cell in sheet[header_row_num]]
    data_rows = []
    for row in sheet.iter_rows(min_row=header_row_num + 1):
        values = [cell.value for cell in row]
        if values[headers.index("No")] is None:
            break
        data_rows.append(dict(zip(headers, values)))
    return headers, data_rows


def test_report_has_an_all_sheet_and_a_sheet_per_agency(tmp_path):
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60001, "Customer ID": "CUST201", "Customer Name": "Customer 201",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60002, "Customer ID": "CUST202", "Customer Name": "Customer 202",
            "Niche/Tablet Price (RM)": 20000,
            "First Instalment Paid Date": today,
            "Agency Code": "AC200",
            "FCC/Agent": "Agent Alpha",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    # AC001 (XEMP) never splits by agent, so it's just one sheet. AC200
    # is a brand-new agency (defaults to splits_by_agent=True), so it
    # gets its own combined group sheet PLUS a separate sheet for its
    # one agent, "Agent Alpha".
    assert set(workbook.sheetnames) == {"All", "AC001", "AC200", "Agent Alpha"}

    _, all_rows = _find_table_rows(workbook["All"])
    po_numbers = {r["PO No"] for r in all_rows}
    assert po_numbers == {60001, 60002}


def test_cancelled_po_stays_visible_beige_with_commission_cleared(tmp_path):
    """
    The report is a standing ledger, not a due-items list - a cancelled
    PO must never disappear. Confirmed with the business: it stays
    listed under its agency, shaded beige, with its commission figure
    cleared (not just tinted over) since nothing is owed on it any
    more.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60030, "Customer ID": "CUST230", "Customer Name": "Customer 230",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60031, "Customer ID": "CUST231", "Customer Name": "Customer 231",
            "Niche/Tablet Price (RM)": 15000,
            "Agency Code": "AC001",
            "Remarks": "Cancelled by customer request",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    confirm_all_pending(db_path, result["commission_run_id"])

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["AC001"]
    headers, ac001_rows = _find_table_rows(sheet)
    po_numbers = {r["PO No"] for r in ac001_rows}
    assert po_numbers == {60030, 60031}  # the cancelled PO is still there

    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    cancelled_row = next(r for r in ac001_rows if r["PO No"] == 60031)
    assert cancelled_row["Full Payment Commission (RM)"] is None  # value cleared, not just tinted

    data_row_num = header_row_num + [r["PO No"] for r in ac001_rows].index(60031) + 1
    po_col = headers.index("PO No") + 1
    assert sheet.cell(row=data_row_num, column=po_col).fill.start_color.rgb in ("00FBE5D6", "FFFBE5D6")

    # Cancelled PO's (would-be) 1500.00 commission must not count
    # toward the Total row - only the active PO's 1500.00 does.
    total_row_num = header_row_num + len(ac001_rows) + 1
    commission_col = headers.index("Full Payment Commission (RM)") + 1
    assert sheet.cell(row=total_row_num, column=commission_col).value == 1500.0


def test_a_po_cancelled_after_confirmation_is_excluded_from_the_movement_line_too(tmp_path):
    """
    Regression test: status is re-derived from Remarks on every
    upload, not a one-way lock - a PO confirmed as due in one run can
    later be marked cancelled. The Total row already excluded a
    cancelled PO's commission, but the "movement as at" line (the
    figure Accounts actually pays out for this cycle) and the cell's
    yellow highlight did not - re-downloading an earlier run's report
    after the PO was cancelled would still tell Accounts to pay out a
    commission that no longer exists.
    """
    today = datetime.date.today()

    xlsx1 = tmp_path / "run1.xlsx"
    build_master_report(xlsx1, [{
        "No": 1, "PO No": 60040, "Customer ID": "CUST240", "Customer Name": "Customer 240",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])
    db_path = _db_path(tmp_path)
    result1 = process_upload(db_path, str(xlsx1), run_date=today)
    confirm_all_pending(db_path, result1["commission_run_id"])

    # A later upload marks the same PO cancelled - status is re-derived
    # from Remarks every time, not a one-way lock.
    xlsx2 = tmp_path / "run2.xlsx"
    build_master_report(xlsx2, [{
        "No": 1, "PO No": 60040, "Customer ID": "CUST240", "Customer Name": "Customer 240",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
        "Remarks": "Cancelled by customer request",
    }])
    process_upload(db_path, str(xlsx2), run_date=today)

    # Re-download run 1's report now that the PO has since been
    # cancelled - the standing ledger always reflects current status.
    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result1["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["AC001"]
    headers, ac001_rows = _find_table_rows(sheet)
    row = next(r for r in ac001_rows if r["PO No"] == 60040)
    assert row["1st Half Commission (RM)"] is None  # cleared, cancelled

    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    commission_col = headers.index("1st Half Commission (RM)") + 1
    data_row_num = header_row_num + [r["PO No"] for r in ac001_rows].index(60040) + 1
    total_row_num = header_row_num + len(ac001_rows) + 1
    movement_row_num = total_row_num + 1

    assert sheet.cell(row=total_row_num, column=commission_col).value == 0
    assert sheet.cell(row=movement_row_num, column=commission_col).value == 0  # not 750 - cancelled money isn't a payout
    cell_fill = sheet.cell(row=data_row_num, column=commission_col).fill.start_color.rgb
    assert cell_fill not in ("00FFFF00", "FFFFFF00")  # not yellow - nothing to highlight as newly due


def test_unpaid_po_stays_visible_with_no_color(tmp_path):
    """A PO with nothing paid yet must still appear - not omitted, and
    not colored, since nothing has happened on it either way."""
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60032, "Customer ID": "CUST232", "Customer Name": "Customer 232",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": today - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60033, "Customer ID": "CUST233", "Customer Name": "Customer 233",
            "Niche/Tablet Price (RM)": 10000,
            # nothing paid at all
            "Agency Code": "AC001",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    confirm_all_pending(db_path, result["commission_run_id"])

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["AC001"]
    headers, ac001_rows = _find_table_rows(sheet)
    assert {r["PO No"] for r in ac001_rows} == {60032, 60033}

    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    data_row_num = header_row_num + [r["PO No"] for r in ac001_rows].index(60033) + 1
    po_col = headers.index("PO No") + 1
    assert sheet.cell(row=data_row_num, column=po_col).fill.start_color.rgb in ("00000000", None)


def test_an_earlier_confirmation_shows_plain_on_a_later_download_not_yellow(tmp_path):
    """
    Confirmed against the real file: once something is confirmed, it
    keeps showing on every future download (this is a standing ledger),
    but only the run that actually confirmed it gets to paint it
    yellow - an older confirmed figure must carry forward plainly, not
    get re-highlighted every single time someone downloads again.
    """
    db_path = _db_path(tmp_path)

    day1 = datetime.date.today() - datetime.timedelta(days=20)
    xlsx1 = tmp_path / "run1.xlsx"
    build_master_report(xlsx1, [{
        "No": 1, "PO No": 60034, "Customer ID": "CUST234", "Customer Name": "Customer 234",
        "Niche/Tablet Price (RM)": 10000,
        "Full Settlement Paid Date": day1 - datetime.timedelta(days=6),
        "Agency Code": "AC001",
    }])
    result1 = process_upload(db_path, str(xlsx1), run_date=day1)
    confirm_all_pending(db_path, result1["commission_run_id"])

    day2 = datetime.date.today()
    xlsx2 = tmp_path / "run2.xlsx"
    build_master_report(xlsx2, [
        {
            "No": 1, "PO No": 60034, "Customer ID": "CUST234", "Customer Name": "Customer 234",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": day1 - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60035, "Customer ID": "CUST235", "Customer Name": "Customer 235",
            "Niche/Tablet Price (RM)": 20000,
            "Full Settlement Paid Date": day2 - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
    ])
    result2 = process_upload(db_path, str(xlsx2), run_date=day2)
    confirm_all_pending(db_path, result2["commission_run_id"])

    report_path = tmp_path / "report2.xlsx"
    generate_report(db_path, result2["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["AC001"]
    headers, ac001_rows = _find_table_rows(sheet)
    assert {r["PO No"] for r in ac001_rows} == {60034, 60035}  # run 1's PO still shows up

    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    commission_col = headers.index("Full Payment Commission (RM)") + 1
    old_po_row_num = header_row_num + [r["PO No"] for r in ac001_rows].index(60034) + 1
    new_po_row_num = header_row_num + [r["PO No"] for r in ac001_rows].index(60035) + 1

    # Both are full-payment rows, so both are green (fully paid off is
    # a lifetime fact) - but full payment doesn't get its own cell
    # highlight either way (whole row green, no separate yellow), so
    # what distinguishes "confirmed this run" is the movement line.
    assert sheet.cell(row=old_po_row_num, column=commission_col).value == 1500.0
    assert sheet.cell(row=new_po_row_num, column=commission_col).value == 3000.0

    header_row_num_2 = header_row_num
    total_row_num = header_row_num_2 + len(ac001_rows) + 1
    movement_row_num = total_row_num + 1
    assert sheet.cell(row=total_row_num, column=commission_col).value == 4500.0  # lifetime total
    assert sheet.cell(row=movement_row_num, column=commission_col).value == 3000.0  # just run 2's own addition


def test_a_po_with_two_triggers_in_one_run_is_a_single_row_not_two(tmp_path):
    """
    Matches the real Master Report's layout: a PO due for both
    instalment 1 and instalment 6 in the same run gets ONE row with
    both commission columns filled in, not two separate rows.
    """
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60003, "Customer ID": "CUST203", "Customer Name": "Customer 203",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": today - datetime.timedelta(days=60),
        "Sixth Instalment Paid Date": today - datetime.timedelta(days=5),
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, all_rows = _find_table_rows(workbook["All"])
    assert len(all_rows) == 1
    row = all_rows[0]
    assert row["1st Half Commission (RM)"] == 750
    assert row["Balance Half Commission (RM)"] == 750
    # The commission-paid-date columns are left for Accounts to fill in later.
    assert row["Full Commission Paid Date"] is None
    assert row["1st Half Commission Paid Date"] is None
    assert row["Balance Half Commission Paid Date"] is None


def test_commission_paid_dates_are_read_back_from_the_sheet(tmp_path):
    """
    Full Commission Paid Date / 1st Half / Balance Half Paid Date are
    Accounts' own manual entries on the real file - this tool never
    computes or writes them, but once Accounts has filled them in and
    the file gets re-uploaded, the report should carry that "confirmed
    paid" status forward rather than leaving the column blank forever.
    """
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60027, "Customer ID": "CUST227", "Customer Name": "Customer 227",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": today - datetime.timedelta(days=60),
        "Sixth Instalment Paid Date": today - datetime.timedelta(days=5),
        "1st Half Commission Paid Date": today - datetime.timedelta(days=50),
        "Balance Half Commission Paid Date": today,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, all_rows = _find_table_rows(workbook["All"])
    row = all_rows[0]
    assert row["Full Commission Paid Date"] is None  # Accounts never filled this one in
    assert row["1st Half Commission Paid Date"] == (today - datetime.timedelta(days=50)).isoformat()
    assert row["Balance Half Commission Paid Date"] == today.isoformat()


def test_full_payment_row_is_shaded_green_matching_the_real_file(tmp_path):
    """
    Confirmed against the real sample file's actual cell formatting
    (not guessed): full-payment rows there are shaded green across the
    whole row - not just the commission cell.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60004, "Customer ID": "CUST204", "Customer Name": "Customer 204",
        "Niche/Tablet Price (RM)": 10000,
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["All"]
    headers, _ = _find_table_rows(sheet)
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    data_row_num = header_row_num + 1

    commission_col = headers.index("Full Payment Commission (RM)") + 1
    unrelated_col = headers.index("Customer Name") + 1

    assert sheet.cell(row=data_row_num, column=commission_col).fill.start_color.rgb in ("00C6DEB5", "FFC6DEB5")
    # Whole row, not just the commission cell.
    assert sheet.cell(row=data_row_num, column=unrelated_col).fill.start_color.rgb in ("00C6DEB5", "FFC6DEB5")


def test_balance_half_row_is_also_shaded_green_not_just_full_payment(tmp_path):
    """
    Confirmed with the business: green means the PO is fully paid off
    commission-wise, not specifically "one-off full payment" only. An
    installment plan's Balance Half (installment 6) is its last ever
    commission trigger - nothing more is due on that PO after this, so
    it gets the same whole-row green as a full payment does, even
    though no "Full Payment Commission" was ever raised on this PO.
    """
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60022, "Customer ID": "CUST222", "Customer Name": "Customer 222",
        "Niche/Tablet Price (RM)": 10000,
        "Sixth Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["All"]
    headers, _ = _find_table_rows(sheet)
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    data_row_num = header_row_num + 1

    unrelated_col = headers.index("Customer Name") + 1
    balance_col = headers.index("Balance Half Commission (RM)") + 1

    # Whole row green, including columns unrelated to the trigger itself.
    assert sheet.cell(row=data_row_num, column=unrelated_col).fill.start_color.rgb in ("00C6DEB5", "FFC6DEB5")
    # The Balance Half cell itself: yellow (newly due this run) wins
    # over the row's green, same rule as any other highlighted cell.
    assert sheet.cell(row=data_row_num, column=balance_col).fill.start_color.rgb in ("00FFFF00", "FFFFFF00")


def test_instalment_commission_cell_is_highlighted_yellow_not_the_whole_row(tmp_path):
    """
    Instalments get a yellow highlight on just the specific commission
    cell that became due, not the whole row - unlike full payment.
    """
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60016, "Customer ID": "CUST216", "Customer Name": "Customer 216",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["All"]
    headers, _ = _find_table_rows(sheet)
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    data_row_num = header_row_num + 1

    commission_col = headers.index("1st Half Commission (RM)") + 1
    unrelated_col = headers.index("Customer Name") + 1

    assert sheet.cell(row=data_row_num, column=commission_col).fill.start_color.rgb in ("00FFFF00", "FFFFFF00")
    assert sheet.cell(row=data_row_num, column=unrelated_col).fill.start_color.rgb in ("00000000", None)


def test_yellow_survives_on_a_row_that_is_also_shaded_green(tmp_path):
    """
    A PO whose very first import already has both full payment and an
    instalment due (both triggers are checked independently, with no
    rule against both firing at once) must still show the instalment
    cell in yellow, not have the row's green silently swallow it.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60017, "Customer ID": "CUST217", "Customer Name": "Customer 217",
        "Niche/Tablet Price (RM)": 10000,
        "Full Settlement Paid Date": settlement_date,
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    # Sanity check this scenario actually raises both triggers.
    assert {e["trigger_type"] for e in result["raised_events"]} == {"full_payment", "installment_1"}

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["All"]
    headers, _ = _find_table_rows(sheet)
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    data_row_num = header_row_num + 1

    instalment_col = headers.index("1st Half Commission (RM)") + 1
    unrelated_col = headers.index("Customer Name") + 1

    assert sheet.cell(row=data_row_num, column=instalment_col).fill.start_color.rgb in ("00FFFF00", "FFFFFF00")
    assert sheet.cell(row=data_row_num, column=unrelated_col).fill.start_color.rgb in ("00C6DEB5", "FFC6DEB5")


def test_cooling_off_period_shows_expired_on_an_instalment_row_too(tmp_path):
    """
    Confirmed against the real Master Report: Cooling Off Period is a
    general per-PO field (EXPIRED once COOLING_OFF_TOTAL_DAYS have
    passed since Signature Date), shown on every row old enough - not
    just on rows where a full-payment trigger happens to fire. A PO
    whose only trigger this run is an instalment must still show
    EXPIRED here if its signature date is old enough.
    """
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60020, "Customer ID": "CUST220", "Customer Name": "Customer 220",
        "Niche/Tablet Price (RM)": 10000,
        "Signature Date": today - datetime.timedelta(days=60),
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, all_rows = _find_table_rows(workbook["All"])
    row = next(r for r in all_rows if r["PO No"] == 60020)
    assert row["Cooling Off Period"] == "EXPIRED"


def test_cooling_off_period_is_blank_when_signature_date_is_too_recent(tmp_path):
    """The flip side: a PO signed only a few days ago hasn't cleared
    the cooling-off window yet, so the column stays blank, not EXPIRED."""
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60021, "Customer ID": "CUST221", "Customer Name": "Customer 221",
        "Niche/Tablet Price (RM)": 10000,
        "Signature Date": today - datetime.timedelta(days=2),
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, all_rows = _find_table_rows(workbook["All"])
    row = next(r for r in all_rows if r["PO No"] == 60021)
    assert row["Cooling Off Period"] is None


def test_no_split_agency_sheet_is_one_flat_table(tmp_path):
    """AC001 (XEMP) must not be broken down by agent - one flat table."""
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60005, "Customer ID": "CUST205", "Customer Name": "Customer 205",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001", "FCC/Agent": "Staff A",
        },
        {
            "No": 2, "PO No": 60006, "Customer ID": "CUST206", "Customer Name": "Customer 206",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001", "FCC/Agent": "Staff B",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, ac001_rows = _find_table_rows(workbook["AC001"])
    assert len(ac001_rows) == 2  # one flat table, both agents mixed together


def test_splitting_agency_has_a_combined_sheet_and_separate_agent_sheets(tmp_path):
    """
    An agency with splits_by_agent on gets its own combined sheet with
    every agent's rows together FIRST, then separate standalone sheets
    per individual agent - not sub-sections within one sheet.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60007, "Customer ID": "CUST207", "Customer Name": "Customer 207",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC210", "FCC/Agent": "Agent One",
        },
        {
            "No": 2, "PO No": 60008, "Customer ID": "CUST208", "Customer Name": "Customer 208",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC210", "FCC/Agent": "Agent Two",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert set(workbook.sheetnames) == {"All", "AC210", "Agent One", "Agent Two"}

    # The combined "AC210" sheet has both agents' rows together, flat.
    _, ac210_rows = _find_table_rows(workbook["AC210"])
    assert {r["PO No"] for r in ac210_rows} == {60007, 60008}

    # Each agent also gets their own standalone sheet with just their row.
    _, agent_one_rows = _find_table_rows(workbook["Agent One"])
    assert {r["PO No"] for r in agent_one_rows} == {60007}
    _, agent_two_rows = _find_table_rows(workbook["Agent Two"])
    assert {r["PO No"] for r in agent_two_rows} == {60008}


def test_agency_group_combines_subcodes_into_one_sheet_before_agent_sheets(tmp_path):
    """
    Several agency_codes can be sub-codes of one real-world agency (AW
    Consultancy's AC108-01/-02/-03) - they must land on ONE combined
    "AW Consultancy" sheet together, not separate AC108-01/AC108-02
    sheets, with individual agent sheets still following after.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60018, "Customer ID": "CUST218", "Customer Name": "Customer 218",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC108-01", "FCC/Agent": "Agent AW1",
        },
        {
            "No": 2, "PO No": 60019, "Customer ID": "CUST219", "Customer Name": "Customer 219",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC108-02", "FCC/Agent": "Agent AW2",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert "AC108-01" not in workbook.sheetnames
    assert "AC108-02" not in workbook.sheetnames
    assert set(workbook.sheetnames) == {"All", "AW Consultancy", "Agent AW1", "Agent AW2"}

    _, group_rows = _find_table_rows(workbook["AW Consultancy"])
    assert {r["PO No"] for r in group_rows} == {60018, 60019}

    _, agent_aw1_rows = _find_table_rows(workbook["Agent AW1"])
    assert {r["PO No"] for r in agent_aw1_rows} == {60018}
    _, agent_aw2_rows = _find_table_rows(workbook["Agent AW2"])
    assert {r["PO No"] for r in agent_aw2_rows} == {60019}


def test_aw_consultancy_sheet_has_the_agency_agent_split_columns(tmp_path):
    """
    Confirmed against the real file: an agency_agent_split group's
    sheet (and its per-agent sheets) get 9 extra columns to the right
    of the main table - Full Payment/First Half/Balance Half, each
    broken into (FB-lead deduction, Agency %, Agent %). A flat agency's
    sheet (AC001) must NOT get these extra columns at all.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60023, "Customer ID": "CUST223", "Customer Name": "Customer 223",
            "Niche/Tablet Price (RM)": 20000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC108-01", "FCC/Agent": "Agent AW1",
        },
        {
            "No": 2, "PO No": 60024, "Customer ID": "CUST224", "Customer Name": "Customer 224",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)

    aw_sheet = workbook["AW Consultancy"]
    header_row_num = next(row[0].row for row in aw_sheet.iter_rows() if any(c.value == "PO No" for c in row))
    headers = [cell.value for cell in aw_sheet[header_row_num]]
    assert headers[25:33] == [
        "3% FB leads from XEKL  (to be deducted from AW Consultancy)", "AW Consultancy\n7%", "Agent\n8%",
        "1.5% FB leads from XEKL  (to be deducted from AW Consultancy)", "AW Consultancy\n3.5%", "Agent\n4%",
        "1.5% FB leads from XEKL  (to be deducted from AW Consultancy)", "AW Consultancy\n3.5%",
    ]
    super_header_row = aw_sheet[1]
    assert super_header_row[25].value == "Full Payment Commissioin (RM)"

    data_row = aw_sheet[header_row_num + 1]
    assert data_row[26].value == 1400  # 20000 * 7%
    assert data_row[27].value == 1600  # 20000 * 8%

    total_row = aw_sheet[header_row_num + 2]
    assert total_row[26].value == 1400
    assert total_row[27].value == 1600

    ac001_sheet = workbook["AC001"]
    ac001_header_row_num = next(row[0].row for row in ac001_sheet.iter_rows() if any(c.value == "PO No" for c in row))
    ac001_headers = [cell.value for cell in ac001_sheet[ac001_header_row_num]]
    assert len(ac001_headers) == 24  # no split columns tacked on for a flat agency


def test_fb_lead_deduction_shows_amount_and_no_deduction_is_greyed_out(tmp_path):
    """
    fb_lead_referred is a purely manual flag (no Excel column for it),
    so this test sets it directly rather than through an upload.
    Confirmed against the real file: when set, the FB-lead deduction
    cell shows the actual RM amount deducted from the agency's share;
    when not set, that cell is left blank and greyed out (this shade
    was approximated from a screenshot, not matched byte-for-byte
    against a real .xlsx like green/yellow/beige were).
    """
    import app.db.connection as db_connection
    from app import commission
    from app.report import generate_commission_run_report

    db_path = str(tmp_path / "ledger.db")
    db_connection.init_db(db_path)
    conn = db_connection.get_connection(db_path)
    today = datetime.date.today()
    settlement = (today - datetime.timedelta(days=6)).isoformat()

    conn.execute(
        "INSERT INTO agencies (agency_code, splits_by_agent, commission_split_type, agency_group) "
        "VALUES ('AC108-01', 1, 'agency_agent_split', 'AW Consultancy')"
    )
    conn.execute(
        "INSERT INTO contracts (po_no, agent_name, agency_code, net_price, case_type, status, "
        "full_settlement_paid_date, fb_lead_referred) VALUES (93001, 'Agent AW1', 'AC108-01', "
        "20000, 'pre_need', 'active', ?, 1)",
        (settlement,),
    )
    conn.execute(
        "INSERT INTO contracts (po_no, agent_name, agency_code, net_price, case_type, status, "
        "full_settlement_paid_date, fb_lead_referred) VALUES (93002, 'Agent AW1', 'AC108-01', "
        "20000, 'pre_need', 'active', ?, 0)",
        (settlement,),
    )
    conn.commit()

    run_id, raised = commission.process_commission_run(
        conn, as_of=today, run_date=today, source_filename="test", created_by_user="test"
    )
    commission.confirm_commission_events(conn, [e["id"] for e in raised], "test")
    conn.commit()

    report_path = tmp_path / "report.xlsx"
    generate_commission_run_report(conn, run_id, str(report_path))
    conn.close()

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["AW Consultancy"]
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))

    referred_row = sheet[header_row_num + 1]
    assert referred_row[25].value == 600  # 20000 * 3%, deducted
    assert referred_row[25].fill.fill_type is None

    not_referred_row = sheet[header_row_num + 2]
    assert not_referred_row[25].value is None
    assert not_referred_row[25].fill.start_color.rgb in ("00BFBFBF", "FFBFBFBF")


def test_report_total_matches_sum_of_commission_columns(tmp_path):
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60009, "Customer ID": "CUST209", "Customer Name": "Customer 209",
            "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60010, "Customer ID": "CUST210", "Customer Name": "Customer 210",
            "Niche/Tablet Price (RM)": 20000,  # 15% = 3000.00
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, all_rows = _find_table_rows(workbook["All"])
    sheet = workbook["All"]
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    total_row_num = header_row_num + len(all_rows) + 1
    headers = [cell.value for cell in sheet[header_row_num]]
    total_row_values = dict(zip(headers, [cell.value for cell in sheet[total_row_num]]))
    assert total_row_values["Customer Name"] == "Total"
    # Niche/Promotion/Discount/Nett Price each get their own real sum
    # too, not just the three commission columns - confirmed against
    # the real file's own Total row.
    assert total_row_values["Niche/Tablet Price (RM)"] == 30000.0
    assert total_row_values["Promotion (RM)"] == 0.0
    assert total_row_values["Discount (RM)"] == 0.0
    assert total_row_values["Nett Price (RM)"] == 30000.0
    assert total_row_values["Full Payment Commission (RM)"] == 4500.0
    assert total_row_values["1st Half Commission (RM)"] == 0.0
    assert total_row_values["Balance Half Commission (RM)"] == 0.0


def test_summary_table_accumulates_across_multiple_runs(tmp_path):
    """
    The Date Record summary table reflects every run ever processed,
    not just the current one, with a correctly accumulating running
    total.
    """
    db_path = _db_path(tmp_path)

    day1 = datetime.date.today() - datetime.timedelta(days=20)
    xlsx1 = tmp_path / "run1.xlsx"
    build_master_report(xlsx1, [{
        "No": 1, "PO No": 60011, "Customer ID": "CUST211", "Customer Name": "Customer 211",
        "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
        "Full Settlement Paid Date": day1 - datetime.timedelta(days=6),
        "Agency Code": "AC001",
    }])
    result1 = process_upload(db_path, str(xlsx1), run_date=day1)
    confirm_all_pending(db_path, result1["commission_run_id"])

    day2 = datetime.date.today()
    xlsx2 = tmp_path / "run2.xlsx"
    build_master_report(xlsx2, [
        {
            "No": 1, "PO No": 60011, "Customer ID": "CUST211", "Customer Name": "Customer 211",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": day1 - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60012, "Customer ID": "CUST212", "Customer Name": "Customer 212",
            "Niche/Tablet Price (RM)": 20000,  # 15% = 3000.00
            "Full Settlement Paid Date": day2 - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
    ])
    result2 = process_upload(db_path, str(xlsx2), run_date=day2)

    report_path = tmp_path / "report2.xlsx"
    confirm_all_pending(db_path, result2["commission_run_id"])
    generate_report(db_path, result2["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    all_values = [tuple(r) for r in workbook["All"].iter_rows(values_only=True) if any(v is not None for v in r)]
    summary_rows = [r for r in all_values if isinstance(r[0], str) and r[0].startswith("As at")]

    assert len(summary_rows) == 2
    assert summary_rows[0][4] == 1500.0   # running total after run 1
    assert summary_rows[1][4] == 4500.0   # running total after run 2 (1500 + 3000)

    grand_total_row = next(r for r in all_values if isinstance(r[0], str) and r[0].startswith("Total Sum of Commission Payout"))
    assert grand_total_row[4] == 4500.0


def test_only_the_newest_summary_row_is_highlighted_yellow(tmp_path):
    """
    Confirmed against the real file: every prior "As at" row in the
    Date Record table stays plain once it's been through a download -
    only the row for whichever run this download is actually for (and
    the grand total line under it) gets shaded yellow. History is never
    overwritten, and yellow never spreads to rows that were already
    there before this run.
    """
    db_path = _db_path(tmp_path)

    day1 = datetime.date.today() - datetime.timedelta(days=20)
    xlsx1 = tmp_path / "run1.xlsx"
    build_master_report(xlsx1, [{
        "No": 1, "PO No": 60025, "Customer ID": "CUST225", "Customer Name": "Customer 225",
        "Niche/Tablet Price (RM)": 10000,
        "Full Settlement Paid Date": day1 - datetime.timedelta(days=6),
        "Agency Code": "AC001",
    }])
    result1 = process_upload(db_path, str(xlsx1), run_date=day1)
    confirm_all_pending(db_path, result1["commission_run_id"])

    day2 = datetime.date.today()
    xlsx2 = tmp_path / "run2.xlsx"
    build_master_report(xlsx2, [
        {
            "No": 1, "PO No": 60025, "Customer ID": "CUST225", "Customer Name": "Customer 225",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": day1 - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60026, "Customer ID": "CUST226", "Customer Name": "Customer 226",
            "Niche/Tablet Price (RM)": 20000,
            "Full Settlement Paid Date": day2 - datetime.timedelta(days=6),
            "Agency Code": "AC001",
        },
    ])
    result2 = process_upload(db_path, str(xlsx2), run_date=day2)

    report_path = tmp_path / "report2.xlsx"
    confirm_all_pending(db_path, result2["commission_run_id"])
    generate_report(db_path, result2["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["All"]
    summary_cells = [row[0] for row in sheet.iter_rows(min_col=1, max_col=1)
                      if row[0].value and isinstance(row[0].value, str)
                      and (row[0].value.startswith("As at") or row[0].value.startswith("Total Sum"))]

    assert len(summary_cells) == 3  # 2 "As at" rows + 1 grand total line
    run1_cell, run2_cell, total_cell = summary_cells

    assert run1_cell.fill.start_color.rgb in ("00000000", None)  # older row: untouched
    assert run2_cell.fill.start_color.rgb in ("00FFFF00", "FFFFFF00")  # this run: yellow
    assert total_cell.fill.start_color.rgb in ("00FFFF00", "FFFFFF00")  # grand total: yellow too


def _summary_grand_total(sheet):
    for row in sheet.iter_rows(values_only=True):
        if row[0] and isinstance(row[0], str) and row[0].startswith("Total Sum"):
            return row[4]
    return None


def test_every_sheet_gets_its_own_date_record_scoped_to_itself(tmp_path):
    """
    Confirmed against the real file: XEMP's own sheet has a Date Record
    table totaling RM33,070.50 - a genuine subset of the "All" sheet's
    combined total, not the same table repeated everywhere. Each
    agency's sheet must show only its own commissions, not bleed in
    another agency's numbers.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60028, "Customer ID": "CUST228", "Customer Name": "Customer 228",
            "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60029, "Customer ID": "CUST229", "Customer Name": "Customer 229",
            "Niche/Tablet Price (RM)": 20000,  # 15% = 3000.00
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC210", "FCC/Agent": "Agent Beta",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    confirm_all_pending(db_path, result["commission_run_id"])

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert _summary_grand_total(workbook["All"]) == 4500.0       # 1500 + 3000, everyone
    assert _summary_grand_total(workbook["AC001"]) == 1500.0     # only its own PO
    assert _summary_grand_total(workbook["AC210"]) == 3000.0     # only its own PO
    # The per-agent sheet under AC210 is scoped the same way as its
    # group sheet, since there's only one agent in this test.
    assert _summary_grand_total(workbook["Agent Beta"]) == 3000.0


def test_no_agency_sheet_gets_its_own_date_record_too(tmp_path):
    """
    Regression test: a contract with no Agency Code at all lands on a
    "(No Agency)" sheet whose agency_group is also the literal string
    "(No Agency)" (see _load_master_rows's fallback) - not NULL. The
    Date Record query's filter has to fall back to that same literal
    string too, or COALESCE(NULL, NULL) never matches it and that
    sheet's own summary table comes up empty even though its main
    ledger table correctly shows the confirmed commission.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60036, "Customer ID": "CUST236", "Customer Name": "Customer 236",
        "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
        "Full Settlement Paid Date": settlement_date,
        # no Agency Code at all
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    confirm_all_pending(db_path, result["commission_run_id"])

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert "(No Agency)" in workbook.sheetnames
    assert _summary_grand_total(workbook["(No Agency)"]) == 1500.0


def test_long_agency_codes_that_collide_after_truncation_get_distinct_sheets(tmp_path):
    """
    Excel sheet names cap at 31 chars. Two different agency codes that
    both sanitize down to the same 31-char prefix must still end up as
    two distinct sheets, not hang forever trying to disambiguate them.
    """
    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)
    long_code_a = "A" * 31
    long_code_b = "A" * 31 + "-DIFFERENT-SUFFIX"  # same first 31 chars as long_code_a

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60013, "Customer ID": "CUST213", "Customer Name": "Customer 213",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": long_code_a, "FCC/Agent": "Agent LongA",
        },
        {
            "No": 2, "PO No": 60014, "Customer ID": "CUST214", "Customer Name": "Customer 214",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": long_code_b, "FCC/Agent": "Agent LongB",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    confirm_all_pending(db_path, result["commission_run_id"])
    generate_report(db_path, result["commission_run_id"], str(report_path))  # must not hang

    workbook = openpyxl.load_workbook(report_path)
    # "All" + two distinct (truncated/suffixed) agency group sheets +
    # two distinct agent sheets - every name unique, nothing silently
    # overwritten or looped forever trying to disambiguate.
    assert len(workbook.sheetnames) == len(set(workbook.sheetnames)) == 5


def test_generate_report_raises_clear_error_when_nothing_was_due(tmp_path):
    today = datetime.date.today()

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60015, "Customer ID": "CUST215", "Customer Name": "Customer 215",
        # nothing paid yet - no commission_run gets created
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    assert result["commission_run_id"] is None

    with pytest.raises(ValueError):
        generate_report(db_path, result["commission_run_id"], str(tmp_path / "report.xlsx"))
