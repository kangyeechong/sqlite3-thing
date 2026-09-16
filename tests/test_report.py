"""
Step 3 verification: the Excel report handed to Accounts, in the real
Master Report column layout.

Run with: pytest tests/test_report.py -v
"""

import datetime

import openpyxl
import pytest

from app.pipeline import process_upload, generate_report
from tests.helpers import build_master_report


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
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert set(workbook.sheetnames) == {"All", "AC001", "AC200"}

    _, all_rows = _find_table_rows(workbook["All"])
    po_numbers = {r["PO No"] for r in all_rows}
    assert po_numbers == {60001, 60002}


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


def test_newly_due_commission_cells_are_highlighted_yellow(tmp_path):
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
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    sheet = workbook["All"]
    headers, data_rows = _find_table_rows(sheet)
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    data_row_num = header_row_num + 1
    commission_col = headers.index("Full Payment Commission (RM)") + 1
    other_col = headers.index("Customer Name") + 1

    assert sheet.cell(row=data_row_num, column=commission_col).fill.start_color.rgb in ("00FFFF00", "FFFFFF00")
    assert sheet.cell(row=data_row_num, column=other_col).fill.start_color.rgb in ("00000000", None)


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
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, ac001_rows = _find_table_rows(workbook["AC001"])
    assert len(ac001_rows) == 2  # one flat table, both agents mixed together


def test_splitting_agency_sheet_has_one_section_per_agent(tmp_path):
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
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    ac210_rows = [tuple(r) for r in workbook["AC210"].iter_rows(values_only=True) if any(v is not None for v in r)]
    agent_titles = {r[0] for r in ac210_rows if r[0] in ("Agent One", "Agent Two")}
    assert agent_titles == {"Agent One", "Agent Two"}


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
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    _, all_rows = _find_table_rows(workbook["All"])
    # Both rows are Full Payment triggers, so the subtotal lands under
    # that specific column - Nett Price ("Total" label) and the other
    # two commission columns stay at their per-row values (not summed
    # across rows in this dict form), so check via the raw cell values.
    sheet = workbook["All"]
    header_row_num = next(row[0].row for row in sheet.iter_rows() if any(c.value == "PO No" for c in row))
    total_row_num = header_row_num + len(all_rows) + 1
    headers = [cell.value for cell in sheet[header_row_num]]
    total_row_values = dict(zip(headers, [cell.value for cell in sheet[total_row_num]]))
    assert total_row_values["Nett Price (RM)"] == "Total"
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
    generate_report(db_path, result2["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    all_values = [tuple(r) for r in workbook["All"].iter_rows(values_only=True) if any(v is not None for v in r)]
    summary_rows = [r for r in all_values if isinstance(r[0], str) and r[0].startswith("As at")]

    assert len(summary_rows) == 2
    assert summary_rows[0][4] == 1500.0   # running total after run 1
    assert summary_rows[1][4] == 4500.0   # running total after run 2 (1500 + 3000)

    grand_total_row = next(r for r in all_values if isinstance(r[0], str) and r[0].startswith("Total Sum of Commission Payout"))
    assert grand_total_row[4] == 4500.0


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
            "Agency Code": long_code_a,
        },
        {
            "No": 2, "PO No": 60014, "Customer ID": "CUST214", "Customer Name": "Customer 214",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": long_code_b,
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result["commission_run_id"], str(report_path))  # must not hang

    workbook = openpyxl.load_workbook(report_path)
    assert len(workbook.sheetnames) == 3  # "All" + two distinct agency sheets


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
