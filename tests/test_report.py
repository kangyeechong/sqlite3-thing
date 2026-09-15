"""
Step 3 verification: the Excel report handed to Accounts.

Run with: pytest tests/test_report.py -v
"""

import datetime

import openpyxl
import pytest

from app.pipeline import process_upload, generate_report
from tests.helpers import build_master_report


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def _sheet_rows(sheet):
    """Every non-empty row in a sheet, as tuples - for easy assertions."""
    return [tuple(r) for r in sheet.iter_rows(values_only=True) if any(v is not None for v in r)]


def test_report_has_an_all_sheet_and_a_sheet_per_agency(tmp_path):
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60001, "Customer ID": "CUST201", "Customer Name": "Customer 201",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",  # no-split agency
        },
        {
            "No": 2, "PO No": 60002, "Customer ID": "CUST202", "Customer Name": "Customer 202",
            "Niche/Tablet Price (RM)": 20000,
            "First Instalment Paid Date": today,
            "Agency Code": "AC200",  # splits by agent (default)
            "FCC/Agent": "Agent Alpha",
        },
    ])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)

    report_path = tmp_path / "report.xlsx"
    generate_report(db_path, result["commission_run_id"], str(report_path))

    workbook = openpyxl.load_workbook(report_path)
    assert set(workbook.sheetnames) == {"All", "AC001", "AC200"}

    all_rows = _sheet_rows(workbook["All"])
    po_numbers_in_all = {row[0] for row in all_rows if isinstance(row[0], int)}
    assert po_numbers_in_all == {60001, 60002}


def test_no_split_agency_sheet_is_one_flat_table(tmp_path):
    """AC001 (XEMP) must not be broken down by agent - one flat table."""
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60003, "Customer ID": "CUST203", "Customer Name": "Customer 203",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001", "FCC/Agent": "Staff A",
        },
        {
            "No": 2, "PO No": 60004, "Customer ID": "CUST204", "Customer Name": "Customer 204",
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
    ac001_rows = _sheet_rows(workbook["AC001"])
    # Exactly one title row ("AC001 - Commission Due"), not two agent-titled sections.
    title_rows = [r for r in ac001_rows if r[0] and isinstance(r[0], str) and "Commission Due" in r[0]]
    assert len(title_rows) == 1


def test_splitting_agency_sheet_has_one_section_per_agent(tmp_path):
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60005, "Customer ID": "CUST205", "Customer Name": "Customer 205",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC210", "FCC/Agent": "Agent One",
        },
        {
            "No": 2, "PO No": 60006, "Customer ID": "CUST206", "Customer Name": "Customer 206",
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
    ac210_rows = _sheet_rows(workbook["AC210"])
    agent_titles = {r[0] for r in ac210_rows if r[0] in ("Agent One", "Agent Two")}
    assert agent_titles == {"Agent One", "Agent Two"}


def test_report_total_matches_sum_of_amounts(tmp_path):
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60007, "Customer ID": "CUST207", "Customer Name": "Customer 207",
            "Niche/Tablet Price (RM)": 10000,  # 15% = 1500.00
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 60008, "Customer ID": "CUST208", "Customer Name": "Customer 208",
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
    all_rows = _sheet_rows(workbook["All"])
    total_row = next(r for r in all_rows if r[-2] == "Total")
    assert total_row[-1] == 4500.0


def test_long_agency_codes_that_collide_after_truncation_get_distinct_sheets(tmp_path):
    """
    Excel sheet names cap at 31 chars. Two different agency codes that
    both sanitize down to the same 31-char prefix must still end up as
    two distinct sheets, not hang forever trying to disambiguate them.
    """
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)
    long_code_a = "A" * 31
    long_code_b = "A" * 31 + "-DIFFERENT-SUFFIX"  # same first 31 chars as long_code_a

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 60010, "Customer ID": "CUST210", "Customer Name": "Customer 210",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": long_code_a,
        },
        {
            "No": 2, "PO No": 60011, "Customer ID": "CUST211", "Customer Name": "Customer 211",
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
    today = datetime.date(2026, 8, 20)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 60009, "Customer ID": "CUST209", "Customer Name": "Customer 209",
        # nothing paid yet - no commission_run gets created
    }])

    db_path = _db_path(tmp_path)
    result = process_upload(db_path, str(xlsx_path), run_date=today)
    assert result["commission_run_id"] is None

    with pytest.raises(ValueError):
        generate_report(db_path, result["commission_run_id"], str(tmp_path / "report.xlsx"))
