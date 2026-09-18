"""
Builds fake Commission Base Report Excel files shaped exactly like the
real Kenjin export (title rows 1-5, headers on row 6, data from row 7)
so tests exercise the real importer code path, not a shortcut.
"""

import openpyxl

from app.pipeline import confirm_events, load_review

HEADERS = [
    "No", "PO No", "PO Date", "Signature Date", "Customer ID",
    "Customer Name", "Lot No", "Niche/Tablet Price (RM)",
    "Promotion (RM)", "Discount (RM)", "Nett Price (RM)",
    "Cooling Off Period", "Full Settlement Paid Date",
    "Full Payment Commission (RM)", "Full Commission Paid Date",
    "First Instalment Paid Date", "1st Half Commission (RM)",
    "1st Half Commission Paid Date", "Sixth Instalment Paid Date",
    "Balance Half Commission (RM)", "Balance Half Commission Paid Date",
    "FCC/Agent", "Agency Code", "Remarks",
]

# Sensible defaults so a test only needs to specify the fields it
# actually cares about.
_ROW_DEFAULTS = {
    "PO Date": None,
    "Signature Date": None,
    "Lot No": "L00-TEST-0000-00",
    "Niche/Tablet Price (RM)": 10000,
    "Promotion (RM)": 0,
    "Discount (RM)": 0,
    "Cooling Off Period": None,
    "Full Settlement Paid Date": None,
    "Full Payment Commission (RM)": None,
    "Full Commission Paid Date": None,
    "First Instalment Paid Date": None,
    "1st Half Commission (RM)": None,
    "1st Half Commission Paid Date": None,
    "Sixth Instalment Paid Date": None,
    "Balance Half Commission (RM)": None,
    "Balance Half Commission Paid Date": None,
    "FCC/Agent": "Test Agent",
    "Agency Code": None,
    "Remarks": None,
}


def build_master_report(path, rows, historical_summary_rows=None):
    """
    rows: list of dicts, each keyed by column header (only the keys
    you care about - everything else gets a sensible default). Each
    dict must at least include "No", "PO No", and "Customer ID".

    historical_summary_rows: optional list of dicts with keys
    "date_record" (a datetime.date), "full_commission",
    "first_half_commission", "second_half_commission" - writes the
    sheet's own trailing "Date Record" table (starting a few columns
    right of the PO data, matching the real file's layout) so tests
    can exercise app.importer's historical-import path.
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "XEKL Master - TEST"

    sheet.cell(row=2, column=1, value="XEKL OVERALL COMMISSION PAYOUT AS AT TEST DATE")
    # Rows 3-5 intentionally blank, matching the real file's layout.

    for col, header in enumerate(HEADERS, start=1):
        sheet.cell(row=6, column=col, value=header)

    for i, row in enumerate(rows):
        full_row = {**_ROW_DEFAULTS, **row}
        for col, header in enumerate(HEADERS, start=1):
            sheet.cell(row=7 + i, column=col, value=full_row.get(header))

    if historical_summary_rows:
        header_row = 7 + len(rows) + 3  # a few blank rows after "Total", matching the real layout
        header_col = len(HEADERS) - 3   # a few columns right of the PO data, matching the real layout
        labels = ["DATE RECORD", "Full Commission", "First Half Commission", "Second Half Commission", "Running Total", "Remarks"]
        for col, label in enumerate(labels, start=header_col):
            sheet.cell(row=header_row, column=col, value=label)
        for i, h in enumerate(historical_summary_rows):
            row_num = header_row + 1 + i
            sheet.cell(row=row_num, column=header_col, value=f"As at {h['date_record'].strftime('%d/%m/%Y')}")
            sheet.cell(row=row_num, column=header_col + 1, value=h["full_commission"])
            sheet.cell(row=row_num, column=header_col + 2, value=h["first_half_commission"])
            sheet.cell(row=row_num, column=header_col + 3, value=h["second_half_commission"])

    workbook.save(path)


AOR_HEADERS = [
    "No", "Acknowledgment Receipt No", "OR Receipt", "Acknowledgment Receipt Date",
    "Purchase Statement No", "Purchase Statement Date", "PO No", "Lot No",
    "Customer ID", "Customer Name", "Payor Name", "Payment Mode",
    "Instalment Plan Status", "Reference No", "Payment Received (RM)",
    "Created By", "Date Created",
]

_AOR_ROW_DEFAULTS = {
    "OR Receipt": None,
    "Purchase Statement No": "INS-TEST-00001",
    "Purchase Statement Date": None,
    "Lot No": "L00-TEST-0000-00",
    "Payor Name": "Test Payor",
    "Payment Mode": "Cheque",
    "Instalment Plan Status": None,
    "Payment Received (RM)": 500,
    "Created By": "test",
    "Date Created": None,
}


def build_aor_report(path, rows):
    """
    rows: list of dicts, each keyed by an AOR_HEADERS column (only the
    keys you care about - everything else gets a sensible default).
    Each dict must at least include "No", "Acknowledgment Receipt No",
    "PO No", "Acknowledgment Receipt Date", and "Reference No".

    Shaped like the real AOR export: a title block, headers a few rows
    down, data below - close enough to the real layout for
    app.aor._is_aor_shaped/_read_aor_rows to read it the same way they
    read a real file.
    """
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "fin_dmy_collect_report-TEST"

    sheet.cell(row=5, column=1, value="List of Acknowledgment of Receipt Report")

    for col, header in enumerate(AOR_HEADERS, start=1):
        sheet.cell(row=22, column=col, value=header)

    for i, row in enumerate(rows):
        full_row = {**_AOR_ROW_DEFAULTS, **row}
        for col, header in enumerate(AOR_HEADERS, start=1):
            sheet.cell(row=23 + i, column=col, value=full_row.get(header))

    workbook.save(path)


def confirm_all_pending(db_path, commission_run_id):
    """
    Test convenience: confirms every pending event on a commission run,
    the way a human clicking through the review page would one at a
    time. Most existing tests were written before the review-and-
    confirm step existed and just want "everything detected this run
    is now due" - this is that, in one call, so they can go straight
    from process_upload to generate_report/download the same way they
    always did.
    """
    events = load_review(db_path, commission_run_id)
    pending_ids = {e["id"] for e in events if e["status"] == "pending"}
    confirm_events(db_path, commission_run_id, pending_ids, "test-user")
