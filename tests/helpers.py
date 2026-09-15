"""
Builds fake Commission Base Report Excel files shaped exactly like the
real Kenjin export (title rows 1-5, headers on row 6, data from row 7)
so tests exercise the real importer code path, not a shortcut.
"""

import openpyxl

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


def build_master_report(path, rows):
    """
    rows: list of dicts, each keyed by column header (only the keys
    you care about - everything else gets a sensible default). Each
    dict must at least include "No", "PO No", and "Customer ID".
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

    workbook.save(path)
