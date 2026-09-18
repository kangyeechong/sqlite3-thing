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
