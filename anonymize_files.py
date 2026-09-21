"""
Temporary utility: anonymizes customer and agent names across the Base
Report, AOR, and "overall commission" reference files so they can be
shared and tested with without exposing real customer/agent data.

Run this locally, on your own machine, against your own real files -
nothing gets uploaded anywhere. Every file you pass gets a new
anonymized copy written to --out-dir; your originals are never
modified.

The SAME real customer/agent gets the SAME fake replacement everywhere
it appears, across every file passed in this one run - so PO No,
dates, prices, agency codes, and Lot No all stay real and
cross-referenceable between files; only the people's names change.

Usage:
    python3 anonymize_files.py base_report.xlsx aor_report.xlsx overall_commission.xlsx --out-dir anonymized

Not part of the web app - safe to delete once you're done testing with
real data and don't need it anymore.
"""

import argparse
import os

import openpyxl
from openpyxl.cell.cell import MergedCell

# Header text (case-insensitive, exact match) that marks a column as
# holding a customer/agent name or ID worth anonymizing. If the
# "overall commission" reference file uses different wording for the
# same thing, add it to the matching set below.
_CUSTOMER_ID_HEADERS = {"customer id"}
_CUSTOMER_NAME_HEADERS = {"customer name"}
_PAYOR_NAME_HEADERS = {"payor name"}
_AGENT_NAME_HEADERS = {"fcc/agent", "agent", "agent name"}


class _FakeValuePool:
    """Hands out sequential fake values, remembering every original ->
    fake mapping so the same real value always gets the same fake one
    everywhere it appears across every file passed in this run."""

    def __init__(self, template):
        self._template = template  # e.g. "Test Customer {n}" or "CUSTTEST{n:04d}"
        self.mapping = {}

    def get(self, original):
        if original is None or str(original).strip() == "":
            return original
        key = str(original).strip()
        if key not in self.mapping:
            self.mapping[key] = self._template.format(n=len(self.mapping) + 1)
        return self.mapping[key]


def _matches(header, target_set):
    return header is not None and str(header).strip().lower() in target_set


def anonymize_workbook(path, pools):
    """
    pools: dict of field -> _FakeValuePool, shared across every file in
    this run so the same real value maps to the same fake one
    everywhere.

    Searches every sheet's first 30 rows for a header row containing
    any of the target headers (the same "search, don't assume a fixed
    row number" approach the rest of the tool uses, since title-block
    height varies between file types), then anonymizes every data row
    under it.
    """
    workbook = openpyxl.load_workbook(path)
    found_columns = []

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=30):
            header_hits = []
            for cell in row:
                if _matches(cell.value, _CUSTOMER_ID_HEADERS):
                    header_hits.append((cell.column, "customer_id"))
                elif _matches(cell.value, _CUSTOMER_NAME_HEADERS):
                    header_hits.append((cell.column, "customer_name"))
                elif _matches(cell.value, _PAYOR_NAME_HEADERS):
                    header_hits.append((cell.column, "payor_name"))
                elif _matches(cell.value, _AGENT_NAME_HEADERS):
                    header_hits.append((cell.column, "agent_name"))
            if not header_hits:
                continue

            header_row_num = row[0].row
            found_columns.append((sheet.title, header_row_num, header_hits))

            for data_row in sheet.iter_rows(min_row=header_row_num + 1):
                for col_idx, field in header_hits:
                    cell = sheet.cell(row=data_row[0].row, column=col_idx)
                    # A cell that's part of a merged range (other than
                    # its top-left anchor) has no independent value to
                    # set - real business spreadsheets merge cells for
                    # formatting all the time, so this is expected, not
                    # an error condition worth stopping the run over.
                    if isinstance(cell, MergedCell):
                        continue
                    cell.value = pools[field].get(cell.value)

    return workbook, found_columns


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("files", nargs="+", help="xlsx files to anonymize (Base Report, AOR, overall commission)")
    parser.add_argument("--out-dir", default="anonymized")
    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    pools = {
        "customer_id": _FakeValuePool("CUSTTEST{n:04d}"),
        "customer_name": _FakeValuePool("Test Customer {n}"),
        "payor_name": _FakeValuePool("Test Payor {n}"),
        "agent_name": _FakeValuePool("Test Agent {n}"),
    }

    for path in args.files:
        workbook, found_columns = anonymize_workbook(path, pools)
        out_path = os.path.join(args.out_dir, os.path.basename(path))
        workbook.save(out_path)

        print(f"\n{path} -> {out_path}")
        if not found_columns:
            print(
                "  WARNING: no recognizable Customer ID / Customer Name / Payor "
                "Name / FCC-Agent column found - nothing anonymized in this "
                "file. Check the header names against the sets at the top of "
                "this script and add yours if they differ."
            )
        for sheet_title, header_row_num, hits in found_columns:
            fields = ", ".join(field for _, field in hits)
            print(f"  sheet '{sheet_title}' row {header_row_num}: anonymized [{fields}]")

    print(
        f"\n{len(pools['customer_id'].mapping)} unique customer ID(s), "
        f"{len(pools['customer_name'].mapping)} unique customer name(s), "
        f"{len(pools['payor_name'].mapping)} unique payor name(s), "
        f"{len(pools['agent_name'].mapping)} unique agent name(s) anonymized "
        f"consistently across all {len(args.files)} file(s)."
    )


if __name__ == "__main__":
    main()
