"""
Plain-English demo of Step 1 (full-payment flagging) using fake
contracts, meant to be read, not debugged. Run it with:

    python3 demo.py

No pytest, no test-runner jargon in the output - just a handful of
made-up contracts and what the system decided about each one, so you
can sanity-check the logic yourself without reading any code.
"""

import datetime
import shutil
from pathlib import Path

from app.pipeline import process_upload
from tests.helpers import build_master_report

DEMO_DIR = Path(__file__).parent / "_demo_run"
DB_PATH = DEMO_DIR / "demo_ledger.db"
UPLOAD_PATH = DEMO_DIR / "demo_upload.xlsx"

TODAY = datetime.date(2026, 8, 20)

# Each of these is a made-up contract, not a real customer. Comments
# say what we EXPECT the system to decide, so you can check the actual
# result against that expectation.
FAKE_CONTRACTS = [
    {
        # Paid in full 6 days ago -> past the 5-day gate -> SHOULD be flagged.
        "No": 1, "PO No": 80001, "Customer ID": "DEMO001", "Customer Name": "Tan Ah Kow (fake)",
        "Niche/Tablet Price (RM)": 20000, "Discount (RM)": 500,
        "Full Settlement Paid Date": TODAY - datetime.timedelta(days=6),
        "Agency Code": "AC001",
    },
    {
        # Paid in full only 2 days ago -> still inside the gate -> should NOT be flagged yet.
        "No": 2, "PO No": 80002, "Customer ID": "DEMO002", "Customer Name": "Lim Mei Ling (fake)",
        "Niche/Tablet Price (RM)": 15000,
        "Full Settlement Paid Date": TODAY - datetime.timedelta(days=2),
        "Agency Code": "AC001",
    },
    {
        # At-Need case, paid today, inurnment date on file -> no waiting -> SHOULD be flagged immediately.
        "No": 3, "PO No": 80003, "Customer ID": "DEMO003", "Customer Name": "Ong Siew Hua (fake)",
        "Niche/Tablet Price (RM)": 12000,
        "Full Settlement Paid Date": TODAY,
        "Remarks": "At need case\nInurnment on 20/08/2026",
        "Agency Code": "AC109",
    },
    {
        # Cancelled -> should be excluded no matter what the paid date says.
        "No": 4, "PO No": 80004, "Customer ID": "DEMO004", "Customer Name": "Wong Kar Wai (fake)",
        "Niche/Tablet Price (RM)": 18000,
        "Full Settlement Paid Date": TODAY - datetime.timedelta(days=30),
        "Remarks": "Cancelled PO",
    },
    {
        # Still on an installment plan, nothing paid yet -> nothing to do this cycle.
        "No": 5, "PO No": 80005, "Customer ID": "DEMO005", "Customer Name": "Chong Wei Ming (fake)",
        "Niche/Tablet Price (RM)": 24000,
        "Agency Code": "AC001",
    },
]


def main():
    if DEMO_DIR.exists():
        shutil.rmtree(DEMO_DIR)
    DEMO_DIR.mkdir()

    build_master_report(UPLOAD_PATH, FAKE_CONTRACTS)

    print("=" * 72)
    print(f"Uploading a fake Commission Base Report, processed as-of {TODAY}")
    print("=" * 72)

    result = process_upload(str(DB_PATH), str(UPLOAD_PATH), run_date=TODAY, created_by_user="demo")

    ir = result["import_result"]
    print(f"\nRead {ir.contracts_seen} contracts from the file.\n")

    raised_by_po = {e["po_no"]: e for e in result["raised_events"]}

    for c in FAKE_CONTRACTS:
        po_no = c["PO No"]
        name = c["Customer ID"] + " - " + c.get("Customer Name", "")
        if po_no in raised_by_po:
            event = raised_by_po[po_no]
            print(f"PO {po_no} ({name})")
            print(f"  -> COMMISSION DUE: RM{event['amount']:,.2f} (full payment)")
        else:
            print(f"PO {po_no} ({name})")
            print(f"  -> nothing due this cycle")
        print()

    if ir.review_flags:
        print("-" * 72)
        print("Import review panel flagged the following for a human to check:")
        for flag in ir.review_flags:
            print(f"  - PO {flag.po_no}: {flag.message}")
    else:
        print("Import review panel: nothing flagged.")

    print()
    print("Expected result, for comparison:")
    print("  PO 80001 (paid 6 days ago)      -> DUE, RM2,925.00 (15% of 19,500)")
    print("  PO 80002 (paid 2 days ago)       -> nothing yet (inside the 5-day gate)")
    print("  PO 80003 (At-Need, paid today)   -> DUE, RM1,800.00 (15% of 12,000)")
    print("  PO 80004 (cancelled)             -> nothing (excluded)")
    print("  PO 80005 (nothing paid yet)      -> nothing")


if __name__ == "__main__":
    main()
