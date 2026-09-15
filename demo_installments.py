"""
Plain-English demo of Step 2 (installment tracking), walking one fake
contract through several upload cycles - the exact scenario from the
original brief: progressing through installments, a cycle where
nothing changes, and installment 6 arriving much later.

Run with: python3 demo_installments.py
"""

import datetime
import shutil
from pathlib import Path

from app.pipeline import process_upload
from tests.helpers import build_master_report

DEMO_DIR = Path(__file__).parent / "_demo_run"
DB_PATH = DEMO_DIR / "demo_installments.db"

PO_NO = 85001
NET_PRICE_INPUTS = {"Niche/Tablet Price (RM)": 24000, "Discount (RM)": 0}  # Net Price = 24000


def upload_and_report(cycle_label, run_date, extra_fields):
    upload_path = DEMO_DIR / f"{cycle_label.replace(' ', '_')}.xlsx"
    build_master_report(upload_path, [{
        "No": 1, "PO No": PO_NO, "Customer ID": "DEMO201", "Customer Name": "Goh Bee Choo (fake)",
        **NET_PRICE_INPUTS, "Agency Code": "AC001",
        **extra_fields,
    }])

    result = process_upload(str(DB_PATH), str(upload_path), run_date=run_date, created_by_user="demo")

    print(f"--- {cycle_label} (processed as-of {run_date}) ---")
    if result["raised_events"]:
        for event in result["raised_events"]:
            print(f"  -> COMMISSION DUE: RM{event['amount']:,.2f} ({event['trigger_type']})")
    else:
        print("  -> nothing due this cycle (carries forward silently)")
    print()


def main():
    if DEMO_DIR.exists():
        shutil.rmtree(DEMO_DIR)
    DEMO_DIR.mkdir()

    print("=" * 72)
    print("One fake contract (PO 85001, Net Price RM24,000), followed across")
    print("five upload cycles - same file structure re-uploaded each time,")
    print("with more columns filled in as real payments would arrive.")
    print("=" * 72)
    print()

    upload_and_report("Cycle 1 - just signed, nothing paid", datetime.date(2026, 6, 5), {})

    upload_and_report("Cycle 2 - installment 1 just paid", datetime.date(2026, 6, 22), {
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
    })

    upload_and_report("Cycle 3 - nothing new happened", datetime.date(2026, 7, 6), {
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
    })

    upload_and_report("Cycle 4 - installment 6 paid, months later", datetime.date(2026, 11, 20), {
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
        "Sixth Instalment Paid Date": datetime.date(2026, 11, 18),
    })

    upload_and_report("Cycle 5 - re-uploaded again, nothing new", datetime.date(2026, 12, 6), {
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
        "Sixth Instalment Paid Date": datetime.date(2026, 11, 18),
    })

    print("Expected result, for comparison:")
    print("  Cycle 1 -> nothing")
    print("  Cycle 2 -> DUE, RM1,800.00 (installment_1, 7.5% of 24,000)")
    print("  Cycle 3 -> nothing (already flagged, carries forward silently)")
    print("  Cycle 4 -> DUE, RM1,800.00 (installment_6, 7.5% of 24,000)")
    print("  Cycle 5 -> nothing (both already flagged)")
    print()
    print(f"Total commissioned across the whole plan: RM3,600.00 (=15% of 24,000,")
    print("split across two dates, months apart, across five separate uploads)")


if __name__ == "__main__":
    main()
