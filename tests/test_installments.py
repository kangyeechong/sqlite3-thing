"""
Step 2 verification: installment tracking (7.5% at installment 1,
7.5% at installment 6), across multiple upload cycles.

Run with: pytest tests/test_installments.py -v
"""

import datetime

from app.pipeline import process_upload
from tests.helpers import build_master_report


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def test_installment_1_paid_flags_first_half(tmp_path):
    today = datetime.date(2026, 8, 20)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 70001, "Customer ID": "CUST101", "Customer Name": "Customer 101",
        "Niche/Tablet Price (RM)": 20000, "Discount (RM)": 400,
        "First Instalment Paid Date": today,
        "Agency Code": "AC001",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert len(result["raised_events"]) == 1
    event = result["raised_events"][0]
    assert event["po_no"] == 70001
    assert event["trigger_type"] == "installment_1"
    # Net Price = 20000 - 400 = 19600; 7.5% = 1470.00
    assert event["amount"] == 1470.0


def test_installment_progresses_across_cycles_and_carries_forward_when_nothing_changes(tmp_path):
    """
    The exact scenario from the original brief: a contract progressing
    through installment 1 and installment 6 across separate upload
    cycles, including a cycle in between where nothing new happened -
    which must carry forward silently, not error or re-flag anything.
    """
    db_path = _db_path(tmp_path)
    po_no = 70002
    net_price_inputs = {"Niche/Tablet Price (RM)": 24000, "Discount (RM)": 0}  # Net Price = 24000

    # Cycle 1: contract exists, nothing paid yet.
    cycle1 = tmp_path / "cycle1.xlsx"
    build_master_report(cycle1, [{
        "No": 1, "PO No": po_no, "Customer ID": "CUST102", "Customer Name": "Customer 102",
        **net_price_inputs, "Agency Code": "AC001",
    }])
    result1 = process_upload(db_path, str(cycle1), run_date=datetime.date(2026, 6, 5))
    assert result1["raised_events"] == []

    # Cycle 2: installment 1 just got paid.
    cycle2 = tmp_path / "cycle2.xlsx"
    build_master_report(cycle2, [{
        "No": 1, "PO No": po_no, "Customer ID": "CUST102", "Customer Name": "Customer 102",
        **net_price_inputs, "Agency Code": "AC001",
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
    }])
    result2 = process_upload(db_path, str(cycle2), run_date=datetime.date(2026, 6, 22))
    assert len(result2["raised_events"]) == 1
    assert result2["raised_events"][0]["trigger_type"] == "installment_1"
    assert result2["raised_events"][0]["amount"] == 1800.0  # 24000 * 7.5%

    # Cycle 3: nothing new - same file re-exported, installment 1 still
    # the only thing paid. Must carry forward silently: no events, no
    # errors, nothing re-flagged.
    cycle3 = tmp_path / "cycle3.xlsx"
    build_master_report(cycle3, [{
        "No": 1, "PO No": po_no, "Customer ID": "CUST102", "Customer Name": "Customer 102",
        **net_price_inputs, "Agency Code": "AC001",
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
    }])
    result3 = process_upload(db_path, str(cycle3), run_date=datetime.date(2026, 7, 6))
    assert result3["raised_events"] == []

    # Cycle 4: installment 6 finally paid, months later.
    cycle4 = tmp_path / "cycle4.xlsx"
    build_master_report(cycle4, [{
        "No": 1, "PO No": po_no, "Customer ID": "CUST102", "Customer Name": "Customer 102",
        **net_price_inputs, "Agency Code": "AC001",
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
        "Sixth Instalment Paid Date": datetime.date(2026, 11, 18),
    }])
    result4 = process_upload(db_path, str(cycle4), run_date=datetime.date(2026, 11, 20))
    assert len(result4["raised_events"]) == 1
    assert result4["raised_events"][0]["trigger_type"] == "installment_6"
    assert result4["raised_events"][0]["amount"] == 1800.0

    # Cycle 5: nothing new again - re-uploading the same completed
    # state must not re-raise either installment.
    cycle5 = tmp_path / "cycle5.xlsx"
    build_master_report(cycle5, [{
        "No": 1, "PO No": po_no, "Customer ID": "CUST102", "Customer Name": "Customer 102",
        **net_price_inputs, "Agency Code": "AC001",
        "First Instalment Paid Date": datetime.date(2026, 6, 20),
        "Sixth Instalment Paid Date": datetime.date(2026, 11, 18),
    }])
    result5 = process_upload(db_path, str(cycle5), run_date=datetime.date(2026, 12, 6))
    assert result5["raised_events"] == []


def test_both_installments_already_paid_on_first_ever_import(tmp_path):
    """
    Importing a plan that's already complete (both installment 1 and
    installment 6 paid) for the first time must raise BOTH events in
    the same run, not just one.
    """
    today = datetime.date(2026, 8, 20)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 70003, "Customer ID": "CUST103", "Customer Name": "Customer 103",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": datetime.date(2026, 1, 5),
        "Sixth Instalment Paid Date": datetime.date(2026, 7, 5),
        "Agency Code": "AC001",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    trigger_types = {e["trigger_type"] for e in result["raised_events"]}
    assert trigger_types == {"installment_1", "installment_6"}


def test_installment_commission_blocked_by_non_positive_net_price(tmp_path):
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 70004, "Customer ID": "CUST104", "Customer Name": "Customer 104",
        "Niche/Tablet Price (RM)": 1000, "Discount (RM)": 2000,  # net_price negative
        "First Instalment Paid Date": datetime.date(2026, 8, 1),
        "Agency Code": "AC001",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    assert result["raised_events"] == []


def test_installment_commission_blocked_on_cancelled_po(tmp_path):
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 70005, "Customer ID": "CUST105", "Customer Name": "Customer 105",
        "Niche/Tablet Price (RM)": 10000,
        "First Instalment Paid Date": datetime.date(2026, 8, 1),
        "Remarks": "Cancelled PO",
    }])

    result = process_upload(_db_path(tmp_path), str(xlsx_path))

    assert result["raised_events"] == []


def test_full_payment_and_installment_can_coexist_in_the_same_run(tmp_path):
    """
    A single upload can contain one contract newly due for full
    payment and a different contract newly due for installment 1 - both
    must show up together in the same commission_run.
    """
    today = datetime.date(2026, 8, 20)
    settlement_date = today - datetime.timedelta(days=6)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [
        {
            "No": 1, "PO No": 70006, "Customer ID": "CUST106", "Customer Name": "Customer 106",
            "Niche/Tablet Price (RM)": 10000,
            "Full Settlement Paid Date": settlement_date,
            "Agency Code": "AC001",
        },
        {
            "No": 2, "PO No": 70007, "Customer ID": "CUST107", "Customer Name": "Customer 107",
            "Niche/Tablet Price (RM)": 10000,
            "First Instalment Paid Date": today,
            "Agency Code": "AC001",
        },
    ])

    result = process_upload(_db_path(tmp_path), str(xlsx_path), run_date=today)

    assert result["commission_run_id"] is not None
    trigger_by_po = {e["po_no"]: e["trigger_type"] for e in result["raised_events"]}
    assert trigger_by_po == {70006: "full_payment", 70007: "installment_1"}
