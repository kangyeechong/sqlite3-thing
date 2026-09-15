"""
Commission calculation logic.

STEP 1 SCOPE: only full-payment commission is implemented here.
Installment commission (7.5% at installment 1, 7.5% at installment 6)
is Step 2 and deliberately not built yet - see the project's phased
build plan. Calling process_full_payment_run() will not touch
installment-plan contracts at all.
"""

import datetime
from decimal import ROUND_HALF_UP, Decimal

from . import rules


def full_payment_is_due(contract_row, as_of):
    """
    contract_row: a sqlite3.Row (or dict) from the `contracts` table.
    as_of: a datetime.date - "today" for the purposes of the cooling-off
    check. Passed in explicitly (never computed as datetime.date.today()
    inside this function) so tests can control it precisely.

    Returns True if this contract's full-payment commission should be
    raised as newly due right now.
    """
    if contract_row["status"] != "active":
        return False
    if contract_row["full_commission_flagged"]:
        return False
    # A non-positive Net Price means the source data is wrong (e.g. a
    # discount larger than the price) - never calculate a commission
    # from it. Deliberately NOT setting full_commission_flagged here:
    # once staff fix the underlying data and re-upload, this contract
    # must still be eligible to be flagged correctly. The import review
    # panel's "non_positive_net_price" check is what surfaces this to a
    # human; this is the hard stop that keeps it from being paid out.
    net_price = contract_row["net_price"]
    if net_price is None or net_price <= 0:
        return False

    settlement_date = contract_row["full_settlement_paid_date"]
    if not settlement_date:
        return False
    settlement_date = datetime.date.fromisoformat(settlement_date)

    if contract_row["case_type"] == "at_need":
        # No cooling-off wait for At-Need, but an inurnment date must
        # be on file first.
        return bool(contract_row["inurnment_date"])

    gate_clears_on = settlement_date + datetime.timedelta(
        days=rules.COOLING_OFF_DAYS_BEFORE_RELEASE
    )
    return as_of >= gate_clears_on


def calculate_full_payment_commission(net_price):
    """
    Uses Decimal with explicit half-up rounding rather than Python's
    built-in round() (which does banker's rounding on binary floats)
    so a commission landing exactly on a half-cent boundary rounds the
    way a manual calculation - and an accountant - would expect.
    """
    amount = Decimal(str(net_price)) * Decimal(str(rules.FULL_PAYMENT_COMMISSION_PCT))
    return float(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def process_full_payment_run(conn, as_of, run_date, source_filename, created_by_user):
    """
    Scans every active contract for newly-due full-payment commission,
    logs a commission_event for each, marks the contract's flag so it's
    never raised again, and groups everything under one new
    commission_run row.

    Returns (commission_run_id, list of dicts describing what was
    raised) - an empty list is a normal, expected outcome (nothing
    newly crossed since the last run), not an error.
    """
    cursor = conn.execute("SELECT * FROM contracts WHERE status = 'active'")
    candidates = cursor.fetchall()

    raised = []
    for contract in candidates:
        if not full_payment_is_due(contract, as_of):
            continue
        amount = calculate_full_payment_commission(contract["net_price"])
        raised.append({
            "po_no": contract["po_no"],
            "trigger_type": "full_payment",
            "trigger_date": contract["full_settlement_paid_date"],
            "amount": amount,
        })

    if not raised:
        return None, []

    now_iso = datetime.datetime.now().isoformat()
    run_cursor = conn.execute(
        "INSERT INTO commission_runs (run_date, source_filename, created_by_user) "
        "VALUES (?, ?, ?)",
        (run_date.isoformat(), source_filename, created_by_user),
    )
    run_id = run_cursor.lastrowid

    for event in raised:
        conn.execute(
            """
            INSERT INTO commission_events (
                po_no, trigger_type, trigger_date, amount,
                detected_at, detected_by_user, commission_run_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                event["po_no"], event["trigger_type"], event["trigger_date"],
                event["amount"], now_iso, created_by_user, run_id,
            ),
        )
        conn.execute(
            "UPDATE contracts SET full_commission_flagged = 1 WHERE po_no = ?",
            (event["po_no"],),
        )

    return run_id, raised
