"""
Commission calculation logic - determines what's newly due and logs it.

Covers all three Phase 1 triggers: full payment (gated by the
cooling-off wait), installment 1, and installment 6. See
docs/data_model.md section 4 for the state machine these implement.
"""

import datetime
from decimal import ROUND_HALF_UP, Decimal

from . import rules


def _has_valid_net_price(contract_row):
    net_price = contract_row["net_price"]
    return net_price is not None and net_price > 0


def _calculate_commission(net_price, pct):
    """
    Shared by every trigger type. Uses Decimal with explicit half-up
    rounding rather than Python's built-in round() (which does
    banker's rounding on binary floats) so a commission landing
    exactly on a half-cent boundary rounds the way a manual
    calculation - and an accountant - would expect.
    """
    amount = Decimal(str(net_price)) * Decimal(str(pct))
    return float(amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


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
    # from it. Deliberately NOT setting the flag here: once staff fix
    # the underlying data and re-upload, this contract must still be
    # eligible to be flagged correctly. The import review panel's
    # "non_positive_net_price" check is what surfaces this to a human.
    if not _has_valid_net_price(contract_row):
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


def installment_1_is_due(contract_row):
    """
    True if the first 7.5% should be raised as newly due. Unlike full
    payment, there is no cooling-off gate on installments - the written
    commission rules only apply that condition to one-off/full payment.
    """
    if contract_row["status"] != "active":
        return False
    if contract_row["installment_1_commission_flagged"]:
        return False
    if not _has_valid_net_price(contract_row):
        return False
    return bool(contract_row["first_installment_paid_date"])


def installment_6_is_due(contract_row):
    """
    True if the second 7.5% should be raised as newly due. This is
    checked independently of installment_1_is_due - a contract's very
    first import could already have both installment 1 and installment
    6 paid (e.g. importing a plan that's already mid-way through), and
    both must be raised then, not just one.
    """
    if contract_row["status"] != "active":
        return False
    if contract_row["installment_6_commission_flagged"]:
        return False
    if not _has_valid_net_price(contract_row):
        return False
    return bool(contract_row["sixth_installment_paid_date"])


def calculate_full_payment_commission(net_price):
    return _calculate_commission(net_price, rules.FULL_PAYMENT_COMMISSION_PCT)


def calculate_installment_1_commission(net_price):
    return _calculate_commission(net_price, rules.INSTALLMENT_1_COMMISSION_PCT)


def calculate_installment_6_commission(net_price):
    return _calculate_commission(net_price, rules.INSTALLMENT_6_COMMISSION_PCT)


# Fixed, hardcoded SQL per trigger type - deliberately not built by
# string-formatting a column name into a query, even though the only
# input is our own code's trigger_type, not user data. Keeping every
# SQL statement a plain literal is one less thing to have to reason
# about later.
_FLAG_COLUMN_UPDATE_SQL = {
    "full_payment": "UPDATE contracts SET full_commission_flagged = 1 WHERE po_no = ?",
    "installment_1": "UPDATE contracts SET installment_1_commission_flagged = 1 WHERE po_no = ?",
    "installment_6": "UPDATE contracts SET installment_6_commission_flagged = 1 WHERE po_no = ?",
}


def process_commission_run(conn, as_of, run_date, source_filename, created_by_user):
    """
    Scans every active contract for newly-due commission across all
    three Phase 1 triggers, logs one commission_event per trigger
    raised, marks the corresponding flag so it's never raised again,
    and groups everything under one new commission_run row.

    A single contract can raise more than one event in the same run -
    e.g. the first time a mid-plan contract is ever imported, both
    installment 1 and installment 6 might already be paid.

    Returns (commission_run_id, list of dicts describing what was
    raised) - an empty list is a normal, expected outcome (nothing
    newly crossed since the last run), not an error.
    """
    candidates = conn.execute("SELECT * FROM contracts WHERE status = 'active'").fetchall()

    raised = []
    for contract in candidates:
        if full_payment_is_due(contract, as_of):
            raised.append({
                "po_no": contract["po_no"],
                "trigger_type": "full_payment",
                "trigger_date": contract["full_settlement_paid_date"],
                "amount": calculate_full_payment_commission(contract["net_price"]),
            })
        if installment_1_is_due(contract):
            raised.append({
                "po_no": contract["po_no"],
                "trigger_type": "installment_1",
                "trigger_date": contract["first_installment_paid_date"],
                "amount": calculate_installment_1_commission(contract["net_price"]),
            })
        if installment_6_is_due(contract):
            raised.append({
                "po_no": contract["po_no"],
                "trigger_type": "installment_6",
                "trigger_date": contract["sixth_installment_paid_date"],
                "amount": calculate_installment_6_commission(contract["net_price"]),
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
        conn.execute(_FLAG_COLUMN_UPDATE_SQL[event["trigger_type"]], (event["po_no"],))

    return run_id, raised
