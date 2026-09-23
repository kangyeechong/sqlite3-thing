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


def calculate_agency_agent_split(net_price, agency_pct, agent_pct, fb_lead_referred, deduction_pct):
    """
    AW Consultancy-style split: agency and agent are paid separately,
    at different percentages of Net Price. The FB-lead deduction (a
    purely manual flag - see contracts.fb_lead_referred) reduces only
    the agency's share, never the agent's, and only when set.

    Returns (agency_amount, agent_amount, total_amount) - total_amount
    is agency_amount + agent_amount, i.e. it already reflects the
    deduction, not the pre-deduction full split.
    """
    agent_amount = _calculate_commission(net_price, agent_pct)
    agency_pct_after_deduction = agency_pct - deduction_pct if fb_lead_referred else agency_pct
    agency_amount = _calculate_commission(net_price, agency_pct_after_deduction)
    total_amount = round(agency_amount + agent_amount, 2)
    return agency_amount, agent_amount, total_amount


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

# The reverse of _FLAG_COLUMN_UPDATE_SQL - see void_commission_event.
_FLAG_COLUMN_CLEAR_SQL = {
    "full_payment": "UPDATE contracts SET full_commission_flagged = 0 WHERE po_no = ?",
    "installment_1": "UPDATE contracts SET installment_1_commission_flagged = 0 WHERE po_no = ?",
    "installment_6": "UPDATE contracts SET installment_6_commission_flagged = 0 WHERE po_no = ?",
}

# Per trigger type: the flat percentage (used for 'flat' agencies), the
# agency/agent percentages (used for 'agency_agent_split' agencies),
# and the FB-lead deduction percentage for that trigger.
_TRIGGER_RULES = {
    "full_payment": {
        "flat_pct": rules.FULL_PAYMENT_COMMISSION_PCT,
        "agency_pct": rules.AW_AGENCY_FULL_PAYMENT_PCT,
        "agent_pct": rules.AW_AGENT_FULL_PAYMENT_PCT,
        "deduction_pct": rules.AW_FB_LEAD_DEDUCTION_FULL_PAYMENT_PCT,
    },
    "installment_1": {
        "flat_pct": rules.INSTALLMENT_1_COMMISSION_PCT,
        "agency_pct": rules.AW_AGENCY_INSTALLMENT_PCT,
        "agent_pct": rules.AW_AGENT_INSTALLMENT_PCT,
        "deduction_pct": rules.AW_FB_LEAD_DEDUCTION_INSTALLMENT_PCT,
    },
    "installment_6": {
        "flat_pct": rules.INSTALLMENT_6_COMMISSION_PCT,
        "agency_pct": rules.AW_AGENCY_INSTALLMENT_PCT,
        "agent_pct": rules.AW_AGENT_INSTALLMENT_PCT,
        "deduction_pct": rules.AW_FB_LEAD_DEDUCTION_INSTALLMENT_PCT,
    },
}


def _build_event(contract, trigger_type, trigger_date):
    """
    Computes one commission_event dict for a trigger already confirmed
    due. Branches on the contract's agency's commission_split_type -
    'flat' (the default, one figure) vs 'agency_agent_split' (AW
    Consultancy - two figures, with the manual FB-lead deduction
    applied to the agency's share only).
    """
    trigger_rules = _TRIGGER_RULES[trigger_type]
    net_price = contract["net_price"]

    if contract["commission_split_type"] == "agency_agent_split":
        agency_amount, agent_amount, amount = calculate_agency_agent_split(
            net_price,
            trigger_rules["agency_pct"],
            trigger_rules["agent_pct"],
            bool(contract["fb_lead_referred"]),
            trigger_rules["deduction_pct"],
        )
    else:
        amount = _calculate_commission(net_price, trigger_rules["flat_pct"])
        agency_amount, agent_amount = None, None

    return {
        "po_no": contract["po_no"],
        "trigger_type": trigger_type,
        "trigger_date": trigger_date,
        "amount": amount,
        "agency_amount": agency_amount,
        "agent_amount": agent_amount,
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
    candidates = conn.execute(
        """
        SELECT c.*, COALESCE(a.commission_split_type, 'flat') AS commission_split_type
        FROM contracts c
        LEFT JOIN agencies a ON a.agency_code = c.agency_code
        WHERE c.status = 'active'
        """
    ).fetchall()

    raised = []
    for contract in candidates:
        if full_payment_is_due(contract, as_of):
            raised.append(_build_event(contract, "full_payment", contract["full_settlement_paid_date"]))
        if installment_1_is_due(contract):
            raised.append(_build_event(contract, "installment_1", contract["first_installment_paid_date"]))
        if installment_6_is_due(contract):
            raised.append(_build_event(contract, "installment_6", contract["sixth_installment_paid_date"]))

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
        # status defaults to 'pending' in the schema too, but spelled
        # out here since it's the whole point of this INSERT: detecting
        # something is never the same as it being confirmed due - see
        # confirm_commission_events, the only other place status ever
        # changes.
        event_cursor = conn.execute(
            """
            INSERT INTO commission_events (
                po_no, trigger_type, trigger_date, amount,
                agency_amount, agent_amount,
                detected_at, detected_by_user, commission_run_id, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending')
            """,
            (
                event["po_no"], event["trigger_type"], event["trigger_date"],
                event["amount"], event["agency_amount"], event["agent_amount"],
                now_iso, created_by_user, run_id,
            ),
        )
        event["id"] = event_cursor.lastrowid
        conn.execute(_FLAG_COLUMN_UPDATE_SQL[event["trigger_type"]], (event["po_no"],))

    return run_id, raised


def load_events_for_run(conn, commission_run_id):
    """
    Every event (pending, confirmed, or voided) raised in one
    commission run, with enough contract context for a human to
    recognize what they're approving - the review page
    (app/web/routes.py) is the only caller.
    """
    return conn.execute(
        """
        SELECT e.id, e.po_no, e.trigger_type, e.trigger_date, e.amount, e.status,
               e.voided_at, e.voided_by_user, e.void_reason,
               c.agency_code, c.agent_name, cu.name AS customer_name
        FROM commission_events e
        JOIN contracts c ON c.po_no = e.po_no
        LEFT JOIN customers cu ON cu.customer_id = c.customer_id
        WHERE e.commission_run_id = ?
        ORDER BY e.po_no
        """,
        (commission_run_id,),
    ).fetchall()


def confirm_commission_events(conn, event_ids, confirmed_by_user):
    """
    Marks the given commission_events as confirmed - the one action
    that makes a detected candidate actually due. Only ever moves
    'pending' rows to 'confirmed'; an id that's already confirmed, or
    doesn't exist, is silently skipped rather than raising, so a
    double-submitted form or a stale checkbox can't cause an error.

    Returns how many rows were actually confirmed by this call.
    """
    if not event_ids:
        return 0
    now_iso = datetime.datetime.now().isoformat()
    event_ids = list(event_ids)

    # Batched well under SQLite's default SQLITE_MAX_VARIABLE_NUMBER
    # (999) - a "confirm all" selection on a very large run must not
    # blow that limit and fail the whole confirm with an
    # sqlite3.OperationalError, confirming nothing at all.
    batch_size = 500
    confirmed_count = 0
    for start in range(0, len(event_ids), batch_size):
        batch = event_ids[start : start + batch_size]
        placeholders = ",".join("?" for _ in batch)
        cursor = conn.execute(
            f"UPDATE commission_events SET status = 'confirmed', confirmed_at = ?, "
            f"confirmed_by_user = ? WHERE status = 'pending' AND id IN ({placeholders})",
            (now_iso, confirmed_by_user, *batch),
        )
        confirmed_count += cursor.rowcount
    return confirmed_count


def void_commission_event(conn, event_id, voided_by_user, reason):
    """
    Reverses a confirmed commission event that turned out to be wrong
    (a bad price, a receipt matched to the wrong PO, confirmed by
    mistake) - never deletes it, so the mistake and who corrected it
    stay visible forever (same "nothing hidden" standing-ledger
    philosophy as the rest of this app; see the 'voided' status
    comment in schema.sql), just excludes it from every report the same
    way a still-pending event already is.

    Clears the matching contracts.*_commission_flagged column so the
    next commission run can detect this PO's trigger fresh, once
    whatever was actually wrong has been corrected at the source (a
    re-uploaded Master report, typically). Voiding by itself never
    creates a new figure - it only un-sticks the PO so the ordinary
    detect -> review -> confirm pipeline can redo it correctly, the
    exact same human check a first-time detection gets.

    Only a 'confirmed' event can be voided - a still-pending one just
    shouldn't be confirmed in the first place, and an already-voided
    one can't be voided twice. Returns True if it actually voided
    something, False if event_id doesn't exist or isn't confirmed (a
    stale page, or a double-submitted form).
    """
    event = conn.execute(
        "SELECT po_no, trigger_type FROM commission_events WHERE id = ? AND status = 'confirmed'",
        (event_id,),
    ).fetchone()
    if event is None:
        return False

    now_iso = datetime.datetime.now().isoformat()
    conn.execute(
        "UPDATE commission_events SET status = 'voided', voided_at = ?, "
        "voided_by_user = ?, void_reason = ? WHERE id = ?",
        (now_iso, voided_by_user, reason, event_id),
    )
    conn.execute(_FLAG_COLUMN_CLEAR_SQL[event["trigger_type"]], (event["po_no"],))
    return True
