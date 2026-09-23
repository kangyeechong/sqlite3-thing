"""
Builds the downloadable Excel report - a full standing ledger, not a
"what's newly due this run" list. Every contract in the database
appears, always: paid, still pending, or cancelled/withdrawn - so a PO
that hasn't paid yet, or was cancelled, never disappears from view.
`commission_run_id` only controls which cells are highlighted yellow
(whatever got confirmed in that specific run) and gates the download
itself (refuses if that run confirmed nothing at all, since there'd be
nothing new to justify a fresh file).

Column layout deliberately mirrors the real Kenjin Master Report
(docs/data_model.md section 2) rather than a simplified summary: one
row per PO, with the Full/1st Half/Balance Half columns filled in from
whatever has EVER been confirmed for that PO (not just this run) - so
a PO with two triggers confirmed in the same run (rare, but real - see
docs/data_model.md section 5) still gets ONE row, not two, exactly
like the source file would show it.

Grouping mirrors how your team already splits the Master report today
(section 6a): every agency GROUP gets its own combined sheet first
(several agency_codes can be sub-codes of one real-world agency, like
AW Consultancy's AC108-01/-02/-03 - see agencies.agency_group), and
where any code in that group has `splits_by_agent` on, separate
standalone sheets follow, one per individual agent - matching the real
file's actual sheet order (a combined agency sheet, then individually
named agent sheets), not jumping straight from raw codes to per-agent
sheets with no combined view in between.
"""

import datetime

from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import commission, rules

_BODY_FONT = Font(name="Arial", size=11)
_HEADER_FONT = Font(name="Arial", size=11, bold=True)
_TITLE_FONT = Font(name="Arial", size=13, bold=True)
_MONEY_FORMAT = "#,##0.00"

# Green matched against the real sample file's actual cell formatting
# (not guessed): rows there use theme accent6 (#70AD47) tinted 0.6,
# reproduced here as plain RGB (Excel's tint formula applied by hand)
# since openpyxl's fill doesn't need to reference the workbook's theme
# to look the same. Confirmed with the business: green means the PO is
# fully paid off commission-wise, not specifically "one-off full
# payment" - an installment plan's Balance Half (its last commission
# trigger) gets the same whole-row green as a full payment does.
_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
_GREEN_FILL = PatternFill(start_color="C6DEB5", end_color="C6DEB5", fill_type="solid")   # fully paid off rows

# Matched against the real sample file's actual cell formatting: theme
# accent2 (#ED7D31) tinted 0.8. Cancelled/withdrawn POs get this shade
# across the whole row - they still appear (see _load_master_rows),
# just visibly marked as "nothing more is ever due here" rather than
# disappearing.
_BEIGE_FILL = PatternFill(start_color="FBE5D6", end_color="FBE5D6", fill_type="solid")

# Marks the FB-lead-deduction cell in the agency/agent split columns
# (see _write_agency_agent_split_columns) when a row's relevant trigger
# fired but fb_lead_referred is off, i.e. "this slot exists, nothing
# was deducted here". Approximated from a screenshot, not an actual
# .xlsx file this time (unlike green/yellow/beige, which were matched
# against real file bytes) - flag if the shade needs adjusting once
# there's a real file with an FB-lead deduction to check against.
_GREY_FILL = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

# The agency/agent split table is a dense block of otherwise-identical
# money columns bolted onto the right of the main table - a thin
# border around every cell (header and data alike) is what makes it
# read as its own table at a glance instead of bleeding into the
# columns next to it.
_THIN_SIDE = Side(style="thin", color="000000")
_THIN_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)

COMPANY_SHORT_NAME = "XEKL"

# Public (not underscore-prefixed): the web layer's results page
# imports this too, so the on-screen review a user checks before
# downloading always shows the same labels as the file they then
# download - one definition, not two copies that can quietly drift
# apart.
TRIGGER_LABELS = {
    "full_payment": "Full Payment",
    "installment_1": "Instalment 1",
    "installment_6": "Instalment 6 (Balance)",
}

# (column header, key into the row dict) - same order and wording as
# the real Master Report, confirmed against the sample file.
_COLUMNS = [
    ("No", "row_no"),
    ("PO No", "po_no"),
    ("PO Date", "po_date"),
    ("Signature Date", "signature_date"),
    ("Customer ID", "customer_id"),
    ("Customer Name", "customer_name"),
    ("Lot No", "lot_no"),
    ("Niche/Tablet Price (RM)", "niche_price"),
    ("Promotion (RM)", "promotion"),
    ("Discount (RM)", "discount"),
    ("Nett Price (RM)", "net_price"),
    ("Cooling Off Period", "cooling_off_period"),
    ("Full Settlement Paid Date", "full_settlement_paid_date"),
    ("Full Payment Commission (RM)", "full_payment_commission"),
    ("Full Commission Paid Date", "full_commission_paid_date"),
    ("First Instalment Paid Date", "first_installment_paid_date"),
    ("1st Half Commission (RM)", "installment_1_commission"),
    ("1st Half Commission Paid Date", "installment_1_commission_paid_date"),
    ("Sixth Instalment Paid Date", "sixth_installment_paid_date"),
    ("Balance Half Commission (RM)", "installment_6_commission"),
    ("Balance Half Commission Paid Date", "installment_6_commission_paid_date"),
    ("FCC/Agent", "agent_name"),
    ("Agency Code", "agency_code"),
    ("Remarks", "remarks"),
]
_MONEY_COLUMNS = {
    "Niche/Tablet Price (RM)", "Promotion (RM)", "Discount (RM)", "Nett Price (RM)",
    "Full Payment Commission (RM)", "1st Half Commission (RM)", "Balance Half Commission (RM)",
}
# Instalment columns that get a yellow CELL highlight when THIS run put
# a value in them. Full payment doesn't get a cell highlight - it gets
# the whole ROW shaded green instead (see _GREEN_FILL / _write_table),
# matching the real file's convention exactly: full payment marks the
# entire row, instalments mark just the specific commission cell.
_HIGHLIGHT_COLUMNS = {"1st Half Commission (RM)", "Balance Half Commission (RM)"}

_SUMMARY_COLUMNS = [
    "DATE RECORD", "Full Commission", "First Half Commission",
    "Second Half Commission", "Running Total", "Remarks",
]


def _format_title_date(iso_date_str):
    return datetime.date.fromisoformat(iso_date_str).strftime("%d %B %Y").upper()


def _format_short_date(iso_date_str):
    return datetime.date.fromisoformat(iso_date_str).strftime("%d/%m/%Y")


def _cooling_off_status(signature_date_str, run_date):
    """
    "Cooling Off Period" in the real file is a general per-PO field -
    it shows EXPIRED once COOLING_OFF_TOTAL_DAYS have passed since
    Signature Date, on every row that has a signature date on file, no
    matter which commission trigger (if any) brought that row into
    this report. Confirmed against the real file: rows with an
    instalment due, not just a full payment due, show EXPIRED too, and
    a PO with no payment activity at all still shows EXPIRED as long
    as it's old enough - this is unrelated to
    COOLING_OFF_DAYS_BEFORE_RELEASE (which only gates when full
    payment commission itself becomes releasable).
    """
    if not signature_date_str:
        return None
    signature_date = datetime.date.fromisoformat(signature_date_str)
    as_of = run_date if isinstance(run_date, datetime.date) else datetime.date.fromisoformat(run_date)
    if (as_of - signature_date).days >= rules.COOLING_OFF_TOTAL_DAYS:
        return "EXPIRED"
    return None


# Exact-match only, not a substring purge - these are the only two
# real values seen for this in the actual export, always as the
# WHOLE remarks text on their own, never combined with anything else
# worth keeping. A genuine remark that happens to mention "addendum"
# as part of a longer note should still come through untouched.
_ADDENDUM_ONLY_REMARKS = {"with addendum a", "without addendum a"}


def _clean_remarks(remarks):
    """
    Strips a purely legal/admin note ("With Addendum A" / "Without
    Addendum A") that doesn't belong in the commission report - not
    something Accounts or an agent needs to see when reading this
    file, confirmed with the business. Every other Remarks value
    (cancellation notes, inurnment dates, Lot switches, ...) passes
    through unchanged.
    """
    if remarks and str(remarks).strip().lower() in _ADDENDUM_ONLY_REMARKS:
        return None
    return remarks


def _load_master_rows(conn, commission_run_id, run_date, period_start=None, period_end=None, latest_run_id=None):
    """
    Every contract in the database, always - paid, still pending, or
    cancelled/withdrawn. A PO that hasn't paid anything yet, or was
    cancelled, still gets a row here; it just shows blank/beige instead
    of a commission figure. This is what makes the report a standing
    ledger rather than a "what's due this run" list - see the module
    docstring.

    Commission figures come from whatever has EVER been confirmed for
    that PO (any run, not just this one) - a PO's Full Payment
    Commission, once confirmed, keeps showing on every future download,
    not just the run that confirmed it. Each trigger also carries an
    "confirmed_this_run" flag so _write_table can tell which cells (if
    any) are the newly-added ones to highlight yellow, versus older
    confirmed figures that just carry forward plainly.

    `run_date` is needed to compute each row's Cooling Off Period
    status (EXPIRED once COOLING_OFF_TOTAL_DAYS have passed since
    Signature Date - see _cooling_off_status).

    period_start/period_end (both required together, or both left
    None): switches to the scoped mode generate_period_report uses.
    Confirmed with the business: "August's report" means every PO
    PURCHASED in August (po_date in range) - a June-purchased PO whose
    installment happens to get confirmed in August stays on JUNE's
    report, not August's (re-download June's report any time to pick
    up a late-confirmed payment for it - same "whatever's confirmed so
    far" logic as the unscoped report, just pre-filtered to that
    month's own contracts). Every commission event ever confirmed for
    an in-period PO shows here, regardless of when it was confirmed -
    this is NOT limited to events confirmed within the period itself.
    The pre-tool historical-absorption fallback (see the unscoped
    branch below) is still skipped in this mode - by agreement, kept
    simple rather than also reconstructing a per-month breakdown of
    that lump-sum figure.

    latest_run_id: in period mode, which commission_run_id counts as
    "just confirmed" for the yellow-cell highlight - see
    _latest_confirmed_run_for_period. Regression fix: this used to be
    hardcoded to "every confirmed event in period mode is this run's
    own movement", which meant EVERY cell ever confirmed for an
    in-period PO stayed yellow forever, on every single re-download -
    round 1's payments never faded back to plain once round 2 added
    more. Ignored entirely outside period mode, where commission_run_id
    (this function's own first argument) is compared directly instead.
    """
    period_clause = "WHERE c.po_date BETWEEN ? AND ?" if period_start is not None else ""
    contract_rows = conn.execute(
        f"""
        SELECT
            c.po_no, c.po_date, c.signature_date, c.customer_id, c.lot_no,
            c.niche_price, c.promotion, c.discount, c.net_price, c.status,
            c.full_settlement_paid_date, c.first_installment_paid_date,
            c.sixth_installment_paid_date, c.agent_name, c.agency_code, c.remarks,
            c.fb_lead_referred,
            c.full_commission_flagged, c.installment_1_commission_flagged,
            c.installment_6_commission_flagged,
            c.full_commission_paid_date, c.installment_1_commission_paid_date,
            c.installment_6_commission_paid_date,
            cu.name AS customer_name,
            a.splits_by_agent, a.agency_group, a.commission_split_type
        FROM contracts c
        LEFT JOIN customers cu ON cu.customer_id = c.customer_id
        LEFT JOIN agencies a ON a.agency_code = c.agency_code
        {period_clause}
        ORDER BY c.po_no
        """,
        (period_start, period_end) if period_start is not None else (),
    ).fetchall()
    if period_start is not None and not contract_rows:
        return []

    by_po = {}
    for r in contract_rows:
        by_po[r["po_no"]] = {
            "po_no": r["po_no"],
            "po_date": r["po_date"],
            "signature_date": r["signature_date"],
            "customer_id": r["customer_id"],
            "customer_name": r["customer_name"],
            "lot_no": r["lot_no"],
            "niche_price": r["niche_price"],
            "promotion": r["promotion"],
            "discount": r["discount"],
            "net_price": r["net_price"],
            # Cancelled/withdrawn rows are shaded beige in _write_table
            # regardless of any commission history - see is_cancelled_row.
            "status": r["status"],
            "cooling_off_period": _cooling_off_status(r["signature_date"], run_date),
            # The paid-date itself is a raw fact from the sheet, shown
            # whether or not the resulting commission has been
            # confirmed yet; the commission figure below only appears
            # once actually confirmed.
            "full_settlement_paid_date": r["full_settlement_paid_date"],
            "full_payment_commission": None,
            "full_payment_confirmed_this_run": False,
            # Accounts fills these in by hand on the real file once
            # they've actually sent the money - read back from
            # whatever the most recent upload had on file, never
            # computed or written by this tool.
            "full_commission_paid_date": r["full_commission_paid_date"],
            "first_installment_paid_date": r["first_installment_paid_date"],
            "installment_1_commission": None,
            "installment_1_confirmed_this_run": False,
            "installment_1_commission_paid_date": r["installment_1_commission_paid_date"],
            "sixth_installment_paid_date": r["sixth_installment_paid_date"],
            "installment_6_commission": None,
            "installment_6_confirmed_this_run": False,
            "installment_6_commission_paid_date": r["installment_6_commission_paid_date"],
            "agent_name": r["agent_name"] or "(unassigned)",
            "agency_code": r["agency_code"] or "(No Agency)",
            # Several agency_codes can share one real-world agency
            # (AW Consultancy's AC108-01/-02/-03) - agency_group is
            # the label the combined top-level sheet groups under.
            # Falls back to the raw agency_code when no group is
            # set, so every other agency still gets its own single
            # sheet exactly as before this existed.
            "agency_group": r["agency_group"] or r["agency_code"] or "(No Agency)",
            "remarks": _clean_remarks(r["remarks"]),
            # No agency on file -> nothing to group by agent for
            # either; default to a flat listing rather than
            # splitting by agent.
            "splits_by_agent": bool(r["splits_by_agent"]) if r["agency_code"] else False,
            "commission_split_type": r["commission_split_type"] or "flat",
            "fb_lead_referred": bool(r["fb_lead_referred"]),
            # Used only below, to detect a trigger that's flagged but
            # has no commission_event at all (see the historical-
            # fallback loop right after this) - not read anywhere else.
            "full_commission_flagged": bool(r["full_commission_flagged"]),
            "installment_1_commission_flagged": bool(r["installment_1_commission_flagged"]),
            "installment_6_commission_flagged": bool(r["installment_6_commission_flagged"]),
            # Only populated for agency_agent_split agencies (AW
            # Consultancy) - the per-trigger Agency/Agent split
            # figures shown in the extra columns to the right of
            # the main table (see _write_agency_agent_split_columns).
            "full_payment_agency_amount": None,
            "full_payment_agent_amount": None,
            "installment_1_agency_amount": None,
            "installment_1_agent_amount": None,
            "installment_6_agency_amount": None,
            "installment_6_agent_amount": None,
        }

    # Unfiltered by date on purpose, in both modes - by_po (built from
    # contract_rows above) is already scoped to the right PO No's; an
    # event for a PO not in that set is simply skipped below (`row is
    # None`). In period mode this means EVERY confirmed event ever
    # raised for an in-period PO shows here, whenever it was confirmed
    # - not just ones confirmed "during" the period - matching "show
    # whatever's confirmed so far for this month's contracts".
    event_rows = conn.execute(
        """
        SELECT po_no, trigger_type, amount, agency_amount, agent_amount, commission_run_id
        FROM commission_events
        WHERE status = 'confirmed'
        """
    ).fetchall()

    for e in event_rows:
        row = by_po.get(e["po_no"])
        if row is None:
            continue  # a contract row always exists for a real event; defensive only
        # confirmed_this_run drives two things in _write_table: the
        # yellow cell highlight on an installment commission cell, and
        # the "movement as at" total beneath it. In period mode there's
        # no single commission_run_id tied to "this download" the way
        # the unscoped report has one - latest_run_id (computed once by
        # generate_period_report, see _latest_confirmed_run_for_period)
        # stands in for it: only the round of confirmations that most
        # recently touched this period counts as new, so an earlier
        # round's cells fade back to plain on a later re-download
        # instead of staying yellow forever.
        confirmed_this_run = (
            e["commission_run_id"] == latest_run_id if period_start is not None
            else e["commission_run_id"] == commission_run_id
        )
        if e["trigger_type"] == "full_payment":
            row["full_payment_commission"] = e["amount"]
            row["full_payment_confirmed_this_run"] = confirmed_this_run
            row["full_payment_agency_amount"] = e["agency_amount"]
            row["full_payment_agent_amount"] = e["agent_amount"]
        elif e["trigger_type"] == "installment_1":
            row["installment_1_commission"] = e["amount"]
            row["installment_1_confirmed_this_run"] = confirmed_this_run
            row["installment_1_agency_amount"] = e["agency_amount"]
            row["installment_1_agent_amount"] = e["agent_amount"]
        elif e["trigger_type"] == "installment_6":
            row["installment_6_commission"] = e["amount"]
            row["installment_6_confirmed_this_run"] = confirmed_this_run
            row["installment_6_agency_amount"] = e["agency_amount"]
            row["installment_6_agent_amount"] = e["agent_amount"]

    # A trigger can be flagged (contracts.*_commission_flagged) with no
    # commission_event at all, pending or confirmed - that only ever
    # happens when _flag_historically_accounted_commissions
    # (app/importer.py) found it already covered by the sheet's own
    # trailing "Date Record" history at import time: a fresh detection
    # (process_commission_run) always creates a matching event in the
    # same breath it sets the flag, so "flagged with zero events" can
    # only mean "already accounted for before this tool ever saw it."
    # Without a fallback here, that real, already-known commission
    # simply vanishes from the report - no amount, no green "fully
    # paid" shading, and (for an agency_agent_split agency) nothing in
    # the Agency/Agent split columns either, even though the real
    # Master Report this data came from already had it filled in by
    # hand. Recomputed the exact same way a fresh detection would (net
    # price x the fixed percentage, via the same commission._build_event
    # every live detection uses) rather than read back verbatim from
    # the sheet - for the same reason net_price itself is always
    # recomputed rather than trusted, and already cross-checked exactly
    # against the real file's own historical figures (see rules.py).
    # Nothing here counts as newly confirmed: no confirmed_this_run
    # flag gets touched, so it never highlights yellow or contributes
    # to a "movement as at" figure.
    # Skipped entirely in period mode - see this function's docstring:
    # pre-tool historical-absorption money has no clean per-day date to
    # filter by, so it never belongs in a period-scoped download.
    if period_start is None:
        has_event = {(e["po_no"], e["trigger_type"]) for e in conn.execute(
            "SELECT DISTINCT po_no, trigger_type FROM commission_events"
        )}
        historical_fallback_triggers = (
            ("full_payment", "full_commission_flagged", "full_payment_commission",
             "full_payment_agency_amount", "full_payment_agent_amount"),
            ("installment_1", "installment_1_commission_flagged", "installment_1_commission",
             "installment_1_agency_amount", "installment_1_agent_amount"),
            ("installment_6", "installment_6_commission_flagged", "installment_6_commission",
             "installment_6_agency_amount", "installment_6_agent_amount"),
        )
        for row in by_po.values():
            for trigger_type, flag_key, amount_key, agency_key, agent_key in historical_fallback_triggers:
                if row[flag_key] and (row["po_no"], trigger_type) not in has_event:
                    fallback = commission._build_event(row, trigger_type, trigger_date=None)
                    row[amount_key] = fallback["amount"]
                    row[agency_key] = fallback["agency_amount"]
                    row[agent_key] = fallback["agent_amount"]

    return list(by_po.values())


def _reconstruct_scoped_historical_entries(conn, agency_group, agent_name):
    """
    The sheet's own pre-existing "Date Record" history has no
    per-agency or per-agent breakdown to read (the real file only ever
    carries one company-wide total per "As at" date) - reconstructed
    here instead for a scoped (agency/agent) summary table, so it isn't
    left permanently empty for every sheet except "All".

    Attributes each historically-absorbed trigger (flagged, with no
    commission_event at all - see _flag_historically_accounted_commissions
    in app/importer.py, the only way that combination arises) to the
    EARLIEST "As at" cutoff on or after its own paid-date - the same
    cycle the old manual process would have first recognized it in -
    using the exact same commission math live detection uses
    (commission._build_event). Verified against the real file: summed
    back up across every agency, this reproduces the company-wide
    historical totals to the cent (4,155.00 / 15,569.25 / 25,822.50 /
    1,698.75 / 16,785.00 / 1,462.50, matching every real "As at" row
    exactly, not just the RM65,493.00 grand total).

    Uses the combined trigger amount (agency+agent together) when
    scoped to an agency/agency-group, matching the agency's own main
    table on that sheet - but the AGENT's own cut when scoped to a
    specific agent, for the same reason: a per-agent sheet's main
    table already shows just that agent's money (see the historical
    fallback in _load_master_rows), so this summary has to match it,
    not the full agency+agent total. Confirmed against a real
    per-agent reference table (RM1,292.00 / RM772.00 / RM440.00 per
    cycle, RM2,504.00 total - exactly the agent's 8%/4% cut, not the
    combined 15%/7.5%).
    """
    cutoffs = [r["date_record"] for r in conn.execute(
        "SELECT date_record FROM historical_summary_rows ORDER BY date_record"
    )]
    if not cutoffs:
        return []

    candidates = conn.execute(
        """
        SELECT c.*, COALESCE(a.commission_split_type, 'flat') AS commission_split_type,
               a.agency_group
        FROM contracts c
        LEFT JOIN agencies a ON a.agency_code = c.agency_code
        WHERE c.status = 'active'
        """
    ).fetchall()
    has_event = {
        (e["po_no"], e["trigger_type"])
        for e in conn.execute("SELECT po_no, trigger_type FROM commission_events")
    }

    triggers = (
        ("full_payment", "full_commission_flagged", "full_settlement_paid_date", "full_commission"),
        ("installment_1", "installment_1_commission_flagged", "first_installment_paid_date", "first_half_commission"),
        ("installment_6", "installment_6_commission_flagged", "sixth_installment_paid_date", "second_half_commission"),
    )

    buckets = {}
    for contract in candidates:
        contract_agency_group = contract["agency_group"] or contract["agency_code"] or "(No Agency)"
        contract_agent_name = contract["agent_name"] or "(unassigned)"
        if agency_group is not None and contract_agency_group != agency_group:
            continue
        if agent_name is not None and contract_agent_name != agent_name:
            continue

        for trigger_type, flag_col, date_col, bucket_key in triggers:
            if not contract[flag_col] or (contract["po_no"], trigger_type) in has_event:
                continue
            paid_date = contract[date_col]
            if not paid_date:
                continue
            cutoff = next((co for co in cutoffs if co >= paid_date), None)
            if cutoff is None:
                continue  # paid after even the latest known cutoff - not historical, live detection covers it
            event = commission._build_event(dict(contract), trigger_type, trigger_date=None)
            bucket = buckets.setdefault(
                cutoff, {
                    "full_commission": 0.0, "first_half_commission": 0.0, "second_half_commission": 0.0,
                    "agent_commission": 0.0, "fb_lead_deduction": 0.0,
                }
            )
            # Agent-scoped: the agent's own cut (falls back to the
            # full amount for a flat agency, where agent_amount is
            # NULL). Agency/group-scoped: the full combined amount,
            # unchanged, plus the group's own "Less Agent Commission" /
            # "Less FB leads" breakdown (see _write_summary_table) -
            # only meaningful at the group level, where both shares are
            # visible side by side, same reasoning as the main table's
            # own split columns.
            if agent_name is not None:
                bucket[bucket_key] += event["agent_amount"] if event["agent_amount"] is not None else event["amount"]
            else:
                bucket[bucket_key] += event["amount"]
                if event["agent_amount"] is not None:
                    bucket["agent_commission"] += event["agent_amount"]
                    deduction_pct = _DEDUCTION_PCT_BY_TRIGGER.get(trigger_type)
                    if deduction_pct is not None:
                        deduction = _fb_deduction_amount(contract["net_price"], contract["fb_lead_referred"], deduction_pct)
                        bucket["fb_lead_deduction"] += deduction or 0.0

    return [
        (
            cutoff, None, b["full_commission"], b["first_half_commission"], b["second_half_commission"], None,
            b["agent_commission"], b["fb_lead_deduction"],
        )
        for cutoff, b in sorted(buckets.items())
    ]


def _load_summary_rows(conn, agency_group=None, agent_name=None, period_start=None, period_end=None):
    """
    Every commission run ever processed (not just this one), grouped
    and pivoted into the "Date Record" running-total table from the
    original workflow. Built entirely from commission_events/
    commission_runs, which already record everything needed - this is
    a new view over existing data, not new calculation logic.

    Only confirmed events count - a run sitting fully pending (nothing
    approved on its review page yet) simply doesn't produce a row here
    yet, the same way it doesn't produce one in the main table either.

    Every sheet gets its own copy of this table, scoped to that
    slice's own history - confirmed against the real file: XEMP's own
    sheet (agency_group falls back to agency_code AC001, since it has
    no group) has a running total of RM33,070.50, a real subset of the
    RM65,493.00 grand total on the "All" sheet, not a different figure
    entirely. Passing neither filter (the "All" sheet's case) sums
    everything, exactly like before this existed.

    The unscoped/"All" case merges in the sheet's own pre-existing
    "Date Record" history verbatim (historical_summary_rows - imported
    once from the uploaded file itself, see app/importer.py)
    chronologically alongside this tool's own tracked runs, so a real
    file's history from before this tool ever existed carries straight
    through instead of the running total silently starting over from
    zero. A scoped (agency/agent) call instead merges in a
    RECONSTRUCTED per-scope breakdown of that same history - see
    _reconstruct_scoped_historical_entries - since the real file itself
    only ever carries one company-wide total per "As at" date, not a
    per-agency one.

    Each row keeps its run_id so _write_summary_table can tell which
    row is the one just processed and highlight only that one yellow -
    confirmed against the real file: every prior "As at" row stays
    plain, only the newest addition (and the grand total line under
    it) is highlighted. Historical rows carry run_id=None, which never
    matches a real commission_run_id, so they're never highlighted.
    """
    # fb_lead_deduction is deliberately NOT summed in SQL as
    # `net_price * deduction_pct` - that raw product is never rounded
    # to the cent the way every other commission figure in this app is
    # (see commission._calculate_commission's ROUND_HALF_UP - the same
    # rounding _fb_deduction_amount and _reconstruct_scoped_historical_
    # entries already use for this exact figure). Summed across many
    # confirmed events, that unrounded-per-event total silently drifts
    # away from the per-row deduction figures shown in this same
    # report's own Agency/Agent split columns (_write_agency_agent_
    # split_columns) and from the historically-reconstructed entries
    # merged into this same table below - by as much as half a cent
    # per event. Fetching the per-event fields here and rounding each
    # one exactly like every other commission figure keeps this table
    # internally consistent with the rest of the report.
    query = """
        SELECT r.id AS run_id, r.run_date, e.trigger_type, {amount_expr} AS amount,
               e.agent_amount, a.commission_split_type, c.fb_lead_referred, c.net_price
        FROM commission_runs r
        JOIN commission_events e ON e.commission_run_id = r.id AND e.status = 'confirmed'
        JOIN contracts c ON c.po_no = e.po_no
        LEFT JOIN agencies a ON a.agency_code = c.agency_code
        WHERE 1=1
    """.format(
        # Agent-scoped: the agent's own cut (falls back to the full
        # amount for a flat agency, where agent_amount is NULL).
        # Agency/group-scoped and the unscoped "All" case: the full
        # combined amount, unchanged - see _reconstruct_scoped_
        # historical_entries above for the matching fix on the
        # historically-absorbed side of this same table.
        amount_expr="COALESCE(e.agent_amount, e.amount)" if agent_name is not None else "e.amount"
    )
    params = []
    if agency_group is not None:
        # Matches the same 3-level fallback _load_master_rows uses to
        # build the "(No Agency)" group key in the first place
        # (agency_group -> agency_code -> the literal string) - without
        # that last fallback, COALESCE(NULL, NULL) is NULL, which never
        # equals the string "(No Agency)" the caller actually passes
        # in for a contract with no agency at all, so that sheet's own
        # Date Record summary always came up empty even though its
        # main table correctly showed confirmed commissions.
        query += " AND COALESCE(a.agency_group, c.agency_code, '(No Agency)') = ?"
        params.append(agency_group)
    if agent_name is not None:
        query += " AND COALESCE(c.agent_name, '(unassigned)') = ?"
        params.append(agent_name)
    if period_start is not None:
        # Filters by the PO's own purchase date, not by when the event
        # was confirmed - matches _load_master_rows: "August's report"
        # means August-purchased contracts, whatever's been confirmed
        # for them so far, whenever that happened. A run with events
        # for both August- and July-purchased POs still shows up here,
        # but its row only reflects the August-purchased slice - the
        # "As at" date is still the run's own confirm date (which could
        # itself be well after August), not the PO's purchase date.
        query += " AND c.po_date BETWEEN ? AND ?"
        params.append(period_start)
        params.append(period_end)
    query += " ORDER BY r.run_date, r.id"

    cursor = conn.execute(query, params)

    # agent_commission/fb_lead_deduction are only ever displayed on a
    # split-group's own combined sheet (see _write_summary_table's
    # split_group_name) - accumulated here unconditionally anyway since
    # it's cheap and keeps this code's shape the same regardless of
    # scope, rather than branching on it.
    by_run = {}
    run_order = []
    for row in cursor.fetchall():
        run_id = row["run_id"]
        if run_id not in by_run:
            by_run[run_id] = {
                "run_date": row["run_date"], "full_payment": 0.0, "installment_1": 0.0, "installment_6": 0.0,
                "agent_commission": 0.0, "fb_lead_deduction": 0.0,
            }
            run_order.append(run_id)
        bucket = by_run[run_id]
        bucket[row["trigger_type"]] += row["amount"] or 0.0
        bucket["agent_commission"] += row["agent_amount"] or 0.0
        if row["commission_split_type"] == "agency_agent_split" and row["fb_lead_referred"]:
            deduction_pct = _DEDUCTION_PCT_BY_TRIGGER.get(row["trigger_type"])
            if deduction_pct is not None:
                bucket["fb_lead_deduction"] += commission._calculate_commission(row["net_price"], deduction_pct)

    # (sort_date, run_id_or_None, full, first_half, second_half, remarks,
    #  agent_commission, fb_lead_deduction)
    entries = []
    for run_id in run_order:
        r = by_run[run_id]
        entries.append((
            r["run_date"], run_id, r["full_payment"], r["installment_1"], r["installment_6"], None,
            r["agent_commission"], r["fb_lead_deduction"],
        ))

    # Skipped entirely in period mode, same reasoning as
    # _load_master_rows: pre-tool historical-absorption money has no
    # clean per-day date to filter by.
    if period_start is None:
        if agency_group is None and agent_name is None:
            historical = conn.execute(
                "SELECT date_record, full_commission, first_half_commission, second_half_commission, remarks "
                "FROM historical_summary_rows"
            ).fetchall()
            for h in historical:
                entries.append((
                    h["date_record"], None,
                    h["full_commission"], h["first_half_commission"], h["second_half_commission"],
                    h["remarks"], 0.0, 0.0,
                ))
        else:
            entries.extend(_reconstruct_scoped_historical_entries(conn, agency_group, agent_name))

    entries.sort(key=lambda e: e[0])

    # "Running Total" is a misnomer inherited verbatim from the real
    # file's own column header - confirmed against the real file, it is
    # NOT a cumulative sum across rows, just each cycle's own
    # Full+First+Second total. The true across-every-cycle cumulative
    # figure only ever appears once, on the separate "Total Sum of
    # Commission Payout" line below the table (see _write_summary_table)
    # - conflating the two here previously made every row past the
    # first show a progressively larger, wrong figure.
    summary_rows = []
    for sort_date, run_id, full, first_half, second_half, remarks, agent_commission, fb_lead_deduction in entries:
        # Rounded to the cent here, once, rather than left as whatever
        # binary-float noise summing several already-rounded event
        # amounts happens to produce (e.g. 1851.85 - 987.65 landing on
        # 864.1999999999999) - every other money figure in this report
        # is written as a clean 2-decimal value (see _write_table's own
        # `round(totals[key], 2)`), and the exported .xlsx cells here
        # should be too, not just look right thanks to the display-only
        # number_format.
        running_total = round(full + first_half + second_half, 2)
        agent_commission = round(agent_commission, 2)
        fb_lead_deduction = round(fb_lead_deduction, 2)
        summary_rows.append({
            "run_id": run_id,
            "date_record": f"As at {_format_short_date(sort_date)}",
            "full_commission": round(full, 2),
            "first_half_commission": round(first_half, 2),
            "second_half_commission": round(second_half, 2),
            "running_total": running_total,
            # Only meaningful (and only ever displayed - see
            # _write_summary_table's split_group_name) on a split
            # group's own combined sheet; otherwise 0 and unused.
            "agent_commission": agent_commission,
            "fb_lead_deduction": fb_lead_deduction,
            # NOT running_total - agent_commission - fb_lead_deduction:
            # running_total is already built from the combined `amount`
            # column, which already has the FB-lead deduction baked in
            # (agency_amount is computed post-deduction - see
            # commission.calculate_agency_agent_split and the schema
            # comment on commission_events.agency_amount). Subtracting
            # fb_lead_deduction again here would deduct it a second
            # time. fb_lead_deduction is shown purely as an informational
            # breakdown of how much of the reduction already reflected
            # in running_total was FB-lead-related, not a further
            # subtraction to apply.
            "total_for_group": round(running_total - agent_commission, 2),
            "remarks": remarks,
        })
    return summary_rows


def _safe_sheet_title(name):
    """Excel sheet names can't exceed 31 chars or contain []:*?/\\ """
    cleaned = "".join(c for c in str(name) if c not in "[]:*?/\\")
    return cleaned[:31] or "Sheet"


def _unique_sheet_title(name, used_titles):
    """
    Like _safe_sheet_title, but guaranteed not to collide with a title
    already in used_titles. Appends a numeric suffix and re-truncates
    to fit the 31-character limit - not a trailing "_", which a 31+
    character base name would immediately truncate right back off,
    producing the same title forever.
    """
    base = _safe_sheet_title(name)
    if base not in used_titles:
        return base
    for n in range(2, 1000):
        suffix = f"_{n}"
        candidate = base[: 31 - len(suffix)] + suffix
        if candidate not in used_titles:
            return candidate
    raise ValueError(f"Could not find a unique sheet title for '{name}'")


def _title_line(label, rows, run_date):
    """
    "{LABEL} OVERALL COMMISSION PAYOUT AS AT {date}\n(PURCHASE ORDER
    FROM {min} UNTIL {max})" - matching the real Master Report's title
    convention. The PO date range is computed from the rows actually
    in this run (real data, not an invented reporting period) and
    omitted entirely if none of them have a PO Date on file.
    """
    as_at = _format_title_date(run_date.isoformat()) if isinstance(run_date, datetime.date) else _format_title_date(run_date)
    title = f"{label} OVERALL COMMISSION PAYOUT AS AT {as_at}"

    po_dates = sorted(r["po_date"] for r in rows if r["po_date"])
    if po_dates:
        title += f"\n(PURCHASE ORDER FROM {_format_title_date(po_dates[0])} UNTIL {_format_title_date(po_dates[-1])})"
    return title


# Built once from _COLUMNS so the total row can place each subtotal
# under its own real column by key lookup - never by counting columns
# from the end, which silently breaks the moment _COLUMNS changes
# shape (exactly what happened here: the old 9-column offsets were
# left in place after the layout grew to 24 columns, and the "Total"
# label and grand-total figure quietly landed under FCC/Agent and
# Agency Code instead of any money column).
_COLUMN_INDEX = {key: i for i, (_label, key) in enumerate(_COLUMNS, start=1)}
_TOTAL_KEYS = (
    "niche_price", "promotion", "discount", "net_price",
    "full_payment_commission", "installment_1_commission", "installment_6_commission",
)

# The three commission-amount keys - cleared to blank (not just tinted
# beige) on a cancelled/withdrawn row, per the business: once a PO is
# marked cancelled, nothing is owed on it any more, so showing a
# leftover figure there would read as still-payable. Price columns
# (Niche/Promotion/Discount/Nett) are left alone - those are just the
# sale's own record, not a "what's owed" figure.
_COMMISSION_VALUE_KEYS = {"full_payment_commission", "installment_1_commission", "installment_6_commission"}

# Maps each commission-amount column key to the row-dict flag that says
# whether THIS run is the one that confirmed it. Two uses: a highlight
# column only turns yellow when that's true, not merely because it has
# a value (older confirmed figures now carry forward on every future
# download - see _load_master_rows - so "has a value" alone would
# wrongly re-highlight everything, every time); and the "movement as
# at" line sums only the rows where this is true, separately from the
# Total row's lifetime sum - see _write_table.
_CONFIRMED_THIS_RUN_KEY = {
    "full_payment_commission": "full_payment_confirmed_this_run",
    "installment_1_commission": "installment_1_confirmed_this_run",
    "installment_6_commission": "installment_6_confirmed_this_run",
}

# One blank column of separation, then the agency/agent split table
# starts here - see _write_agency_agent_split_columns.
_SPLIT_COLUMNS_START = len(_COLUMNS) + 2

# (trigger_type, real file's super-header text (its "Commissioin" typo
# preserved on purpose - confirmed from the real file, not a mistake
# here), FB-lead deduction %, agency %, agent %, row dict keys for the
# agency/agent amounts) - one entry per trigger, each contributing 3
# columns (FB deduction, Agency, Agent) to the split table.
_SPLIT_COLUMN_GROUPS = [
    ("full_payment", "Full Payment Commissioin (RM)",
     rules.AW_FB_LEAD_DEDUCTION_FULL_PAYMENT_PCT, rules.AW_AGENCY_FULL_PAYMENT_PCT, rules.AW_AGENT_FULL_PAYMENT_PCT,
     "full_payment_agency_amount", "full_payment_agent_amount"),
    ("installment_1", "First Half Commissioin (RM)",
     rules.AW_FB_LEAD_DEDUCTION_INSTALLMENT_PCT, rules.AW_AGENCY_INSTALLMENT_PCT, rules.AW_AGENT_INSTALLMENT_PCT,
     "installment_1_agency_amount", "installment_1_agent_amount"),
    ("installment_6", "Balance Half Commissioin (RM)",
     rules.AW_FB_LEAD_DEDUCTION_INSTALLMENT_PCT, rules.AW_AGENCY_INSTALLMENT_PCT, rules.AW_AGENT_INSTALLMENT_PCT,
     "installment_6_agency_amount", "installment_6_agent_amount"),
]

# Same per-trigger deduction % as _SPLIT_COLUMN_GROUPS above, keyed by
# trigger_type - used by the group-level Date Record summary (see
# _reconstruct_scoped_historical_entries and _load_summary_rows) to
# recompute the FB-lead deduction total for a cycle, the same way the
# main table's own FB-lead columns already do per row.
_DEDUCTION_PCT_BY_TRIGGER = {t: deduction_pct for t, _, deduction_pct, _, _, _, _ in _SPLIT_COLUMN_GROUPS}


def _pct_label(pct):
    """0.035 -> "3.5%", 0.07 -> "7%" - no trailing zeros, matching how
    the real file writes these percentages in its column headers."""
    return f"{pct * 100:g}%"


def _write_agency_agent_split_columns(sheet, rows, start_row, header_row, first_data_row, total_row, group_name):
    """
    Writes the extra columns the real file adds to the right of the
    main table for an agency that splits commission between agency and
    agent (currently AW Consultancy only - see
    agencies.commission_split_type). One 3-column group per trigger
    (FB-lead deduction, Agency %, Agent %); a row only gets values in
    the group matching whichever trigger(s) actually fired for it this
    run, exactly like the main table's own commission columns.

    Row positions are passed in rather than recomputed, so this stays
    perfectly aligned with whatever _write_table already wrote for the
    very same rows in the very same sheet.
    """
    col = _SPLIT_COLUMNS_START
    totals = {}
    movement_totals = {}
    for trigger_type, super_header, deduction_pct, agency_pct, agent_pct, agency_key, agent_key in _SPLIT_COLUMN_GROUPS:
        fb_col, agency_col, agent_col = col, col + 1, col + 2
        totals[fb_col] = 0.0
        totals[agency_col] = 0.0
        totals[agent_col] = 0.0
        movement_totals[agency_col] = 0.0
        movement_totals[agent_col] = 0.0
        confirmed_this_run_key = f"{trigger_type}_confirmed_this_run"

        title_cell = sheet.cell(row=start_row, column=fb_col, value=super_header)
        title_cell.font = _TITLE_FONT
        sheet.merge_cells(start_row=start_row, start_column=fb_col, end_row=start_row, end_column=agent_col)
        # merge_cells only keeps the top-left cell's style, but a
        # border still needs to be set on every cell in the merged
        # range for the border to actually render along its full width.
        for c in (fb_col, agency_col, agent_col):
            sheet.cell(row=start_row, column=c).border = _THIN_BORDER

        fb_label = f"{_pct_label(deduction_pct)} FB leads from XEKL  (to be deducted from {group_name})"
        agency_label = f"{group_name}\n{_pct_label(agency_pct)}"
        agent_label = f"Agent\n{_pct_label(agent_pct)}"
        for c, label in ((fb_col, fb_label), (agency_col, agency_label), (agent_col, agent_label)):
            cell = sheet.cell(row=header_row, column=c, value=label)
            cell.font = _HEADER_FONT
            cell.border = _THIN_BORDER

        row_num = first_data_row
        for row in rows:
            is_cancelled_row = row.get("status") in ("cancelled", "withdrawn")
            agency_amount = None if is_cancelled_row else row.get(agency_key)
            agent_amount = None if is_cancelled_row else row.get(agent_key)
            trigger_fired = agency_amount is not None or agent_amount is not None

            fb_cell = sheet.cell(row=row_num, column=fb_col)
            if trigger_fired:
                deduction = _fb_deduction_amount(row["net_price"], row["fb_lead_referred"], deduction_pct)
                fb_cell.value = deduction
                fb_cell.number_format = _MONEY_FORMAT
                if deduction is None:
                    fb_cell.fill = _GREY_FILL
                totals[fb_col] += deduction or 0.0

            agency_cell = sheet.cell(row=row_num, column=agency_col, value=agency_amount)
            agent_cell = sheet.cell(row=row_num, column=agent_col, value=agent_amount)
            for cell in (fb_cell, agency_cell, agent_cell):
                cell.font = _BODY_FONT
                cell.border = _THIN_BORDER
                if is_cancelled_row:
                    cell.fill = _BEIGE_FILL
            agency_cell.number_format = _MONEY_FORMAT
            agent_cell.number_format = _MONEY_FORMAT
            if not is_cancelled_row:
                totals[agency_col] += agency_amount or 0.0
                totals[agent_col] += agent_amount or 0.0
                if row.get(confirmed_this_run_key):
                    movement_totals[agency_col] += agency_amount or 0.0
                    movement_totals[agent_col] += agent_amount or 0.0
            row_num += 1

        col += 3

    for c, total in totals.items():
        cell = sheet.cell(row=total_row, column=c, value=round(total, 2))
        cell.font = _HEADER_FONT
        cell.number_format = _MONEY_FORMAT
        cell.border = _THIN_BORDER

    # Same "movement as at" idea as the main table (_write_table) - one
    # row under the lifetime Total, showing just what THIS run added to
    # the Agency/Agent split. This is where the real file's own
    # "movement as at {date}" line was actually first confirmed.
    movement_row = total_row + 1
    for c, amount in movement_totals.items():
        cell = sheet.cell(row=movement_row, column=c, value=round(amount, 2))
        cell.font = _BODY_FONT
        cell.number_format = _MONEY_FORMAT
        cell.fill = _YELLOW_FILL
        cell.border = _THIN_BORDER


def _fb_deduction_amount(net_price, fb_lead_referred, deduction_pct):
    """The RM amount deducted from the agency's share for an FB-lead
    referral, or None when this sale wasn't FB-lead-referred (the
    default - see contracts.fb_lead_referred). Recomputed here rather
    than stored, since commission_events only stores the post-deduction
    agency_amount, not the deduction itself."""
    if not fb_lead_referred:
        return None
    return commission._calculate_commission(net_price, deduction_pct)


def _write_table(sheet, rows, start_row, title, run_date, split_group_name=None):
    """Writes one titled table (header + data + bold total row, plus a
    "movement as at" line for whatever this specific run newly
    confirmed) starting at start_row. Returns the next free row, so
    tables can be stacked."""
    row_num = start_row
    title_cell = sheet.cell(row=row_num, column=1, value=title)
    title_cell.font = _TITLE_FONT
    row_num += 2

    header_row = row_num
    for col, (label, _key) in enumerate(_COLUMNS, start=1):
        cell = sheet.cell(row=header_row, column=col, value=label)
        cell.font = _HEADER_FONT
    row_num += 1
    first_data_row = row_num

    totals = {key: 0.0 for key in _TOTAL_KEYS}
    movement = {key: 0.0 for key in _CONFIRMED_THIS_RUN_KEY}
    for i, row in enumerate(rows, start=1):
        # A cancelled/withdrawn PO shades the whole row beige and its
        # commission figures are cleared, not just tinted - see
        # _COMMISSION_VALUE_KEYS. Checked first so it overrides green
        # below (a PO that was fully paid and *then* cancelled still
        # reads as "nothing owed now", the more important fact).
        is_cancelled_row = row.get("status") in ("cancelled", "withdrawn")
        # Green means the PO is fully paid off, commission-wise - either
        # a one-off full payment, or (for an installment plan) its
        # Balance Half, the last commission trigger that plan will ever
        # raise. Confirmed with the business: the customer may still be
        # paying off later installments after that (commission release
        # is fixed at installment 1 and 6 regardless of a 6/12/18/24
        # month plan's actual length - see rules.py), but nothing more
        # is ever due on this PO for commission purposes once its
        # Balance Half is paid, so it's "done" the same as a full
        # payment is.
        is_fully_paid_row = (not is_cancelled_row) and (
            row.get("full_payment_commission") is not None
            or row.get("installment_6_commission") is not None
        )
        for col, (label, key) in enumerate(_COLUMNS, start=1):
            if key == "row_no":
                value = i
            elif is_cancelled_row and key in _COMMISSION_VALUE_KEYS:
                value = None
            else:
                value = row.get(key)
            cell = sheet.cell(row=row_num, column=col, value=value)
            cell.font = _BODY_FONT
            if label in _MONEY_COLUMNS:
                cell.number_format = _MONEY_FORMAT
            if is_cancelled_row:
                cell.fill = _BEIGE_FILL
            elif is_fully_paid_row:
                cell.fill = _GREEN_FILL
            # Checked with a separate `if`, not `elif` - a PO whose
            # first-ever import already has both full payment AND an
            # instalment due (both triggers are evaluated independently
            # in app/commission.py, with no rule against both firing at
            # once) must still show its instalment cell in yellow, not
            # let the row's green silently swallow it. Gated on the
            # confirmed_this_run flag, not merely "has a value" - a
            # figure confirmed in an earlier run now carries forward
            # plainly on every later download, not re-highlighted every
            # single time (see _load_master_rows).
            if (
                not is_cancelled_row
                and label in _HIGHLIGHT_COLUMNS
                and row.get(_CONFIRMED_THIS_RUN_KEY.get(key), False)
            ):
                cell.fill = _YELLOW_FILL
        for key in totals:
            if is_cancelled_row and key in _COMMISSION_VALUE_KEYS:
                continue
            totals[key] += row.get(key) or 0.0
        if not is_cancelled_row:
            for key, confirmed_key in _CONFIRMED_THIS_RUN_KEY.items():
                if row.get(confirmed_key):
                    movement[key] += row.get(key) or 0.0
        row_num += 1

    # "Total" label sits under Customer Name - well clear of every
    # column that now carries its own numeric subtotal (Niche/Tablet
    # Price, Promotion, Discount, Nett Price, and each of the three
    # commission columns), each directly beneath its own header rather
    # than one merged figure - matches how the original file separates
    # these out rather than lumping them into a single number. This is
    # the lifetime total across every confirmed run ever, not just this
    # one - see the "movement" line right under it for that.
    sheet.cell(row=row_num, column=_COLUMN_INDEX["customer_name"], value="Total").font = _HEADER_FONT
    for key in totals:
        cell = sheet.cell(row=row_num, column=_COLUMN_INDEX[key], value=round(totals[key], 2))
        cell.font = _HEADER_FONT
        cell.number_format = _MONEY_FORMAT
    row_num += 1

    # "movement as at {date}" - confirmed against the real file (the
    # AW Consultancy split table's own version of this line): a second
    # figure under the lifetime Total, showing just what THIS run added
    # - the number Accounts actually needs for this cycle's payout
    # requisition, as opposed to the running lifetime total above it.
    as_at = run_date.isoformat() if isinstance(run_date, datetime.date) else run_date
    movement_label = f"movement as at {_format_short_date(as_at)}"
    label_cell = sheet.cell(row=row_num, column=_COLUMN_INDEX["customer_name"], value=movement_label)
    label_cell.font = _BODY_FONT
    for key, amount in movement.items():
        cell = sheet.cell(row=row_num, column=_COLUMN_INDEX[key], value=round(amount, 2))
        cell.font = _BODY_FONT
        cell.number_format = _MONEY_FORMAT
        cell.fill = _YELLOW_FILL
    row_num += 1

    if split_group_name is not None:
        _write_agency_agent_split_columns(
            sheet, rows, start_row, header_row, first_data_row, row_num - 2, split_group_name
        )

    row_num += 1

    return row_num + 1  # one blank row before whatever comes next


def _write_summary_table(sheet, summary_rows, start_row, current_run_id, split_group_name=None):
    """
    `current_run_id` identifies the commission_run this download is
    actually for - whichever row was added by this exact run (matched
    on run_id, set by _load_summary_rows) gets shaded yellow, the same
    "what did THIS run add" highlighting convention used everywhere
    else in this file (the main table's own "movement as at" row, the
    Agency/Agent split columns' movement row). Confirmed against a
    real reference table: the grand total row underneath is shaded
    yellow too, whenever this run actually contributed a row - not
    unconditionally, so a download for a run that only confirmed
    something historical (no new "As at" row at all) leaves the grand
    total plain like every other row.

    `split_group_name`: only ever passed for a split-type agency
    GROUP's own combined sheet (e.g. "AW Consultancy" - never a flat
    agency's sheet, and never an individual agent's own sheet, whose
    main table already shows just that agent's own cut - see
    _load_master_rows's historical fallback and
    _reconstruct_scoped_historical_entries). Confirmed against a real
    reference table: this sheet's Date Record needs 3 extra columns
    between Running Total and Remarks - "Less Agent Commission",
    "Less FB leads from XEKL", and "Total for {group_name}" (Running
    Total minus both deductions) - since Running Total here is the
    combined agency+agent figure, and the business needs to see what's
    actually left for the agency itself after the agent's cut and any
    FB-lead deduction come out.
    """
    row_num = start_row
    sheet.cell(row=row_num, column=1, value="Summary").font = _TITLE_FONT
    row_num += 2

    if split_group_name is not None:
        columns = [
            "DATE RECORD", "Full Commission", "First Half Commission", "Second Half Commission",
            "Running Total", "Less Agent Commission", "Less FB leads from XEKL",
            f"Total for {split_group_name}", "Remarks",
        ]
    else:
        columns = _SUMMARY_COLUMNS

    header_row = row_num
    for col, label in enumerate(columns, start=1):
        sheet.cell(row=header_row, column=col, value=label).font = _HEADER_FONT
    row_num += 1

    money_cols = {
        "Full Commission", "First Half Commission", "Second Half Commission", "Running Total",
        "Less Agent Commission", "Less FB leads from XEKL",
    }
    for row in summary_rows:
        if split_group_name is not None:
            values = [
                row["date_record"], row["full_commission"], row["first_half_commission"],
                row["second_half_commission"], row["running_total"],
                row["agent_commission"], row["fb_lead_deduction"], row["total_for_group"],
                row["remarks"],
            ]
        else:
            values = [
                row["date_record"], row["full_commission"], row["first_half_commission"],
                row["second_half_commission"], row["running_total"], row["remarks"],
            ]
        is_current_run = current_run_id is not None and row["run_id"] == current_run_id
        for col, (label, value) in enumerate(zip(columns, values), start=1):
            cell = sheet.cell(row=row_num, column=col, value=value)
            cell.font = _BODY_FONT
            if label in money_cols or label.startswith("Total for "):
                cell.number_format = _MONEY_FORMAT
            if is_current_run:
                cell.fill = _YELLOW_FILL
        row_num += 1

    if summary_rows:
        # The grand total across every cycle ever recorded - unlike
        # each row's own "running_total" above, this one genuinely is
        # cumulative, so it's summed fresh here rather than read off
        # the last row.
        final_running_total = round(sum(row["running_total"] for row in summary_rows), 2)
        latest_date_label = summary_rows[-1]["date_record"].replace("As at ", "")
        # Whether THIS run actually added a row to this table (a run
        # confirmed on this download but scoped away from this sheet,
        # or one that only confirmed something already historically
        # accounted for, adds no row here at all) - only then does the
        # grand total line also shade yellow.
        grand_total_is_current = any(
            current_run_id is not None and row["run_id"] == current_run_id for row in summary_rows
        )
        label_cell = sheet.cell(row=row_num, column=1, value=f"Total Sum of Commission Payout as at {latest_date_label}")
        label_cell.font = _HEADER_FONT
        total_cell = sheet.cell(row=row_num, column=5, value=final_running_total)
        total_cell.font = _HEADER_FONT
        total_cell.number_format = _MONEY_FORMAT
        if grand_total_is_current:
            label_cell.fill = _YELLOW_FILL
            total_cell.fill = _YELLOW_FILL
        if split_group_name is not None:
            final_agent_commission = round(sum(row["agent_commission"] for row in summary_rows), 2)
            final_fb_lead_deduction = round(sum(row["fb_lead_deduction"] for row in summary_rows), 2)
            # Not minus final_fb_lead_deduction too - see the matching
            # comment on total_for_group above.
            final_total_for_group = round(final_running_total - final_agent_commission, 2)
            for col, value in ((6, final_agent_commission), (7, final_fb_lead_deduction), (8, final_total_for_group)):
                cell = sheet.cell(row=row_num, column=col, value=value)
                cell.font = _HEADER_FONT
                cell.number_format = _MONEY_FORMAT
                if grand_total_is_current:
                    cell.fill = _YELLOW_FILL
        row_num += 1

    return row_num + 1


def _autosize_columns(sheet, column_count):
    for col in range(1, column_count + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 20


def generate_commission_run_report(conn, commission_run_id, output_path):
    """
    Writes the downloadable Excel file for one commission run. Every
    sheet gets its own Date Record summary table directly under its
    data, scoped to that sheet's own history (confirmed against the
    real file: XEMP's own sheet totals RM33,070.50, a genuine subset
    of the "All" sheet's RM65,493.00, not a separate figure) - not just
    one summary table on "All" covering everything:
      - "All" sheet: every newly-due PO in one list (the Master view),
        followed by the Date Record summary covering every run ever
        processed
      - one combined sheet per agency GROUP (everyone sharing that
        group's rows together), followed by a Date Record scoped to
        just that group
      - for any group with splits_by_agent on, one additional
        standalone sheet per individual agent, after the group sheet,
        each with its own Date Record scoped to just that agent

    Raises ValueError if commission_run_id is None (nothing was newly
    due in this upload), or if this run exists but nothing on it has
    been confirmed yet (everything detected is still sitting on the
    review page) - either way there is nothing meaningful to export,
    and that should be a clear error rather than a blank file quietly
    handed to Accounts.
    """
    if commission_run_id is None:
        raise ValueError(
            "No commission run to report - nothing was newly due in this "
            "upload, so there's nothing to export."
        )

    run_row = conn.execute(
        "SELECT run_date FROM commission_runs WHERE id = ?", (commission_run_id,)
    ).fetchone()
    run_date = run_row["run_date"]

    # The report itself always shows every contract (see
    # _load_master_rows), so an empty result there would only mean an
    # empty database - the real gate is whether THIS run actually
    # confirmed anything, since that's what would make a fresh
    # download meaningfully different from the last one.
    has_confirmed_this_run = conn.execute(
        "SELECT 1 FROM commission_events WHERE commission_run_id = ? AND status = 'confirmed' LIMIT 1",
        (commission_run_id,),
    ).fetchone()
    if not has_confirmed_this_run:
        raise ValueError(
            "Nothing confirmed on this commission run yet - review and confirm "
            "the detected commissions before downloading."
        )

    rows = _load_master_rows(conn, commission_run_id, run_date)
    column_count = len(_COLUMNS)

    workbook = Workbook()
    all_sheet = workbook.active
    all_sheet.title = "All"
    next_row = _write_table(all_sheet, rows, start_row=1, title=_title_line(COMPANY_SHORT_NAME, rows, run_date), run_date=run_date)
    _write_summary_table(all_sheet, _load_summary_rows(conn), start_row=next_row, current_run_id=commission_run_id)
    _autosize_columns(all_sheet, column_count)

    # Top-level grouping is by agency_group, not raw agency_code - an
    # agency with several sub-codes (AW Consultancy's AC108-01/-02/-03)
    # gets ONE combined sheet with every code's rows together first,
    # matching the real file's actual structure. An agency with no
    # group set (every agency except AW Consultancy so far) falls back
    # to its own agency_code as its "group", so it still gets exactly
    # one sheet, same as before this existed.
    groups = {}
    for row in rows:
        groups.setdefault(row["agency_group"], []).append(row)

    # "(No Agency)" is a fallback bucket, not a real agency - moved to
    # the very end of the sheet order (regardless of where its POs
    # happen to fall by PO No) so it reads as the leftover/miscellany
    # tab it is, rather than being interleaved with real agencies.
    if "(No Agency)" in groups:
        groups["(No Agency)"] = groups.pop("(No Agency)")

    split_column_count = _SPLIT_COLUMNS_START + len(_SPLIT_COLUMN_GROUPS) * 3 - 1

    used_titles = {"All"}
    for group_name, group_rows in groups.items():
        sheet_title = _unique_sheet_title(group_name, used_titles)
        used_titles.add(sheet_title)
        sheet = workbook.create_sheet(sheet_title)
        # AW Consultancy (and any future agency seeded the same way -
        # see rules.AGENCIES_WITH_AGENCY_AGENT_SPLIT) gets the extra
        # agency/agent split columns to the right of the main table;
        # every other agency's sheet stays exactly as before.
        is_split_group = any(row["commission_split_type"] == "agency_agent_split" for row in group_rows)
        group_next_row = _write_table(
            sheet, group_rows, start_row=1, title=_title_line(group_name, group_rows, run_date), run_date=run_date,
            split_group_name=group_name if is_split_group else None,
        )
        _write_summary_table(
            sheet, _load_summary_rows(conn, agency_group=group_name),
            start_row=group_next_row, current_run_id=commission_run_id,
            split_group_name=group_name if is_split_group else None,
        )
        _autosize_columns(sheet, split_column_count if is_split_group else column_count)

        # If any code in this group splits by agent, ALSO create a
        # separate standalone sheet per individual agent - not a
        # sub-section of the combined sheet, a genuinely separate sheet
        # in the workbook, exactly like the real file's sheet list
        # (one "AW Consultancy" sheet, then "TAN HER JIE", "HOO CHEW
        # YOON", etc. each as their own tab).
        if any(row["splits_by_agent"] for row in group_rows):
            agents = {}
            for row in group_rows:
                agents.setdefault(row["agent_name"], []).append(row)
            for agent_name, agent_rows in agents.items():
                agent_sheet_title = _unique_sheet_title(agent_name, used_titles)
                used_titles.add(agent_sheet_title)
                agent_sheet = workbook.create_sheet(agent_sheet_title)
                # The Agency/Agent split columns show how one PO's
                # commission divides between the agency and the agent -
                # that only makes sense on the combined agency-group
                # sheet above, where both shares are visible side by
                # side. An individual agent's own sheet never gets the
                # split table, regardless of whether their agency uses
                # agency_agent_split - confirmed with the business: only
                # the agency-level view computes the split, the agent
                # view just shows what's owed to them.
                is_split_agent = any(row["commission_split_type"] == "agency_agent_split" for row in agent_rows)
                if is_split_agent:
                    # "What's owed to them" means just the agent's own
                    # cut (full_payment_agent_amount etc.), not the
                    # agency+agent combined total the group sheet above
                    # shows - shallow copies, so this never mutates the
                    # same row dicts the group sheet already rendered
                    # from.
                    display_rows = [
                        {
                            **row,
                            "full_payment_commission": row["full_payment_agent_amount"],
                            "installment_1_commission": row["installment_1_agent_amount"],
                            "installment_6_commission": row["installment_6_agent_amount"],
                        }
                        for row in agent_rows
                    ]
                else:
                    display_rows = agent_rows
                agent_next_row = _write_table(
                    agent_sheet, display_rows, start_row=1, title=_title_line(agent_name, agent_rows, run_date),
                    run_date=run_date, split_group_name=None,
                )
                _write_summary_table(
                    agent_sheet, _load_summary_rows(conn, agency_group=group_name, agent_name=agent_name),
                    start_row=agent_next_row, current_run_id=commission_run_id,
                )
                _autosize_columns(agent_sheet, column_count)

    workbook.save(output_path)
    return output_path


def _latest_confirmed_run_for_period(conn, period_start, period_end):
    """
    The single commission_run_id that most recently confirmed
    something for an in-period PO (scoped by po_date alone, never by
    agency or agent), or None if nothing has been confirmed yet for
    this period at all.

    A period report can be regenerated at any time and combines
    however many separate runs/uploads happened to land in it - there's
    no single "the run that was just processed" the way a per-run
    download has. This is the closest useful equivalent, and it's
    computed ONCE for the whole workbook rather than per sheet: "round
    2 of processing" is one moment in time company-wide, not a
    separate concept per agency, so the All sheet, every agency group,
    and every agent all highlight the exact same cells as new. Used for
    both the main table's yellow cell highlight (_load_master_rows'
    latest_run_id) and every sheet's Date Record summary table
    (_write_summary_table's current_run_id) - see generate_period_report.
    """
    row = conn.execute(
        """
        SELECT e.commission_run_id AS run_id
        FROM commission_events e
        JOIN commission_runs r ON r.id = e.commission_run_id
        JOIN contracts c ON c.po_no = e.po_no
        WHERE e.status = 'confirmed' AND c.po_date BETWEEN ? AND ?
        ORDER BY r.run_date DESC, r.id DESC
        LIMIT 1
        """,
        (period_start, period_end),
    ).fetchone()
    return row["run_id"] if row is not None else None


def _period_title(label, period_start, period_end, processed_date):
    """
    Unlike the standing per-run report (title "AS AT {run_date}" - see
    _title_line), a period report has no single run date: it can be
    regenerated at any time and combines however many separate
    uploads/confirmations happened to land in the period, picking up
    later confirmations each time it's re-downloaded. So the period
    itself (the PO purchase-date range) goes in the title as before,
    plus a second line stating the actual date THIS copy was generated
    - since the same period, downloaded again next week, may show more
    confirmed commission than it does today.
    """
    return (
        f"{label} OVERALL COMMISSION PAYOUT "
        f"({_format_title_date(period_start)} TO {_format_title_date(period_end)})"
        f"\nPROCESSED AS OF {_format_title_date(processed_date)}"
    )


def generate_period_report(conn, period_start, period_end, output_path):
    """
    Writes the downloadable Excel file for a chosen period (ISO date
    strings, inclusive both ends) - the sibling of
    generate_commission_run_report, same sheet structure (an "All"
    sheet, one combined sheet per agency group, one standalone sheet
    per agent where the group splits by agent), but scoped by PO
    purchase date (po_date) instead of tied to one specific upload's
    commission_run_id.

    Confirmed with the business: "August's report" means every PO
    PURCHASED in August, full stop - not every PO with something
    confirmed in August. A June-purchased PO whose installment happens
    to get confirmed in August stays on JUNE's report (re-download it
    any time to pick up a late-confirmed payment), never August's,
    even though a commission event with an August trigger_date exists
    for it. Every commission event ever confirmed for an in-period PO
    shows here, regardless of when it was confirmed - this is "show
    whatever's confirmed so far for this month's contracts", the same
    logic the standing per-run report uses, just pre-filtered to one
    month's own POs.

    Deliberately does NOT include the pre-tool historical-absorption
    cutoff money (see _load_master_rows/_load_summary_rows) - by
    agreement, kept on the standing per-run report only rather than
    also reconstructing a per-month breakdown of that lump-sum figure.

    Raises ValueError if no PO was purchased in this range at all, or
    none of them have anything confirmed yet - same "nothing
    meaningful to export" reasoning as generate_commission_run_report.
    """
    latest_run_id = _latest_confirmed_run_for_period(conn, period_start, period_end)
    rows = _load_master_rows(conn, commission_run_id=None, run_date=period_end,
                              period_start=period_start, period_end=period_end,
                              latest_run_id=latest_run_id)
    if not rows:
        raise ValueError(
            f"No PO was purchased between {period_start} and {period_end} - "
            f"there's nothing to export for this period."
        )
    column_count = len(_COLUMNS)
    processed_date = datetime.date.today().isoformat()

    workbook = Workbook()
    all_sheet = workbook.active
    all_sheet.title = "All"
    all_title = _period_title(COMPANY_SHORT_NAME, period_start, period_end, processed_date)
    next_row = _write_table(all_sheet, rows, start_row=1, title=all_title, run_date=period_end)
    all_summary_rows = _load_summary_rows(conn, period_start=period_start, period_end=period_end)
    _write_summary_table(
        all_sheet, all_summary_rows, start_row=next_row, current_run_id=latest_run_id,
    )
    _autosize_columns(all_sheet, column_count)

    groups = {}
    for row in rows:
        groups.setdefault(row["agency_group"], []).append(row)
    if "(No Agency)" in groups:
        groups["(No Agency)"] = groups.pop("(No Agency)")

    split_column_count = _SPLIT_COLUMNS_START + len(_SPLIT_COLUMN_GROUPS) * 3 - 1

    used_titles = {"All"}
    for group_name, group_rows in groups.items():
        sheet_title = _unique_sheet_title(group_name, used_titles)
        used_titles.add(sheet_title)
        sheet = workbook.create_sheet(sheet_title)
        is_split_group = any(row["commission_split_type"] == "agency_agent_split" for row in group_rows)
        group_title = _period_title(group_name, period_start, period_end, processed_date)
        group_next_row = _write_table(
            sheet, group_rows, start_row=1, title=group_title, run_date=period_end,
            split_group_name=group_name if is_split_group else None,
        )
        group_summary_rows = _load_summary_rows(
            conn, agency_group=group_name, period_start=period_start, period_end=period_end,
        )
        _write_summary_table(
            sheet, group_summary_rows, start_row=group_next_row,
            current_run_id=latest_run_id,
            split_group_name=group_name if is_split_group else None,
        )
        _autosize_columns(sheet, split_column_count if is_split_group else column_count)

        if any(row["splits_by_agent"] for row in group_rows):
            agents = {}
            for row in group_rows:
                agents.setdefault(row["agent_name"], []).append(row)
            for agent_name, agent_rows in agents.items():
                agent_sheet_title = _unique_sheet_title(agent_name, used_titles)
                used_titles.add(agent_sheet_title)
                agent_sheet = workbook.create_sheet(agent_sheet_title)
                is_split_agent = any(row["commission_split_type"] == "agency_agent_split" for row in agent_rows)
                if is_split_agent:
                    display_rows = [
                        {
                            **row,
                            "full_payment_commission": row["full_payment_agent_amount"],
                            "installment_1_commission": row["installment_1_agent_amount"],
                            "installment_6_commission": row["installment_6_agent_amount"],
                        }
                        for row in agent_rows
                    ]
                else:
                    display_rows = agent_rows
                agent_title = _period_title(agent_name, period_start, period_end, processed_date)
                agent_next_row = _write_table(
                    agent_sheet, display_rows, start_row=1, title=agent_title,
                    run_date=period_end, split_group_name=None,
                )
                agent_summary_rows = _load_summary_rows(
                    conn, agency_group=group_name, agent_name=agent_name,
                    period_start=period_start, period_end=period_end,
                )
                _write_summary_table(
                    agent_sheet, agent_summary_rows, start_row=agent_next_row,
                    current_run_id=latest_run_id,
                )
                _autosize_columns(agent_sheet, column_count)

    workbook.save(output_path)
    return output_path
