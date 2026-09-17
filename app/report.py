"""
Builds the downloadable Excel report for one commission run - the file
that gets handed to Accounts.

Column layout deliberately mirrors the real Kenjin Master Report
(docs/data_model.md section 2) rather than a simplified summary: one
row per PO, with the Full/1st Half/Balance Half columns each filled in
only when that specific trigger fired in this run - so a PO with two
triggers due in the same run (rare, but real - see
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
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from app import commission, rules

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

# The real file also shades cancelled/withdrawn rows a light beige
# (theme accent2 #ED7D31 tinted 0.8, ~#FBE5D6) - confirmed, but NOT
# implemented yet: this report only lists POs with commission newly
# due this run, so a cancelled PO (nothing due) never appears as a row
# at all. There's nothing to paint beige until/unless the report
# becomes a full status listing rather than a due-items list - that's
# a bigger design question, not a missing color constant.

# Marks the FB-lead-deduction cell in the agency/agent split columns
# (see _write_agency_agent_split_columns) when a row's relevant trigger
# fired but fb_lead_referred is off, i.e. "this slot exists, nothing
# was deducted here". Approximated from a screenshot, not an actual
# .xlsx file this time (unlike green/yellow/beige, which were matched
# against real file bytes) - flag if the shade needs adjusting once
# there's a real file with an FB-lead deduction to check against.
_GREY_FILL = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

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


def _load_run_rows(conn, commission_run_id, run_date):
    """
    Joins this run's commission_events against contracts/customers/
    agencies and collapses them to one row per PO - a PO with two
    triggers due in the same run ends up with both sets of columns
    filled on a single row, not two separate rows.

    `run_date` is needed to compute each row's Cooling Off Period
    status (EXPIRED once COOLING_OFF_TOTAL_DAYS have passed since
    Signature Date - see _cooling_off_status).
    """
    cursor = conn.execute(
        """
        SELECT
            c.po_no, c.po_date, c.signature_date, c.customer_id, c.lot_no,
            c.niche_price, c.promotion, c.discount, c.net_price,
            c.full_settlement_paid_date, c.first_installment_paid_date,
            c.sixth_installment_paid_date, c.agent_name, c.agency_code, c.remarks,
            c.fb_lead_referred,
            cu.name AS customer_name,
            a.splits_by_agent, a.agency_group, a.commission_split_type,
            e.trigger_type, e.amount, e.agency_amount, e.agent_amount
        FROM commission_events e
        JOIN contracts c ON c.po_no = e.po_no
        LEFT JOIN customers cu ON cu.customer_id = c.customer_id
        LEFT JOIN agencies a ON a.agency_code = c.agency_code
        WHERE e.commission_run_id = ?
        ORDER BY c.agency_code, c.agent_name, c.po_no
        """,
        (commission_run_id,),
    )

    by_po = {}
    for r in cursor.fetchall():
        po_no = r["po_no"]
        if po_no not in by_po:
            by_po[po_no] = {
                "po_no": po_no,
                "po_date": r["po_date"],
                "signature_date": r["signature_date"],
                "customer_id": r["customer_id"],
                "customer_name": r["customer_name"],
                "lot_no": r["lot_no"],
                "niche_price": r["niche_price"],
                "promotion": r["promotion"],
                "discount": r["discount"],
                "net_price": r["net_price"],
                "cooling_off_period": _cooling_off_status(r["signature_date"], run_date),
                "full_settlement_paid_date": None,
                "full_payment_commission": None,
                "full_commission_paid_date": None,  # filled in later by Accounts, never by this tool
                "first_installment_paid_date": None,
                "installment_1_commission": None,
                "installment_1_commission_paid_date": None,  # ditto
                "sixth_installment_paid_date": None,
                "installment_6_commission": None,
                "installment_6_commission_paid_date": None,  # ditto
                "agent_name": r["agent_name"] or "(unassigned)",
                "agency_code": r["agency_code"] or "(No Agency)",
                # Several agency_codes can share one real-world agency
                # (AW Consultancy's AC108-01/-02/-03) - agency_group is
                # the label the combined top-level sheet groups under.
                # Falls back to the raw agency_code when no group is
                # set, so every other agency still gets its own single
                # sheet exactly as before this existed.
                "agency_group": r["agency_group"] or r["agency_code"] or "(No Agency)",
                "remarks": r["remarks"],
                # No agency on file -> nothing to group by agent for
                # either; default to a flat listing rather than
                # splitting by agent.
                "splits_by_agent": bool(r["splits_by_agent"]) if r["agency_code"] else False,
                "commission_split_type": r["commission_split_type"] or "flat",
                "fb_lead_referred": bool(r["fb_lead_referred"]),
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
        row = by_po[po_no]
        if r["trigger_type"] == "full_payment":
            row["full_settlement_paid_date"] = r["full_settlement_paid_date"]
            row["full_payment_commission"] = r["amount"]
            row["full_payment_agency_amount"] = r["agency_amount"]
            row["full_payment_agent_amount"] = r["agent_amount"]
        elif r["trigger_type"] == "installment_1":
            row["first_installment_paid_date"] = r["first_installment_paid_date"]
            row["installment_1_commission"] = r["amount"]
            row["installment_1_agency_amount"] = r["agency_amount"]
            row["installment_1_agent_amount"] = r["agent_amount"]
        elif r["trigger_type"] == "installment_6":
            row["sixth_installment_paid_date"] = r["sixth_installment_paid_date"]
            row["installment_6_commission"] = r["amount"]
            row["installment_6_agency_amount"] = r["agency_amount"]
            row["installment_6_agent_amount"] = r["agent_amount"]

    return list(by_po.values())


def _load_summary_rows(conn):
    """
    Every commission run ever processed (not just this one), grouped
    and pivoted into the "Date Record" running-total table from the
    original workflow. Built entirely from commission_events/
    commission_runs, which already record everything needed - this is
    a new view over existing data, not new calculation logic.
    """
    cursor = conn.execute(
        """
        SELECT r.id AS run_id, r.run_date, e.trigger_type, SUM(e.amount) AS total
        FROM commission_runs r
        JOIN commission_events e ON e.commission_run_id = r.id
        GROUP BY r.id, e.trigger_type
        ORDER BY r.run_date, r.id
        """
    )

    by_run = {}
    run_order = []
    for row in cursor.fetchall():
        run_id = row["run_id"]
        if run_id not in by_run:
            by_run[run_id] = {"run_date": row["run_date"], "full_payment": 0.0, "installment_1": 0.0, "installment_6": 0.0}
            run_order.append(run_id)
        by_run[run_id][row["trigger_type"]] = row["total"] or 0.0

    running_total = 0.0
    summary_rows = []
    for run_id in run_order:
        r = by_run[run_id]
        run_total = r["full_payment"] + r["installment_1"] + r["installment_6"]
        running_total += run_total
        summary_rows.append({
            "date_record": f"As at {_format_short_date(r['run_date'])}",
            "full_commission": r["full_payment"],
            "first_half_commission": r["installment_1"],
            "second_half_commission": r["installment_6"],
            "running_total": running_total,
            "remarks": None,
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
_TOTAL_KEYS = ("full_payment_commission", "installment_1_commission", "installment_6_commission")

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
    for trigger_type, super_header, deduction_pct, agency_pct, agent_pct, agency_key, agent_key in _SPLIT_COLUMN_GROUPS:
        fb_col, agency_col, agent_col = col, col + 1, col + 2
        totals[fb_col] = 0.0
        totals[agency_col] = 0.0
        totals[agent_col] = 0.0

        sheet.cell(row=start_row, column=fb_col, value=super_header).font = _TITLE_FONT
        sheet.merge_cells(start_row=start_row, start_column=fb_col, end_row=start_row, end_column=agent_col)

        fb_label = f"{_pct_label(deduction_pct)} FB leads from XEKL  (to be deducted from {group_name})"
        agency_label = f"{group_name}\n{_pct_label(agency_pct)}"
        agent_label = f"Agent\n{_pct_label(agent_pct)}"
        for c, label in ((fb_col, fb_label), (agency_col, agency_label), (agent_col, agent_label)):
            cell = sheet.cell(row=header_row, column=c, value=label)
            cell.font = _HEADER_FONT

        row_num = first_data_row
        for row in rows:
            agency_amount = row.get(agency_key)
            agent_amount = row.get(agent_key)
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
            agency_cell.number_format = _MONEY_FORMAT
            agent_cell.number_format = _MONEY_FORMAT
            totals[agency_col] += agency_amount or 0.0
            totals[agent_col] += agent_amount or 0.0
            row_num += 1

        col += 3

    for c, total in totals.items():
        cell = sheet.cell(row=total_row, column=c, value=round(total, 2))
        cell.font = _HEADER_FONT
        cell.number_format = _MONEY_FORMAT


def _fb_deduction_amount(net_price, fb_lead_referred, deduction_pct):
    """The RM amount deducted from the agency's share for an FB-lead
    referral, or None when this sale wasn't FB-lead-referred (the
    default - see contracts.fb_lead_referred). Recomputed here rather
    than stored, since commission_events only stores the post-deduction
    agency_amount, not the deduction itself."""
    if not fb_lead_referred:
        return None
    return commission._calculate_commission(net_price, deduction_pct)


def _write_table(sheet, rows, start_row, title, split_group_name=None):
    """Writes one titled table (header + data + bold total row) starting
    at start_row. Returns the next free row, so tables can be stacked."""
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
    for i, row in enumerate(rows, start=1):
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
        is_fully_paid_row = (
            row.get("full_payment_commission") is not None
            or row.get("installment_6_commission") is not None
        )
        for col, (label, key) in enumerate(_COLUMNS, start=1):
            value = i if key == "row_no" else row.get(key)
            cell = sheet.cell(row=row_num, column=col, value=value)
            cell.font = _BODY_FONT
            if label in _MONEY_COLUMNS:
                cell.number_format = _MONEY_FORMAT
            if is_fully_paid_row:
                cell.fill = _GREEN_FILL
            # Checked with a separate `if`, not `elif` - a PO whose
            # first-ever import already has both full payment AND an
            # instalment due (both triggers are evaluated independently
            # in app/commission.py, with no rule against both firing at
            # once) must still show its instalment cell in yellow, not
            # let the row's green silently swallow it.
            if label in _HIGHLIGHT_COLUMNS and value is not None:
                cell.fill = _YELLOW_FILL
        for key in totals:
            totals[key] += row.get(key) or 0.0
        row_num += 1

    # "Total" label sits under Nett Price (the column immediately left
    # of the three commission columns); each commission column gets its
    # own subtotal directly beneath it, rather than one merged figure -
    # matches how the original file separates Full / First Half /
    # Second Half rather than lumping them into a single number.
    sheet.cell(row=row_num, column=_COLUMN_INDEX["net_price"], value="Total").font = _HEADER_FONT
    for key in totals:
        cell = sheet.cell(row=row_num, column=_COLUMN_INDEX[key], value=round(totals[key], 2))
        cell.font = _HEADER_FONT
        cell.number_format = _MONEY_FORMAT

    if split_group_name is not None:
        _write_agency_agent_split_columns(
            sheet, rows, start_row, header_row, first_data_row, row_num, split_group_name
        )

    row_num += 1

    return row_num + 1  # one blank row before whatever comes next


def _write_summary_table(sheet, summary_rows, start_row):
    row_num = start_row
    sheet.cell(row=row_num, column=1, value="Summary").font = _TITLE_FONT
    row_num += 2

    header_row = row_num
    for col, label in enumerate(_SUMMARY_COLUMNS, start=1):
        sheet.cell(row=header_row, column=col, value=label).font = _HEADER_FONT
    row_num += 1

    money_cols = {"Full Commission", "First Half Commission", "Second Half Commission", "Running Total"}
    for row in summary_rows:
        values = [
            row["date_record"], row["full_commission"], row["first_half_commission"],
            row["second_half_commission"], row["running_total"], row["remarks"],
        ]
        for col, (label, value) in enumerate(zip(_SUMMARY_COLUMNS, values), start=1):
            cell = sheet.cell(row=row_num, column=col, value=value)
            cell.font = _BODY_FONT
            if label in money_cols:
                cell.number_format = _MONEY_FORMAT
        row_num += 1

    if summary_rows:
        final_running_total = summary_rows[-1]["running_total"]
        latest_date_label = summary_rows[-1]["date_record"].replace("As at ", "")
        label_cell = sheet.cell(row=row_num, column=1, value=f"Total Sum of Commission Payout as at {latest_date_label}")
        label_cell.font = _HEADER_FONT
        total_cell = sheet.cell(row=row_num, column=5, value=final_running_total)
        total_cell.font = _HEADER_FONT
        total_cell.number_format = _MONEY_FORMAT
        row_num += 1

    return row_num + 1


def _autosize_columns(sheet, column_count):
    for col in range(1, column_count + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 20


def generate_commission_run_report(conn, commission_run_id, output_path):
    """
    Writes the downloadable Excel file for one commission run:
      - "All" sheet: every newly-due PO in one list (the Master view),
        followed by the cumulative Date Record summary table covering
        every run ever processed
      - one combined sheet per agency GROUP (everyone sharing that
        group's rows together)
      - for any group with splits_by_agent on, one additional
        standalone sheet per individual agent, after the group sheet

    Raises ValueError if commission_run_id is None (nothing was newly
    due in this upload) - there is nothing meaningful to export, and
    that should be a clear error rather than a blank file quietly
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

    rows = _load_run_rows(conn, commission_run_id, run_date)
    column_count = len(_COLUMNS)

    workbook = Workbook()
    all_sheet = workbook.active
    all_sheet.title = "All"
    next_row = _write_table(all_sheet, rows, start_row=1, title=_title_line(COMPANY_SHORT_NAME, rows, run_date))
    _write_summary_table(all_sheet, _load_summary_rows(conn), start_row=next_row)
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
        _write_table(
            sheet, group_rows, start_row=1, title=_title_line(group_name, group_rows, run_date),
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
                is_split_agent = any(row["commission_split_type"] == "agency_agent_split" for row in agent_rows)
                _write_table(
                    agent_sheet, agent_rows, start_row=1, title=_title_line(agent_name, agent_rows, run_date),
                    split_group_name=group_name if is_split_agent else None,
                )
                _autosize_columns(agent_sheet, split_column_count if is_split_agent else column_count)

    workbook.save(output_path)
    return output_path
