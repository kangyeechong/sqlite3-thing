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

_BODY_FONT = Font(name="Arial", size=11)
_HEADER_FONT = Font(name="Arial", size=11, bold=True)
_TITLE_FONT = Font(name="Arial", size=13, bold=True)
_MONEY_FORMAT = "#,##0.00"

# Green matched against the real sample file's actual cell formatting
# (not guessed): full-payment rows there use theme accent6 (#70AD47)
# tinted 0.6, reproduced here as plain RGB (Excel's tint formula
# applied by hand) since openpyxl's fill doesn't need to reference the
# workbook's theme to look the same.
_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
_GREEN_FILL = PatternFill(start_color="C6DEB5", end_color="C6DEB5", fill_type="solid")   # full payment rows

# The real file also shades cancelled/withdrawn rows a light beige
# (theme accent2 #ED7D31 tinted 0.8, ~#FBE5D6) - confirmed, but NOT
# implemented yet: this report only lists POs with commission newly
# due this run, so a cancelled PO (nothing due) never appears as a row
# at all. There's nothing to paint beige until/unless the report
# becomes a full status listing rather than a due-items list - that's
# a bigger design question, not a missing color constant.

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


def _load_run_rows(conn, commission_run_id):
    """
    Joins this run's commission_events against contracts/customers/
    agencies and collapses them to one row per PO - a PO with two
    triggers due in the same run ends up with both sets of columns
    filled on a single row, not two separate rows.
    """
    cursor = conn.execute(
        """
        SELECT
            c.po_no, c.po_date, c.signature_date, c.customer_id, c.lot_no,
            c.niche_price, c.promotion, c.discount, c.net_price,
            c.full_settlement_paid_date, c.first_installment_paid_date,
            c.sixth_installment_paid_date, c.agent_name, c.agency_code, c.remarks,
            cu.name AS customer_name,
            a.splits_by_agent, a.agency_group,
            e.trigger_type, e.amount
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
                "cooling_off_period": None,
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
            }
        row = by_po[po_no]
        if r["trigger_type"] == "full_payment":
            row["full_settlement_paid_date"] = r["full_settlement_paid_date"]
            row["full_payment_commission"] = r["amount"]
            row["cooling_off_period"] = "EXPIRED"
        elif r["trigger_type"] == "installment_1":
            row["first_installment_paid_date"] = r["first_installment_paid_date"]
            row["installment_1_commission"] = r["amount"]
        elif r["trigger_type"] == "installment_6":
            row["sixth_installment_paid_date"] = r["sixth_installment_paid_date"]
            row["installment_6_commission"] = r["amount"]

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


def _write_table(sheet, rows, start_row, title):
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

    totals = {key: 0.0 for key in _TOTAL_KEYS}
    for i, row in enumerate(rows, start=1):
        is_full_payment_row = row.get("full_payment_commission") is not None
        for col, (label, key) in enumerate(_COLUMNS, start=1):
            value = i if key == "row_no" else row.get(key)
            cell = sheet.cell(row=row_num, column=col, value=value)
            cell.font = _BODY_FONT
            if label in _MONEY_COLUMNS:
                cell.number_format = _MONEY_FORMAT
            if is_full_payment_row:
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

    rows = _load_run_rows(conn, commission_run_id)
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

    used_titles = {"All"}
    for group_name, group_rows in groups.items():
        sheet_title = _unique_sheet_title(group_name, used_titles)
        used_titles.add(sheet_title)
        sheet = workbook.create_sheet(sheet_title)
        _write_table(sheet, group_rows, start_row=1, title=_title_line(group_name, group_rows, run_date))
        _autosize_columns(sheet, column_count)

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
                _write_table(agent_sheet, agent_rows, start_row=1, title=_title_line(agent_name, agent_rows, run_date))
                _autosize_columns(agent_sheet, column_count)

    workbook.save(output_path)
    return output_path
