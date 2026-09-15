"""
Builds the downloadable Excel report for one commission run - the file
that gets handed to Accounts.

Grouping mirrors how your team already splits the Master report today
(docs/data_model.md section 6a): every agency gets its own sheet, and
within an agency where `splits_by_agent` is on, that sheet is further
broken into one section per agent. This is pure grouping of already-
calculated flat commission amounts - it does not compute AW
Consultancy's agency/agent split math, which is still a later phase.
"""

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

_BODY_FONT = Font(name="Arial", size=11)
_HEADER_FONT = Font(name="Arial", size=11, bold=True)
_TITLE_FONT = Font(name="Arial", size=13, bold=True)
_MONEY_FORMAT = "#,##0.00"

_TRIGGER_LABELS = {
    "full_payment": "Full Payment",
    "installment_1": "Instalment 1",
    "installment_6": "Instalment 6 (Balance)",
}

# (column header, key into the row dict)
_COLUMNS = [
    ("PO No", "po_no"),
    ("Customer Name", "customer_name"),
    ("Lot No", "lot_no"),
    ("Agent", "agent_name"),
    ("Agency Code", "agency_code"),
    ("Trigger", "trigger_label"),
    ("Trigger Date", "trigger_date"),
    ("Net Price (RM)", "net_price"),
    ("Commission (RM)", "amount"),
]
_MONEY_COLUMNS = {"Net Price (RM)", "Commission (RM)"}


def _load_run_rows(conn, commission_run_id):
    """Joins this run's commission_events against contracts/customers/
    agencies so each row has everything the report needs to display."""
    cursor = conn.execute(
        """
        SELECT
            e.po_no, e.trigger_type, e.trigger_date, e.amount,
            c.agent_name, c.agency_code, c.lot_no, c.net_price,
            cu.name AS customer_name,
            a.splits_by_agent
        FROM commission_events e
        JOIN contracts c ON c.po_no = e.po_no
        LEFT JOIN customers cu ON cu.customer_id = c.customer_id
        LEFT JOIN agencies a ON a.agency_code = c.agency_code
        WHERE e.commission_run_id = ?
        ORDER BY c.agency_code, c.agent_name, e.po_no
        """,
        (commission_run_id,),
    )
    rows = []
    for r in cursor.fetchall():
        rows.append({
            "po_no": r["po_no"],
            "customer_name": r["customer_name"],
            "lot_no": r["lot_no"],
            "agent_name": r["agent_name"] or "(unassigned)",
            "agency_code": r["agency_code"] or "(No Agency)",
            "trigger_label": _TRIGGER_LABELS.get(r["trigger_type"], r["trigger_type"]),
            "trigger_date": r["trigger_date"],
            "net_price": r["net_price"],
            "amount": r["amount"],
            # No agency on file -> nothing to group by agent for either;
            # default to a flat listing rather than splitting by agent.
            "splits_by_agent": bool(r["splits_by_agent"]) if r["agency_code"] else False,
        })
    return rows


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


def _write_table(sheet, rows, start_row, title):
    """Writes one titled table (header + data + bold total row) starting
    at start_row. Returns the next free row, so tables can be stacked."""
    row_num = start_row
    sheet.cell(row=row_num, column=1, value=title).font = _TITLE_FONT
    row_num += 2

    header_row = row_num
    for col, (label, _key) in enumerate(_COLUMNS, start=1):
        cell = sheet.cell(row=header_row, column=col, value=label)
        cell.font = _HEADER_FONT
    row_num += 1

    total = 0.0
    for row in rows:
        for col, (label, key) in enumerate(_COLUMNS, start=1):
            cell = sheet.cell(row=row_num, column=col, value=row[key])
            cell.font = _BODY_FONT
            if label in _MONEY_COLUMNS:
                cell.number_format = _MONEY_FORMAT
        total += row["amount"]
        row_num += 1

    label_col = len(_COLUMNS) - 1
    amount_col = len(_COLUMNS)
    sheet.cell(row=row_num, column=label_col, value="Total").font = _HEADER_FONT
    total_cell = sheet.cell(row=row_num, column=amount_col, value=round(total, 2))
    total_cell.font = _HEADER_FONT
    total_cell.number_format = _MONEY_FORMAT
    row_num += 1

    return row_num + 1  # one blank row before whatever comes next


def _autosize_columns(sheet):
    for col in range(1, len(_COLUMNS) + 1):
        sheet.column_dimensions[get_column_letter(col)].width = 20


def generate_commission_run_report(conn, commission_run_id, output_path):
    """
    Writes the downloadable Excel file for one commission run:
      - "All" sheet: every newly-due item in one flat list (the Master
        view)
      - one sheet per agency
      - within an agency with splits_by_agent on, that sheet is further
        broken into one section per agent

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

    rows = _load_run_rows(conn, commission_run_id)

    workbook = Workbook()
    all_sheet = workbook.active
    all_sheet.title = "All"
    _write_table(all_sheet, rows, start_row=1, title="All Commission Due This Run")
    _autosize_columns(all_sheet)

    agencies = {}
    for row in rows:
        agencies.setdefault(row["agency_code"], []).append(row)

    used_titles = {"All"}
    for agency_code, agency_rows in agencies.items():
        title = _unique_sheet_title(agency_code, used_titles)
        used_titles.add(title)
        sheet = workbook.create_sheet(title)

        if not agency_rows[0]["splits_by_agent"]:
            _write_table(sheet, agency_rows, start_row=1, title=f"{agency_code} - Commission Due")
        else:
            agents = {}
            for row in agency_rows:
                agents.setdefault(row["agent_name"], []).append(row)
            row_num = 1
            for agent_name, agent_rows in agents.items():
                row_num = _write_table(sheet, agent_rows, start_row=row_num, title=agent_name)
        _autosize_columns(sheet)

    workbook.save(output_path)
    return output_path
