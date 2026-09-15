"""
Reads a Kenjin "Commission Base Report" (Master report) Excel file and
loads it into the ledger database.

This is the only input Phase 1 accepts - see docs/data_model.md section
1 for why the separate AOR (Acknowledgment of Receipt) file is
deliberately not parsed here.
"""

import datetime
from dataclasses import dataclass, field

import openpyxl

from . import parsing

# The real Master sheet has a title block in rows 1-5, headers on row 6,
# and data from row 7 onward - confirmed against the real sample file.
# The importer doesn't hardcode "row 6" though: it searches for the row
# containing "PO No", so a sheet with an extra or missing blank row
# doesn't silently misread everything as garbage.
_REQUIRED_HEADERS = ("No", "PO No", "Customer ID")


@dataclass
class ReviewFlag:
    po_no: object
    check: str
    message: str


@dataclass
class ImportResult:
    contracts_seen: int = 0
    contracts_new: int = 0
    contracts_updated: int = 0
    review_flags: list = field(default_factory=list)


def _find_header_row(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=30):
        values = [cell.value for cell in row]
        if all(h in values for h in _REQUIRED_HEADERS):
            return row[0].row
    raise ValueError(
        f"Could not find a header row containing {_REQUIRED_HEADERS} "
        f"in sheet '{sheet.title}' - is this really a Commission Base "
        f"Report?"
    )


def _is_master_shaped(sheet):
    try:
        _find_header_row(sheet)
        return True
    except ValueError:
        return False


def _to_iso_date(value):
    """Normalizes an Excel date cell (datetime, date, or None) to an
    ISO date string, or None."""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date().isoformat()
    if isinstance(value, datetime.date):
        return value.isoformat()
    return None


def _to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _is_positive_whole_number(value):
    """
    True for 1, 2, 3... whether Excel/openpyxl hands it back as an int
    or as a float (a "No" or "PO No" column with General number format
    can come back as 1.0 rather than 1) - checked explicitly because
    isinstance(value, int) alone would silently misread a float-typed
    column as the end of the data and truncate the import.
    """
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return value > 0
    if isinstance(value, float):
        return value > 0 and value.is_integer()
    return False


def _read_rows(sheet):
    """
    Yields one dict per data row, keyed by the sheet's own header text.
    Stops as soon as a row's "No" column isn't a positive integer -
    that's how the real sheet's trailing summary table (Date Record /
    Full Commission / ... - see docs/data_model.md section 2) gets
    excluded without needing to know its exact row position.
    """
    header_row_num = _find_header_row(sheet)
    headers = [cell.value for cell in sheet[header_row_num]]

    for row in sheet.iter_rows(min_row=header_row_num + 1):
        values = [cell.value for cell in row]
        row_no = values[headers.index("No")]
        if not _is_positive_whole_number(row_no):
            break
        yield dict(zip(headers, values))


def _build_contract_fields(raw_row):
    niche_price = _to_number(raw_row.get("Niche/Tablet Price (RM)")) or 0.0
    promotion = _to_number(raw_row.get("Promotion (RM)")) or 0.0
    discount = _to_number(raw_row.get("Discount (RM)")) or 0.0
    net_price = round(niche_price - promotion - discount, 2)

    remarks = raw_row.get("Remarks")

    return {
        "po_no": int(raw_row["PO No"]),
        "customer_id": raw_row.get("Customer ID"),
        "customer_name": raw_row.get("Customer Name"),
        "agent_name": raw_row.get("FCC/Agent"),
        "agency_code": raw_row.get("Agency Code") or None,
        "lot_no": raw_row.get("Lot No"),
        "po_date": _to_iso_date(raw_row.get("PO Date")),
        "signature_date": _to_iso_date(raw_row.get("Signature Date")),
        "niche_price": niche_price,
        "promotion": promotion,
        "discount": discount,
        "net_price": net_price,
        "case_type": parsing.detect_case_type(remarks),
        "inurnment_date": parsing.extract_inurnment_date(remarks),
        "status": parsing.detect_status(remarks),
        "full_settlement_paid_date": _to_iso_date(raw_row.get("Full Settlement Paid Date")),
        "first_installment_paid_date": _to_iso_date(raw_row.get("First Instalment Paid Date")),
        "sixth_installment_paid_date": _to_iso_date(raw_row.get("Sixth Instalment Paid Date")),
        "remarks": remarks,
    }


def _review_checks(conn, fields_list):
    """
    Scans the whole batch for anomalies before anything is written.
    Returns a list of ReviewFlag - informational only, nothing here
    blocks the import. See docs/data_model.md section 6b for the
    rationale behind each check.
    """
    flags = []
    seen_po_nos = set()
    known_agency_codes = {
        row["agency_code"] for row in conn.execute("SELECT agency_code FROM agencies")
    }

    for fields in fields_list:
        po_no = fields["po_no"]

        if po_no in seen_po_nos:
            flags.append(ReviewFlag(po_no, "duplicate_po", "This PO No appears more than once in this upload."))
        seen_po_nos.add(po_no)

        if fields["first_installment_paid_date"] and fields["sixth_installment_paid_date"]:
            if fields["sixth_installment_paid_date"] < fields["first_installment_paid_date"]:
                flags.append(ReviewFlag(
                    po_no, "dates_out_of_order",
                    "Sixth Instalment Paid Date is earlier than First Instalment Paid Date.",
                ))

        if fields["net_price"] is not None and fields["net_price"] <= 0:
            flags.append(ReviewFlag(po_no, "non_positive_net_price", f"Net Price is {fields['net_price']}."))

        if fields["agency_code"] and fields["agency_code"] not in known_agency_codes:
            flags.append(ReviewFlag(
                po_no, "unknown_agency_code",
                f"Agency Code '{fields['agency_code']}' hasn't been seen before - check for a typo.",
            ))
            known_agency_codes.add(fields["agency_code"])  # don't re-flag within the same batch

        if fields["status"] == "active" and not fields["agency_code"]:
            flags.append(ReviewFlag(po_no, "blank_agency_code", "Agency Code is blank on an active PO."))

        has_any_payment = any([
            fields["full_settlement_paid_date"],
            fields["first_installment_paid_date"],
            fields["sixth_installment_paid_date"],
        ])
        if fields["status"] in ("cancelled", "withdrawn", "on_hold") and has_any_payment:
            flags.append(ReviewFlag(
                po_no, "payment_on_inactive_po",
                f"Status is '{fields['status']}' but a paid date is present.",
            ))

        if fields["case_type"] == "at_need" and fields["inurnment_date"] is None:
            flags.append(ReviewFlag(
                po_no, "at_need_missing_inurnment_date",
                "Remarks mention an At Need case but no inurnment date could be read from it.",
            ))

        existing = conn.execute(
            "SELECT status FROM contracts WHERE po_no = ?", (po_no,)
        ).fetchone()
        if existing is not None and existing["status"] != fields["status"]:
            flags.append(ReviewFlag(
                po_no, "status_changed",
                f"Status changed from '{existing['status']}' to '{fields['status']}' "
                f"based on Remarks text - please confirm this is correct.",
            ))

    return flags


def _upsert_contract(conn, fields, now_iso):
    conn.execute(
        "INSERT INTO customers (customer_id, name) VALUES (?, ?) "
        "ON CONFLICT(customer_id) DO UPDATE SET name = excluded.name",
        (fields["customer_id"], fields["customer_name"]),
    )

    if fields["agency_code"]:
        conn.execute(
            "INSERT INTO agencies (agency_code) VALUES (?) "
            "ON CONFLICT(agency_code) DO NOTHING",
            (fields["agency_code"],),
        )

    existing = conn.execute(
        "SELECT po_no FROM contracts WHERE po_no = ?", (fields["po_no"],)
    ).fetchone()

    conn.execute(
        """
        INSERT INTO contracts (
            po_no, customer_id, agent_name, agency_code, lot_no,
            po_date, signature_date, niche_price, promotion, discount,
            net_price, case_type, inurnment_date, status,
            full_settlement_paid_date, first_installment_paid_date,
            sixth_installment_paid_date, remarks, updated_at
        ) VALUES (
            :po_no, :customer_id, :agent_name, :agency_code, :lot_no,
            :po_date, :signature_date, :niche_price, :promotion, :discount,
            :net_price, :case_type, :inurnment_date, :status,
            :full_settlement_paid_date, :first_installment_paid_date,
            :sixth_installment_paid_date, :remarks, :now
        )
        ON CONFLICT(po_no) DO UPDATE SET
            customer_id = excluded.customer_id,
            agent_name = excluded.agent_name,
            agency_code = excluded.agency_code,
            lot_no = excluded.lot_no,
            po_date = excluded.po_date,
            signature_date = excluded.signature_date,
            niche_price = excluded.niche_price,
            promotion = excluded.promotion,
            discount = excluded.discount,
            net_price = excluded.net_price,
            case_type = excluded.case_type,
            inurnment_date = excluded.inurnment_date,
            status = excluded.status,
            full_settlement_paid_date = excluded.full_settlement_paid_date,
            first_installment_paid_date = excluded.first_installment_paid_date,
            sixth_installment_paid_date = excluded.sixth_installment_paid_date,
            remarks = excluded.remarks,
            updated_at = excluded.updated_at
        """,
        {**fields, "now": now_iso},
    )

    return existing is None


def import_master_report(conn, file_path):
    """
    Reads every sheet in the workbook that looks like a Master report
    (has the expected headers), upserts every PO into `contracts`, and
    returns an ImportResult with counts and review flags.

    Deliberately does not commit the transaction - the caller decides
    when to commit, so a review of the flags can happen before the
    import is finalized if desired.
    """
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    now_iso = datetime.datetime.now().isoformat()

    # Only the Master sheet is read, not the per-agency sheets that
    # follow it in the same workbook - those are filtered views of the
    # identical rows (confirmed against the real sample file), not
    # additional data. Reading every sheet would mean seeing each PO
    # multiple times and having to guess whether a repeat is that
    # expected cross-sheet echo or a genuine duplicate-entry mistake;
    # reading only the first matching sheet avoids the ambiguity
    # entirely, so a real duplicate PO *within* that sheet still gets
    # caught by the review panel below instead of being masked by it.
    master_sheet = next(
        (sheet for sheet in workbook.worksheets if _is_master_shaped(sheet)),
        None,
    )
    if master_sheet is None:
        raise ValueError(
            f"No sheet in '{file_path}' has the expected headers "
            f"{_REQUIRED_HEADERS} - is this really a Commission Base Report?"
        )

    all_fields = []
    skip_flags = []
    for raw_row in _read_rows(master_sheet):
        po_no_value = raw_row.get("PO No")
        if not _is_positive_whole_number(po_no_value):
            # Can't import a row with no usable primary key - skip just
            # this row rather than letting int() raise and crash the
            # whole upload for every other valid row in the file.
            skip_flags.append(ReviewFlag(
                raw_row.get("No"), "missing_po_no",
                f"Row 'No'={raw_row.get('No')} has no valid PO No and was not imported.",
            ))
            continue
        all_fields.append(_build_contract_fields(raw_row))

    result = ImportResult(contracts_seen=len(all_fields))
    result.review_flags = skip_flags + _review_checks(conn, all_fields)

    for fields in all_fields:
        is_new = _upsert_contract(conn, fields, now_iso)
        if is_new:
            result.contracts_new += 1
        else:
            result.contracts_updated += 1

    return result
