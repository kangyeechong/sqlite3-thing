"""
Reads a Kenjin AOR (Acknowledgment of Receipt) export and uses it to
fill in contracts' paid-date columns - the same fields Accounts used to
have to type into the Master report by hand after cross-checking this
exact file. Once those dates are on file, the existing detection
pipeline (app.commission.process_commission_run) picks up from there
exactly as if the Master report itself had carried them - nothing about
agency/agent splitting, the Date Record summary, or the review-and-
confirm workflow needed to change for this to work.

The real export's `Reference No` column is free text a staff member
typed by hand, not a structured field - see _classify_reference for
the exact rules (confirmed against a real sample together with the
business, not guessed):
  - An "(INST X/Y)" tag anywhere in the text is authoritative,
    regardless of what word precedes it - "ADVANCE BALANCE PAYMENT
    (INST 01/24)" is installment 1, not a full-payment completion.
    Only installment 1 and 6 matter for commission (the fixed release
    points - see rules.py); any other installment number is recorded
    (so the receipt is never reprocessed) but triggers nothing.
  - No INST tag at all: "FULL PAYMENT" or "BALANCE PAYMENT" means this
    receipt completes the full price. Plain "DEPOSIT" or "PARTIAL
    PAYMENT" (no tag, not the completing one) and "STAMP DUTY" are
    recognized but non-triggering - a step along the way, not a
    commission event.
  - Anything else doesn't match a known pattern and gets flagged for a
    human to look at, rather than silently guessed at or dropped.

The real exports routinely overlap each other (the same underlying
report, exported again later with an extended date range) and a single
export can carry several sheets that are different overlapping
snapshots of the same data - `Acknowledgment Receipt No` is what makes
re-importing safe: every receipt is only ever applied once, ever (see
aor_receipts in schema.sql), the same UNIQUE-constraint-based
idempotency historical_summary_rows already uses for Date Record rows.
"""

import datetime
import re
from dataclasses import dataclass, field

import openpyxl
from openpyxl.styles import PatternFill

from .importer import _is_positive_whole_number, _to_iso_date

# Matches the same green already used elsewhere in this app for "fully
# paid off" (see app/report.py's _GREEN_FILL) - confirmed against a
# real annotated AOR sample, the business's own green highlight (theme
# accent6, tint 0.8) resolves to a very close shade (~DCEDD5) to this
# one, so this keeps one consistent "paid off" color across the whole
# tool rather than introducing a second, slightly different green.
_GREEN_FILL = PatternFill(start_color="C6DEB5", end_color="C6DEB5", fill_type="solid")
_YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

_AOR_REQUIRED_HEADERS = (
    "No", "Acknowledgment Receipt No", "PO No", "Reference No", "Acknowledgment Receipt Date",
)

# Captures the run of digits/slashes/commas/ampersands/whitespace right
# after "INST" - e.g. "15/24" or "22/24, 23/24 & 24/24" - stopping
# naturally at the first character that isn't part of that run (a ")"
# in every real sample seen). Deliberately scoped to just after "INST"
# rather than searching the whole Reference No for any "N/M" pattern,
# since a plain date like "07/08/2026" elsewhere in the same text would
# otherwise falsely match as its own bogus installment number.
_INST_CLUSTER_PATTERN = re.compile(r"INST\s*([\d/,&\s]+)", re.IGNORECASE)
_INST_PAIR_PATTERN = re.compile(r"(\d+)\s*/\s*\d+")

_TRIGGER_TO_DATE_COLUMN = {
    "full_payment": "full_settlement_paid_date",
    "installment_1": "first_installment_paid_date",
    "installment_6": "sixth_installment_paid_date",
}


@dataclass
class AorReviewFlag:
    acknowledgment_receipt_no: object
    po_no: object
    message: str


@dataclass
class AorImportResult:
    receipts_seen: int = 0
    # Newly recorded this upload - 0 for a receipt already imported by
    # an earlier (possibly overlapping) AOR upload, same spirit as
    # ImportResult.historical_rows_imported in app/importer.py.
    receipts_imported: int = 0
    # How many contract paid-date columns actually got filled in - can
    # be smaller than receipts_imported, since several receipts often
    # contribute to one paid-date (see _classify_reference), and a
    # column already on file is never overwritten.
    paid_dates_written: int = 0
    review_flags: list = field(default_factory=list)


def _is_aor_shaped(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=30):
        values = [cell.value for cell in row]
        if all(h in values for h in _AOR_REQUIRED_HEADERS):
            return True
    return False


def _find_header_row(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=30):
        values = [cell.value for cell in row]
        if all(h in values for h in _AOR_REQUIRED_HEADERS):
            return row[0].row
    return None


def _read_aor_rows(sheet):
    """Yields each data row as a dict keyed by header, stopping at the
    first row whose "No" isn't a positive whole number - the same
    "Total"/"Contra"/repeated-header-row footer shape confirmed on the
    real Master report also shows up verbatim at the bottom of a real
    AOR export."""
    header_row_num = _find_header_row(sheet)
    if header_row_num is None:
        return
    headers = [cell.value for cell in sheet[header_row_num]]
    for row in sheet.iter_rows(min_row=header_row_num + 1):
        values = [cell.value for cell in row]
        row_no = values[headers.index("No")]
        if not _is_positive_whole_number(row_no):
            break
        yield dict(zip(headers, values))


def _classify_reference(reference_text):
    """
    Returns one of:
      ("installments", [1, 6, ...])  - an (INST X/Y) tag was found;
                                        every installment number present
      ("full_payment", None)         - completes the full price
      ("skip", None)                 - recognized, non-triggering
      ("unrecognized", None)         - doesn't match any known pattern
    See the module docstring for the confirmed rules this encodes.
    """
    if not reference_text:
        return ("unrecognized", None)
    text = str(reference_text)
    upper = text.upper()

    inst_match = _INST_CLUSTER_PATTERN.search(text)
    if inst_match:
        numbers = [int(n) for n in _INST_PAIR_PATTERN.findall(inst_match.group(1))]
        if numbers:
            return ("installments", numbers)

    if "FULL PAYMENT" in upper or "BALANCE PAYMENT" in upper:
        return ("full_payment", None)
    # "PATRIAL PAYMENT" is a real, recurring typo in the actual export
    # (confirmed against a real sample) - same meaning as "PARTIAL
    # PAYMENT", not a different, unrecognized case.
    if "STAMP DUTY" in upper or "DEPOSIT" in upper or "PARTIAL PAYMENT" in upper or "PATRIAL PAYMENT" in upper:
        return ("skip", None)
    return ("unrecognized", None)


def _targets_for_classification(kind, numbers):
    """Maps a _classify_reference result to the set of trigger_types
    (from _TRIGGER_TO_DATE_COLUMN) this receipt actually contributes
    to. An installment number other than 1 or 6 contributes to none -
    only those two are fixed commission-release points (see rules.py)
    - but the receipt is still recorded as seen either way."""
    if kind == "full_payment":
        return {"full_payment"}
    if kind == "installments":
        targets = set()
        if 1 in numbers:
            targets.add("installment_1")
        if 6 in numbers:
            targets.add("installment_6")
        return targets
    return set()


def import_aor_report(conn, file_path, imported_by_user=None):
    """
    Reads every AOR-shaped sheet in the workbook (a real export often
    has several overlapping ones - see the module docstring), fills in
    whichever of a contract's full_settlement_paid_date /
    first_installment_paid_date / sixth_installment_paid_date columns
    are still blank, and returns an AorImportResult with counts and
    review flags.

    Deliberately never overwrites a paid-date that's already on file
    (whether it came from a Master report upload or an earlier AOR
    one) - this only ever fills in a blank, exactly like Accounts
    manually cross-checking this file and typing the date in once.

    Deliberately does not commit the transaction, matching
    import_master_report - the caller decides when to commit.
    """
    workbook = openpyxl.load_workbook(file_path, data_only=True)
    now_iso = datetime.datetime.now().isoformat()
    source_filename = file_path.split("/")[-1].split("\\")[-1]

    aor_sheets = [sheet for sheet in workbook.worksheets if _is_aor_shaped(sheet)]
    if not aor_sheets:
        raise ValueError(
            f"No sheet in '{file_path}' has the expected AOR headers "
            f"{_AOR_REQUIRED_HEADERS} - is this really an Acknowledgment "
            f"of Receipt export?"
        )

    result = AorImportResult()
    already_imported = {
        row["acknowledgment_receipt_no"]
        for row in conn.execute("SELECT acknowledgment_receipt_no FROM aor_receipts")
    }
    seen_this_upload = set()

    # (po_no, trigger_type) -> latest Acknowledgment Receipt Date seen
    # for it. Several receipts can contribute to the same trigger (e.g.
    # a split "ADVANCE PARTIAL PAYMENT" + "ADVANCE BALANCE PAYMENT",
    # both tagged the same installment number) - the later date is the
    # one that actually completed it.
    paid_date_candidates = {}

    for sheet in aor_sheets:
        for raw_row in _read_aor_rows(sheet):
            result.receipts_seen += 1
            ack_no = raw_row.get("Acknowledgment Receipt No")
            po_no = raw_row.get("PO No")

            if not ack_no:
                result.review_flags.append(AorReviewFlag(
                    None, po_no, "Row has no Acknowledgment Receipt No - can't be safely re-imported, skipped.",
                ))
                continue
            if ack_no in already_imported or ack_no in seen_this_upload:
                continue  # already applied by this or an earlier (possibly overlapping) upload
            seen_this_upload.add(ack_no)

            if not _is_positive_whole_number(po_no):
                result.review_flags.append(AorReviewFlag(
                    ack_no, po_no, "Row has no valid PO No - skipped.",
                ))
                continue
            po_no = int(po_no)

            receipt_date = _to_iso_date(raw_row.get("Acknowledgment Receipt Date"))
            if receipt_date is None:
                result.review_flags.append(AorReviewFlag(
                    ack_no, po_no, "Row has no Acknowledgment Receipt Date - skipped.",
                ))
                continue

            kind, numbers = _classify_reference(raw_row.get("Reference No"))
            if kind == "unrecognized":
                result.review_flags.append(AorReviewFlag(
                    ack_no, po_no,
                    f"Reference No {raw_row.get('Reference No')!r} doesn't match any known "
                    f"pattern - not applied, needs a human to check.",
                ))

            for target in _targets_for_classification(kind, numbers):
                key = (po_no, target)
                if key not in paid_date_candidates or receipt_date > paid_date_candidates[key]:
                    paid_date_candidates[key] = receipt_date

    # Recorded as imported regardless of what it classified as
    # (including "skip" and "unrecognized") - once a human has had the
    # chance to see an unrecognized one flagged, re-flagging the exact
    # same receipt on every future overlapping upload adds nothing.
    for ack_no in seen_this_upload:
        conn.execute(
            "INSERT INTO aor_receipts (acknowledgment_receipt_no, imported_at, imported_by_user, source_filename) "
            "VALUES (?, ?, ?, ?)",
            (ack_no, now_iso, imported_by_user, source_filename),
        )
    result.receipts_imported = len(seen_this_upload)

    for (po_no, target), receipt_date in paid_date_candidates.items():
        contract = conn.execute("SELECT * FROM contracts WHERE po_no = ?", (po_no,)).fetchone()
        if contract is None:
            result.review_flags.append(AorReviewFlag(
                None, po_no, "This PO doesn't exist in the ledger yet - upload the Master report first.",
            ))
            continue
        date_column = _TRIGGER_TO_DATE_COLUMN[target]
        if contract[date_column] is not None:
            continue  # never overwrite an existing paid-date
        conn.execute(
            f"UPDATE contracts SET {date_column} = ? WHERE po_no = ?",
            (receipt_date, po_no),
        )
        result.paid_dates_written += 1

    return result


def annotate_aor_file(file_path, output):
    """
    Writes a copy of the AOR export to `output` (a path or a file-like
    object) with every original sheet left exactly as uploaded, plus
    one new sheet appended - "Filtered" - containing only the rows a
    staff member currently pulls out by hand while going through this
    file: every receipt belonging to a PO that reached full payment in
    this file (the whole group leading up to it, not just the
    completing row - a "PARTIAL"/"PATRIAL PAYMENT" followed by a
    "BALANCE PAYMENT" for the same PO is one completed sale, and both
    rows are included), colored green, and every receipt whose own
    (INST X/Y) tag is installment 1 or 6 specifically, colored yellow.
    Confirmed against a real annotated sample - these are the exact
    colors and grouping the business already uses.

    Reads the uploaded file twice on purpose: once with data_only=True
    to classify rows and pull out values (the same read every other
    function in this module uses), and once completely untouched to
    build the output from - so a cell that happens to hold a formula
    in the original file is never silently flattened to its cached
    value just because this function also had to read it.

    No ledger lookups, no side effects, and independent of
    import_aor_report (this only reflects what's IN the file, not what
    got applied to the database, so it's safe to generate even for
    receipts already seen on an earlier upload).
    """
    values_workbook = openpyxl.load_workbook(file_path, data_only=True)
    aor_sheets = [sheet for sheet in values_workbook.worksheets if _is_aor_shaped(sheet)]

    # First pass: which POs have a full-payment completion anywhere in
    # this file, so every receipt row for that PO (not just the
    # completing one) is pulled into the filtered sheet as green.
    full_payment_pos = set()
    for sheet in aor_sheets:
        for raw_row in _read_aor_rows(sheet):
            po_no = raw_row.get("PO No")
            if not _is_positive_whole_number(po_no):
                continue
            kind, _ = _classify_reference(raw_row.get("Reference No"))
            if kind == "full_payment":
                full_payment_pos.add(int(po_no))

    # Second pass: pick out just the rows that matter, in the order
    # they appear in the file - same spirit as a human scrolling
    # through it top to bottom and filtering as they go.
    headers = None
    filtered_rows = []  # list of (row_values, fill)
    for sheet in aor_sheets:
        header_row_num = _find_header_row(sheet)
        if header_row_num is None:
            continue
        sheet_headers = [cell.value for cell in sheet[header_row_num]]
        if headers is None:
            headers = sheet_headers
        for raw_row in _read_aor_rows(sheet):
            po_no = raw_row.get("PO No")
            kind, numbers = _classify_reference(raw_row.get("Reference No"))
            targets = _targets_for_classification(kind, numbers)
            fill = None
            if "installment_1" in targets or "installment_6" in targets:
                fill = _YELLOW_FILL
            elif _is_positive_whole_number(po_no) and int(po_no) in full_payment_pos:
                fill = _GREEN_FILL
            if fill is not None:
                filtered_rows.append(([raw_row.get(h) for h in sheet_headers], fill))

    output_workbook = openpyxl.load_workbook(file_path)  # untouched - this is what gets kept as-is
    filtered_sheet_name = "Filtered"
    suffix = 2
    while filtered_sheet_name in output_workbook.sheetnames:
        filtered_sheet_name = f"Filtered ({suffix})"
        suffix += 1
    filtered_sheet = output_workbook.create_sheet(filtered_sheet_name)

    if headers is not None:
        for col, header in enumerate(headers, start=1):
            filtered_sheet.cell(row=1, column=col, value=header)
        for row_offset, (values, fill) in enumerate(filtered_rows, start=2):
            for col, value in enumerate(values, start=1):
                cell = filtered_sheet.cell(row=row_offset, column=col, value=value)
                cell.fill = fill

    output_workbook.save(output)
