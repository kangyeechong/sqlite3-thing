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
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .importer import _is_positive_whole_number, _to_iso_date, _to_number

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
    # Only ever nonzero when period_start/period_end were given (see
    # import_aor_report) - a receipt whose own Acknowledgment Receipt
    # Date falls outside the chosen period, or has no date at all, so
    # it's left completely unprocessed (not recorded as seen) rather
    # than applied now - it stays available for a later upload whose
    # period actually covers it.
    receipts_outside_period: int = 0
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


def import_aor_report(conn, file_path, imported_by_user=None, aor_upload_id=None,
                       period_start=None, period_end=None):
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

    aor_upload_id: the aor_uploads row this call is processing on
    behalf of (see app.pipeline.process_aor_upload), stamped onto
    every aor_receipts row this call writes. Optional - None for a
    caller that has no aor_uploads row yet (e.g. a script or test
    calling this directly) - it's only used later to scope
    annotate_aor_file's "Filtered" sheet to one specific upload's own
    newly-introduced receipts, never for the transfer logic above,
    which is already correctly scoped via aor_receipts' own
    UNIQUE(acknowledgment_receipt_no).

    period_start/period_end: ISO date strings (both required together,
    or both left None). The real Kenjin AOR export is never cut on
    clean calendar-month boundaries (one real export covered 1 June -
    17 Aug, the next 18 Aug - 23 Sep) - staff process month by month
    regardless, so a receipt whose own Acknowledgment Receipt Date
    falls outside the given period is left completely untouched: not
    recorded into aor_receipts, not applied to a paid-date, nothing.
    It stays exactly as available for a future upload whose period
    actually covers it as if this upload had never mentioned it -
    critically, this is NOT the same as "already imported, skip
    forever" (that's what aor_receipts itself already guards against);
    it's "not this period's business yet." A receipt with no
    Acknowledgment Receipt Date at all is treated the same way when a
    period is given (there's no date to check it against), rather than
    flagged as it would be with no period given - see
    AorImportResult.receipts_outside_period.

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
    # Whether a receipt's own classification counts as a "valid
    # payment" for build_period_audit_workbook depends on its PO
    # actually existing - snapshotted once up front rather than
    # queried per row.
    existing_po_nos = {row["po_no"] for row in conn.execute("SELECT po_no FROM contracts")}
    seen_this_upload = set()
    # ack_no -> {po_no, receipt_date, reference_text, payment_received,
    # trigger_type} - persisted to aor_receipts below regardless of
    # what a row classified as, so build_period_audit_workbook can
    # later show every receipt in a chosen period (not just the ones
    # that mattered) alongside just the valid payments among them.
    receipt_details = {}

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

            receipt_date = _to_iso_date(raw_row.get("Acknowledgment Receipt Date"))

            if period_start is not None:
                if receipt_date is None or not (period_start <= receipt_date <= period_end):
                    # Checked (and skipped) BEFORE seen_this_upload.add -
                    # this row must not count as "applied" so a later
                    # upload whose period actually covers it can still
                    # pick it up. See the period_start/period_end
                    # docstring above for why this is deliberately
                    # different from the already_imported skip above.
                    result.receipts_outside_period += 1
                    continue

            seen_this_upload.add(ack_no)
            receipt_details[ack_no] = {
                "po_no": None,
                "receipt_date": receipt_date,
                "reference_text": raw_row.get("Reference No"),
                "payment_received": _to_number(raw_row.get("Payment Received (RM)")),
                "trigger_type": None,
            }

            if not _is_positive_whole_number(po_no):
                result.review_flags.append(AorReviewFlag(
                    ack_no, po_no, "Row has no valid PO No - skipped.",
                ))
                continue
            po_no = int(po_no)
            receipt_details[ack_no]["po_no"] = po_no

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

            targets = _targets_for_classification(kind, numbers)
            if targets and po_no in existing_po_nos:
                receipt_details[ack_no]["trigger_type"] = ",".join(sorted(targets))
            for target in targets:
                key = (po_no, target)
                if key not in paid_date_candidates or receipt_date > paid_date_candidates[key]:
                    paid_date_candidates[key] = receipt_date

    # Recorded as imported regardless of what it classified as
    # (including "skip" and "unrecognized") - once a human has had the
    # chance to see an unrecognized one flagged, re-flagging the exact
    # same receipt on every future overlapping upload adds nothing.
    for ack_no in seen_this_upload:
        details = receipt_details[ack_no]
        conn.execute(
            "INSERT INTO aor_receipts "
            "(acknowledgment_receipt_no, po_no, imported_at, imported_by_user, source_filename, "
            "aor_upload_id, receipt_date, reference_text, payment_received, trigger_type) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                ack_no, details["po_no"], now_iso, imported_by_user, source_filename,
                aor_upload_id, details["receipt_date"], details["reference_text"],
                details["payment_received"], details["trigger_type"],
            ),
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


def annotate_aor_file(file_path, output, conn=None, aor_upload_id=None):
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
    colors and grouping the business already uses. Sorted by PO No
    ascending (20260299, 20260300, ...) rather than the file's own row
    order, so every row for one PO sits together and the sheet reads
    top to bottom in order.

    Reads the uploaded file twice on purpose: once with data_only=True
    to classify rows and pull out values (the same read every other
    function in this module uses), and once completely untouched to
    build the output from - so a cell that happens to hold a formula
    in the original file is never silently flattened to its cached
    value just because this function also had to read it.

    conn/aor_upload_id: optional - when both are given, the Filtered
    sheet is scoped to only the receipts THIS SPECIFIC upload actually
    introduced (via aor_receipts.aor_upload_id, stamped at import time
    - see import_aor_report). Staff process month by month, but the
    real Kenjin export is cumulative (an "August" export re-lists every
    receipt back to whenever records began), so without this, an old
    month's receipts would show up as if newly relevant every time a
    later cumulative export gets annotated. Left as None (the default)
    this stays exactly what it always was: no ledger lookups, no side
    effects, purely a reflection of what's IN the file - which is
    still exactly right for the very first time a given upload's own
    bytes get annotated, since every receipt it just imported is, by
    definition, stamped with its own aor_upload_id.
    """
    values_workbook = openpyxl.load_workbook(file_path, data_only=True)
    aor_sheets = [sheet for sheet in values_workbook.worksheets if _is_aor_shaped(sheet)]

    this_upload_receipts = None
    if conn is not None and aor_upload_id is not None:
        this_upload_receipts = {
            row["acknowledgment_receipt_no"] for row in conn.execute(
                "SELECT acknowledgment_receipt_no FROM aor_receipts WHERE aor_upload_id = ?",
                (aor_upload_id,),
            )
        }
        if not this_upload_receipts:
            # Zero linked receipts is genuinely ambiguous on its own -
            # it means either "this upload predates aor_upload_id ever
            # existing" (no receipt of its could ever have been
            # stamped) or "this upload went through period filtering
            # and genuinely matched nothing in its own chosen period"
            # (every row fell outside it, so nothing got recorded) -
            # and those need OPPOSITE handling. Falling back to
            # unscoped for the second case doesn't produce an empty
            # sheet the way the first case's fallback intends - it
            # dumps every trigger-shaped row anywhere in this file,
            # including whatever other months a real cumulative Kenjin
            # export repeats, right back into "Filtered" (found live:
            # a period that genuinely matched nothing still showed a
            # year's worth of unrelated dates). aor_uploads.period_start
            # disambiguates them directly: recorded means this upload
            # DID go through period-aware code, so trust the empty
            # result; NULL means it predates that column ever existing,
            # so fall back exactly as before.
            went_through_period_filtering = conn.execute(
                "SELECT period_start FROM aor_uploads WHERE id = ?", (aor_upload_id,),
            ).fetchone()
            if went_through_period_filtering is None or went_through_period_filtering["period_start"] is None:
                this_upload_receipts = None

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
    filtered_rows = []  # list of (po_no, row_values, fill)
    for sheet in aor_sheets:
        header_row_num = _find_header_row(sheet)
        if header_row_num is None:
            continue
        sheet_headers = [cell.value for cell in sheet[header_row_num]]
        if headers is None:
            headers = sheet_headers
        for raw_row in _read_aor_rows(sheet):
            if this_upload_receipts is not None:
                if raw_row.get("Acknowledgment Receipt No") not in this_upload_receipts:
                    continue
            po_no = raw_row.get("PO No")
            kind, numbers = _classify_reference(raw_row.get("Reference No"))
            targets = _targets_for_classification(kind, numbers)
            fill = None
            if "installment_1" in targets or "installment_6" in targets:
                fill = _YELLOW_FILL
            elif _is_positive_whole_number(po_no) and int(po_no) in full_payment_pos:
                fill = _GREEN_FILL
            if fill is not None:
                filtered_rows.append((po_no, [raw_row.get(h) for h in sheet_headers], fill))

    # Sorted by PO No ascending (20260299, 20260300, ...) rather than
    # the file's own row order, so a PO's group of rows is easy to
    # find and everything reads in one consistent order.
    filtered_rows.sort(key=lambda entry: entry[0] if _is_positive_whole_number(entry[0]) else float("inf"))

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
        for row_offset, (_, values, fill) in enumerate(filtered_rows, start=2):
            for col, value in enumerate(values, start=1):
                cell = filtered_sheet.cell(row=row_offset, column=col, value=value)
                cell.fill = fill

    output_workbook.save(output)


def build_period_audit_workbook(
    conn, period_start, period_end, output, po_period_start=None, po_period_end=None,
):
    """
    Writes an audit workbook for a chosen receipt period (ISO date
    strings, inclusive both ends) - built purely from aor_receipts, so
    it naturally spans however many AOR uploads/files actually cover
    that period, not just one. Answers "let me check every payment for
    this period" directly, without having to dig through several
    separate per-upload annotated copies:

      - "All Receipts": every receipt genuinely dated in this period
        (receipt_date, not upload date), whatever it turned out to be -
        deposits, stamp duty, non-1/6 installments, unrecognized text,
        POs not yet in the ledger, all included. A complete audit trail
        to cross-check against Kenjin's own totals for the period.
      - "Valid Payments": the subset that actually counts toward
        commission - matched a real PO in the ledger AND classified as
        full payment, installment 1, or installment 6 (see
        app.aor._classify_reference). This is "the ones we actually
        want to put into Overall Commission."

    po_period_start/po_period_end (both required together, or both left
    None): adds two more sheets, narrowing "All Receipts" down further
    by the PO's own PURCHASE date - a genuinely different axis from the
    receipt period above (confirmed live: payments received in
    September routinely settle POs purchased back in March, and the
    two date ranges are never the same). Answers "of everything that
    came in this receipt period, which of it is actually March's own
    business":
      - "Payments by PO Date": All Receipts further filtered to just
        the POs purchased in this range.
      - "Valid Payments by PO Date": that same subset, narrowed to just
        the valid/triggering ones - "March's payments that are actually
        going into March's Overall Commission report."

    Deliberately reads persisted aor_receipts state, not the original
    uploaded files - a receipt's classification was already decided at
    import time (see import_aor_report), so this never re-parses or
    re-judges anything, just reports back what's already on record.

    Customer ID/Name/Lot No/Agency Code/Agent Name are pulled from the
    ledger (via po_no), not from the AOR row itself - the real AOR
    export's own Customer Name column is routinely blank (confirmed
    against a real sample), while the ledger's copy is reliable since
    it came from the Master report. A receipt with no po_no at all (no
    valid PO No on that row - see import_aor_report), or a po_no not
    yet in the ledger, simply shows blank for all of these rather than
    failing - and is naturally excluded from the two PO-date-scoped
    sheets, since there's no PO Date to filter by either.

    Colored the same way annotate_aor_file's "Filtered" sheet already
    is (yellow for installment 1/6, green for a full payment) - same
    precedence too (yellow wins if a receipt's own trigger_type somehow
    carries both) - so this reads consistently with what staff already
    know from that sheet. A non-triggering or unmatched row (no
    trigger_type - only ever appears on "All Receipts") is left
    uncolored.
    """
    header_font = Font(bold=True)
    columns = (
        ("Acknowledgment Receipt No", "acknowledgment_receipt_no"),
        ("PO No", "po_no"),
        ("PO Date", "po_date"),
        ("Customer ID", "customer_id"),
        ("Customer Name", "customer_name"),
        ("Lot No", "lot_no"),
        ("Agency Code", "agency_code"),
        ("Agent Name", "agent_name"),
        ("Receipt Date", "receipt_date"),
        ("Reference No", "reference_text"),
        ("Payment Received (RM)", "payment_received"),
        ("Trigger Type", "trigger_type"),
        ("Source File", "source_filename"),
    )

    all_receipts = conn.execute(
        """
        SELECT r.acknowledgment_receipt_no, r.po_no, r.receipt_date, r.reference_text,
               r.payment_received, r.trigger_type, r.source_filename,
               c.customer_id, cu.name AS customer_name, c.lot_no, c.agency_code, c.agent_name,
               c.po_date
        FROM aor_receipts r
        LEFT JOIN contracts c ON c.po_no = r.po_no
        LEFT JOIN customers cu ON cu.customer_id = c.customer_id
        WHERE r.receipt_date BETWEEN ? AND ?
        ORDER BY r.receipt_date, r.acknowledgment_receipt_no
        """,
        (period_start, period_end),
    ).fetchall()
    valid_payments = [r for r in all_receipts if r["trigger_type"]]

    workbook = openpyxl.Workbook()

    def _fill_for(trigger_type):
        if not trigger_type:
            return None
        targets = trigger_type.split(",")
        if "installment_1" in targets or "installment_6" in targets:
            return _YELLOW_FILL
        if "full_payment" in targets:
            return _GREEN_FILL
        return None

    def _write_sheet(sheet, title_rows):
        for col, (label, _) in enumerate(columns, start=1):
            cell = sheet.cell(row=1, column=col, value=label)
            cell.font = header_font
        for row_offset, row in enumerate(title_rows, start=2):
            fill = _fill_for(row["trigger_type"])
            for col, (_, key) in enumerate(columns, start=1):
                cell = sheet.cell(row=row_offset, column=col, value=row[key])
                if fill is not None:
                    cell.fill = fill
        for col in range(1, len(columns) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 20

    all_sheet = workbook.active
    all_sheet.title = "All Receipts"
    _write_sheet(all_sheet, all_receipts)

    valid_sheet = workbook.create_sheet("Valid Payments")
    _write_sheet(valid_sheet, valid_payments)

    if po_period_start is not None:
        po_scoped = [
            r for r in all_receipts
            if r["po_date"] is not None and po_period_start <= r["po_date"] <= po_period_end
        ]
        po_scoped_sheet = workbook.create_sheet("Payments by PO Date")
        _write_sheet(po_scoped_sheet, po_scoped)

        po_scoped_valid = [r for r in po_scoped if r["trigger_type"]]
        po_scoped_valid_sheet = workbook.create_sheet("Valid Payments by PO Date")
        _write_sheet(po_scoped_valid_sheet, po_scoped_valid)

    workbook.save(output)
