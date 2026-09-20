"""
Ties the importer and commission logic together into the single
operation a user actually triggers: "I uploaded this file, tell me
what's newly due."
"""

import contextlib
import datetime
import os

from .aor import import_aor_report
from .db.connection import get_connection, init_db
from .importer import import_master_report
from .commission import confirm_commission_events, load_events_for_run, process_commission_run
from .report import generate_commission_run_report


@contextlib.contextmanager
def _connect(db_path):
    """Every function below needs the same open/close-on-the-way-out
    connection lifecycle; sharing it here means a future change to it
    (e.g. logging, a busy_timeout) can't be missed in one function but
    not another."""
    conn = get_connection(db_path)
    try:
        yield conn
    finally:
        conn.close()


def process_upload(db_path, file_path, run_date=None, created_by_user=None):
    """
    Runs the full pipeline against one uploaded Master report:
      1. import every PO into the ledger (creates the DB if needed)
      2. determine what's newly due - full payment, installment 1, or
         installment 6
      3. log it as *pending* and mark it flagged, grouped under one
         commission_run - detecting something is never the same as it
         being confirmed due; see confirm_events below

    Returns a dict with the import result and the commission run
    result, so a caller (a future web route, or a test) can show both
    to the user - the review flags from step 1 and the newly-detected
    (still pending) events from step 2 are both worth a human's eyes
    before anything gets confirmed and downloaded.
    """
    if run_date is None:
        run_date = datetime.date.today()

    if not os.path.exists(db_path):
        init_db(db_path)

    with _connect(db_path) as conn:
        import_result = import_master_report(conn, file_path, imported_by_user=created_by_user)

        run_id, raised_events = process_commission_run(
            conn,
            as_of=run_date,
            run_date=run_date,
            source_filename=os.path.basename(file_path),
            created_by_user=created_by_user,
        )

        conn.commit()

    return {
        "import_result": import_result,
        "commission_run_id": run_id,
        "raised_events": raised_events,
    }


def process_aor_upload(db_path, file_path, run_date=None, created_by_user=None):
    """
    Runs the AOR pipeline against one uploaded Acknowledgment of
    Receipt export - the exact same two-step shape as process_upload
    above, just fed by a different source file:
      1. fill in whatever paid-date columns this export's receipts
         confirm that aren't already on file (app.aor.import_aor_report)
      2. run the SAME detection process_upload uses, so a newly-filled
         paid-date is picked up exactly as if it had come from the
         Master report itself - agency/agent splitting, the Date
         Record summary, and the review-and-confirm workflow all keep
         working unchanged, since none of them care where a paid-date
         came from.
    """
    if run_date is None:
        run_date = datetime.date.today()

    if not os.path.exists(db_path):
        init_db(db_path)

    with _connect(db_path) as conn:
        import_result = import_aor_report(conn, file_path, imported_by_user=created_by_user)

        run_id, raised_events = process_commission_run(
            conn,
            as_of=run_date,
            run_date=run_date,
            source_filename=os.path.basename(file_path),
            created_by_user=created_by_user,
        )

        # Kept so the "download annotated copy" route can regenerate
        # the colored version fresh from the original upload at any
        # later time (see app/aor.py's annotate_aor_file) - not the
        # annotated bytes themselves, since those are cheap to
        # regenerate and this way a future change to the coloring
        # rules applies retroactively to every past upload's download
        # too. Keyed by its own id rather than run_id, since run_id is
        # None whenever nothing was newly detected this upload (see
        # process_commission_run) but the file is always worth
        # annotating regardless.
        with open(file_path, "rb") as f:
            file_bytes = f.read()
        upload_cursor = conn.execute(
            "INSERT INTO aor_uploads (commission_run_id, filename, file_bytes, uploaded_at) "
            "VALUES (?, ?, ?, ?)",
            (run_id, os.path.basename(file_path), file_bytes, datetime.datetime.now().isoformat()),
        )
        aor_upload_id = upload_cursor.lastrowid

        conn.commit()

    return {
        "import_result": import_result,
        "commission_run_id": run_id,
        "raised_events": raised_events,
        "aor_upload_id": aor_upload_id,
    }


def generate_report(db_path, commission_run_id, output_path):
    """
    Writes the downloadable Excel report for a commission run that was
    already created by process_upload. Deliberately a separate step,
    not bundled into process_upload automatically - the intended flow
    is: upload, review what got raised on screen, confirm it, then
    download, so a problem can be caught before a file ever reaches
    Accounts.
    """
    with _connect(db_path) as conn:
        return generate_commission_run_report(conn, commission_run_id, output_path)


def load_review(db_path, commission_run_id):
    """Every event (pending or already confirmed) on one commission
    run, for the review-and-confirm page."""
    with _connect(db_path) as conn:
        return load_events_for_run(conn, commission_run_id)


def list_confirmed_runs(db_path):
    """
    Every commission run that has at least one confirmed event, newest
    first - what the "past reports" page lists, so a report is always
    reachable even long after the upload that produced it, not just
    from the one link shown right after that specific upload. A run
    sitting fully pending (nothing confirmed on it yet) is left out,
    same as everywhere else - there'd be nothing to download from it.
    """
    with _connect(db_path) as conn:
        return conn.execute(
            """
            SELECT DISTINCT r.id, r.run_date, r.source_filename
            FROM commission_runs r
            JOIN commission_events e ON e.commission_run_id = r.id AND e.status = 'confirmed'
            ORDER BY r.run_date DESC, r.id DESC
            """
        ).fetchall()


def list_runs_with_pending_events(db_path):
    """
    Every commission run that still has at least one pending (not yet
    confirmed) event, newest first - the runs list_confirmed_runs above
    deliberately leaves out. Without this, a run that raised something
    newly due is only ever reachable via the one link shown right on
    its own results page - navigate away before confirming it (or come
    back a different day) and there was previously no way back to it
    at all, since it doesn't qualify for "Past Reports" either (nothing
    on it is confirmed yet). This is what makes /review/<run_id>
    reachable again for exactly that case.
    """
    with _connect(db_path) as conn:
        return conn.execute(
            """
            SELECT DISTINCT r.id, r.run_date, r.source_filename
            FROM commission_runs r
            JOIN commission_events e ON e.commission_run_id = r.id AND e.status = 'pending'
            ORDER BY r.run_date DESC, r.id DESC
            """
        ).fetchall()


def list_aor_uploads(db_path):
    """
    Every AOR export ever uploaded, newest first - so its annotated
    copy (see app/aor.py's annotate_aor_file) stays downloadable from
    the "Past Reports" page long after the upload that produced it,
    not just from the one link shown right after that specific
    upload. Deliberately not filtered by whether anything was newly
    detected that run (unlike list_confirmed_runs above) - an AOR
    upload's file is always worth being able to get back to, whether
    or not it happened to raise anything new that time.
    """
    with _connect(db_path) as conn:
        return conn.execute(
            "SELECT id, filename, uploaded_at FROM aor_uploads ORDER BY uploaded_at DESC, id DESC"
        ).fetchall()


def confirm_events(db_path, commission_run_id, event_ids, confirmed_by_user):
    """
    Confirms the chosen events (must belong to commission_run_id - an
    id for a different run is silently ignored, not just any id the
    caller happens to pass) and returns how many were actually
    confirmed by this call.
    """
    with _connect(db_path) as conn:
        valid_ids = {
            row["id"] for row in load_events_for_run(conn, commission_run_id)
            if row["id"] in event_ids
        }
        confirmed_count = confirm_commission_events(conn, list(valid_ids), confirmed_by_user)
        conn.commit()
        return confirmed_count
