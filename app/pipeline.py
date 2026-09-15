"""
Ties the importer and commission logic together into the single
operation a user actually triggers: "I uploaded this file, tell me
what's newly due."
"""

import datetime
import os

from .db.connection import get_connection, init_db
from .importer import import_master_report
from .commission import process_commission_run


def process_upload(db_path, file_path, run_date=None, created_by_user=None):
    """
    Runs the full pipeline against one uploaded Master report:
      1. import every PO into the ledger (creates the DB if needed)
      2. determine what's newly due - full payment, installment 1, or
         installment 6
      3. log it, mark it flagged, group it under one commission_run

    Returns a dict with the import result and the commission run
    result, so a caller (a future web route, or a test) can show both
    to the user - the review flags from step 1 and the commission run
    from step 2 are both worth a human's eyes before anything is
    downloaded.
    """
    if run_date is None:
        run_date = datetime.date.today()

    if not os.path.exists(db_path):
        init_db(db_path)

    conn = get_connection(db_path)
    try:
        import_result = import_master_report(conn, file_path)

        run_id, raised_events = process_commission_run(
            conn,
            as_of=run_date,
            run_date=run_date,
            source_filename=os.path.basename(file_path),
            created_by_user=created_by_user,
        )

        conn.commit()
    finally:
        conn.close()

    return {
        "import_result": import_result,
        "commission_run_id": run_id,
        "raised_events": raised_events,
    }
