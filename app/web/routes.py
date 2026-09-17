"""
Web routes: login, upload, review results, download.

Deliberately thin - every route just calls into app.pipeline, the same
functions already proven by the test suite and the real-data checks.
The web layer's only job is turning HTTP requests into calls to that
existing logic and rendering the result, not containing new business
logic of its own.
"""

import datetime
import io
import os
import tempfile

from flask import (
    Blueprint, abort, current_app, flash, redirect, render_template,
    request, send_file, session, url_for,
)
from werkzeug.utils import secure_filename

from ..db.connection import get_connection
from ..pipeline import confirm_events, load_review, process_upload
from ..report import TRIGGER_LABELS, generate_commission_run_report
from .auth import find_user_by_email, hash_password, login_required, verify_password
from .csrf import validate_csrf_token

bp = Blueprint("web", __name__)

# A precomputed hash of a placeholder that is never a real password,
# used only so that logging in with an email that doesn't exist takes
# roughly as long as a real (wrong) password check - otherwise a
# missing account returns near-instantly while a wrong password on a
# real account takes measurably longer, letting someone probe which
# emails have accounts on a tool that gates access to real payout data.
_DUMMY_HASH_FOR_TIMING_SAFETY = hash_password("not-a-real-password-used-only-for-timing-safety")


def _abort_if_run_missing(conn, run_id):
    """Shared by review() and download() - both need the same 404
    before doing anything else with a run id that might not exist."""
    run_exists = conn.execute(
        "SELECT 1 FROM commission_runs WHERE id = ?", (run_id,)
    ).fetchone()
    if run_exists is None:
        abort(404, description=f"No commission run with id {run_id}.")


@bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html")

    validate_csrf_token(request.form.get("csrf_token"))

    email = request.form.get("email", "").strip().lower()
    password = request.form.get("password", "")

    conn = get_connection(current_app.config["DB_PATH"])
    try:
        user = find_user_by_email(conn, email)
    finally:
        conn.close()

    stored_hash = user["password_hash"] if user is not None else _DUMMY_HASH_FOR_TIMING_SAFETY
    password_ok = verify_password(password, stored_hash)  # always runs, even for a missing user - see the comment above

    if user is None or not password_ok:
        flash("Incorrect email or password.")
        return render_template("login.html"), 401

    session["user_email"] = user["email"]
    session["display_name"] = user["display_name"] or user["email"]
    return redirect(url_for("web.upload"))


@bp.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("web.login"))


@bp.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "GET":
        return render_template("upload.html")

    validate_csrf_token(request.form.get("csrf_token"))

    uploaded_file = request.files.get("report_file")
    if uploaded_file is None or uploaded_file.filename == "":
        flash("Choose a Commission Base Report file first.")
        return render_template("upload.html"), 400

    # secure_filename strips path separators and traversal sequences
    # (e.g. "../../etc/cron.d/evil" or an absolute path) - without it,
    # a crafted upload filename could make os.path.join() write outside
    # tmp_dir entirely.
    safe_name = secure_filename(uploaded_file.filename) or "upload.xlsx"

    # Saved to a real temp file (not read straight from the upload
    # stream) because openpyxl needs to seek around the file while
    # reading it - fully processed before this block exits, so there's
    # no risk of the file disappearing mid-read.
    with tempfile.TemporaryDirectory() as tmp_dir:
        saved_path = os.path.join(tmp_dir, safe_name)
        uploaded_file.save(saved_path)

        try:
            result = process_upload(
                current_app.config["DB_PATH"],
                saved_path,
                run_date=datetime.date.today(),
                created_by_user=session["user_email"],
            )
        except Exception as exc:
            # Broad on purpose: this is the boundary where an
            # unpredictable, user-supplied file gets parsed (a renamed
            # non-Excel file, a corrupted workbook, an unexpected
            # format openpyxl itself rejects with its own exception
            # types) - every failure here should become a message the
            # person uploading can act on, never a raw server error.
            flash(f"Couldn't process this file: {exc}")
            return render_template("upload.html"), 400

    return render_template(
        "results.html",
        import_result=result["import_result"],
        raised_events=result["raised_events"],
        commission_run_id=result["commission_run_id"],
        trigger_labels=TRIGGER_LABELS,
    )


@bp.route("/review/<int:run_id>")
@login_required
def review(run_id):
    conn = get_connection(current_app.config["DB_PATH"])
    try:
        _abort_if_run_missing(conn, run_id)
    finally:
        conn.close()

    events = load_review(current_app.config["DB_PATH"], run_id)
    pending_events = [e for e in events if e["status"] == "pending"]
    confirmed_events = [e for e in events if e["status"] == "confirmed"]

    return render_template(
        "review.html",
        run_id=run_id,
        pending_events=pending_events,
        confirmed_events=confirmed_events,
        trigger_labels=TRIGGER_LABELS,
    )


@bp.route("/review/<int:run_id>/confirm", methods=["POST"])
@login_required
def confirm_review(run_id):
    validate_csrf_token(request.form.get("csrf_token"))

    # Checkbox values arrive as strings; anything that isn't a valid
    # event id is simply not a valid id to confirm and gets dropped -
    # confirm_events() only ever touches ids that both parse here AND
    # actually belong to this run.
    event_ids = set()
    for raw_id in request.form.getlist("event_id"):
        try:
            event_ids.add(int(raw_id))
        except ValueError:
            continue

    confirmed_count = confirm_events(
        current_app.config["DB_PATH"], run_id, event_ids, session["user_email"]
    )
    if confirmed_count:
        flash(f"Confirmed {confirmed_count} commission(s).")
    else:
        flash("Nothing was confirmed - select at least one row first.")

    return redirect(url_for("web.review", run_id=run_id))


@bp.route("/download/<int:run_id>")
@login_required
def download(run_id):
    # One connection, reused for both the existence check and the
    # report query - generate_commission_run_report() is called
    # directly here (rather than through pipeline.generate_report(),
    # which would open a second connection of its own) specifically to
    # avoid that.
    conn = get_connection(current_app.config["DB_PATH"])
    try:
        _abort_if_run_missing(conn, run_id)

        # Written to an in-memory buffer, not a temp file on disk - a
        # temp file would get cleaned up as soon as this function
        # returns, but send_file actually streams the response body
        # *after* the view function has already returned. Keeping it
        # in memory sidesteps that race entirely.
        buffer = io.BytesIO()
        try:
            generate_commission_run_report(conn, run_id, buffer)
        except ValueError:
            # Nothing confirmed yet on this run - send them to the
            # review page to fix that, rather than a raw stack trace.
            flash("Nothing on this run is confirmed yet - review and confirm it first.")
            return redirect(url_for("web.review", run_id=run_id))
    finally:
        conn.close()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"commission_run_{run_id}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
