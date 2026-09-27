"""
Flask app factory for the Agent Commission Tool web interface.

This wraps the already-proven ledger logic (app.importer, app.commission,
app.report) - it contains no business logic of its own, only HTTP
plumbing. See app/web/routes.py.
"""

import os
import sys
import warnings

from flask import Flask, abort, redirect, render_template, request, session, url_for

from ..db.connection import init_db
from .csrf import get_csrf_token, validate_csrf_token
from .i18n import LANGUAGES, t
from .routes import bp as web_blueprint

_DEV_ONLY_SECRET_KEY = "dev-only-not-for-production"


def create_app(db_path, secret_key=None):
    app = Flask(__name__)
    app.config["DB_PATH"] = db_path

    app.secret_key = secret_key or os.environ.get("SECRET_KEY")
    if not app.secret_key:
        # A silent fallback here would be a real vulnerability: this
        # key signs login session cookies, and the fallback value is
        # public (it's in version control) - anyone who knows it could
        # forge a session and log in as any user with no password.
        # Fine for local testing (printed loudly so it's never missed),
        # never for anything reachable by anyone else.
        warnings.warn(
            "No SECRET_KEY set - using a PUBLIC, well-known development "
            "key. Session cookies can be forged by anyone who has read "
            "this source code. Do not use this outside local testing; "
            "set the SECRET_KEY environment variable before running "
            "this anywhere another person can reach it.",
            stacklevel=2,
        )
        print(
            "\n*** WARNING: running with a public, dev-only SECRET_KEY. "
            "Do not expose this server beyond your own machine. ***\n",
            file=sys.stderr,
        )
        app.secret_key = _DEV_ONLY_SECRET_KEY

    # Commission Base Report exports are at most a few hundred KB in
    # every sample seen so far - 20MB is generous headroom while still
    # bounding upload size, so a very large (accidental or otherwise)
    # file can't exhaust memory/disk before the upload route's own
    # validation ever runs.
    app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

    # SESSION_COOKIE_HTTPONLY defaults to True already (keeps the login
    # cookie out of reach of any injected JS). SAMESITE is set
    # explicitly rather than left to Flask's own default so it's a
    # documented decision, not an accident: "Lax" still lets a plain
    # top-level link into the tool work while blocking the cookie being
    # sent along with a cross-site POST/embed.
    app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

    # SESSION_COOKIE_SECURE makes the browser withhold the cookie over
    # plain HTTP - correct once this is ever served over HTTPS, but
    # today's documented deployment (run_server.py, local HTTP only)
    # would silently break: the cookie would never come back and login
    # would appear to fail for no visible reason. So this only turns on
    # when explicitly opted into via SESSION_COOKIE_SECURE=true, for the
    # day this runs behind real HTTPS.
    app.config["SESSION_COOKIE_SECURE"] = (
        os.environ.get("SESSION_COOKIE_SECURE", "").strip().lower() == "true"
    )

    if not os.path.exists(db_path):
        init_db(db_path)

    app.register_blueprint(web_blueprint)

    # Makes csrf_token() and t() (translation lookup - see .i18n)
    # callable directly in any template without every route having to
    # remember to pass them in explicitly.
    app.context_processor(lambda: {"csrf_token": get_csrf_token, "t": t, "languages": LANGUAGES})

    @app.route("/")
    def index():
        return redirect(url_for("web.upload"))

    @app.route("/lang", methods=["POST"])
    def set_language():
        validate_csrf_token(request.form.get("csrf_token"))
        code = request.form.get("code", "")
        if code not in LANGUAGES:
            abort(400, description="Unknown language.")
        session["lang"] = code
        referrer = request.referrer
        if referrer and referrer.startswith(request.host_url):
            return redirect(referrer)
        return redirect(url_for("web.upload") if session.get("user_email") else url_for("web.login"))

    def _fallback_continue_url():
        return url_for("web.upload") if session.get("user_email") else url_for("web.login")

    def _continue_url():
        # Prefer sending them right back to the page they were on (a
        # fresh GET there regenerates a valid CSRF token automatically,
        # so a retry just works) - only when it's actually this app's
        # own page, never an external URL a crafted Referer could point
        # at, and never the same request path that just failed (a POST
        # endpoint has no matching GET page to land back on safely).
        referrer = request.referrer
        if referrer and referrer.startswith(request.host_url) and referrer.rstrip("/") != request.url.rstrip("/"):
            return referrer
        return _fallback_continue_url()

    @app.errorhandler(400)
    def handle_bad_request(error):
        # Most often this is the stale-CSRF-token case (see
        # app.web.csrf.validate_csrf_token) - real and reproducible any
        # time the session's token has changed (every logout does this)
        # while an old copy of a form is still sitting in the browser
        # (a second tab, the back button). Flask's own default 400 page
        # is bare HTML with no nav and no way back except retyping the
        # URL - to a non-technical user that reads as "the login is
        # broken," not "reload and try again." This renders the same
        # message inside the app's own layout, with a link that either
        # returns to the form they were just on (so a retry just works)
        # or a sensible fallback.
        return render_template(
            "error.html",
            heading=t("error.something_wrong"),
            message=error.description or "That request couldn't be processed.",
            continue_url=_continue_url(),
            continue_label=t("action.try_again"),
        ), 400

    @app.errorhandler(404)
    def handle_not_found(error):
        return render_template(
            "error.html",
            heading=t("error.not_found"),
            message=error.description or "That page or item doesn't exist.",
            continue_url=_fallback_continue_url(),
            continue_label=t("action.back_to_upload") if session.get("user_email") else t("action.back_to_login"),
        ), 404

    return app
