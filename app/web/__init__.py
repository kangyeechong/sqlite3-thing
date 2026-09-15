"""
Flask app factory for the Agent Commission Tool web interface.

This wraps the already-proven ledger logic (app.importer, app.commission,
app.report) - it contains no business logic of its own, only HTTP
plumbing. See app/web/routes.py.
"""

import os
import sys
import warnings

from flask import Flask, redirect, url_for

from ..db.connection import init_db
from .csrf import get_csrf_token
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

    if not os.path.exists(db_path):
        init_db(db_path)

    app.register_blueprint(web_blueprint)

    # Makes csrf_token() callable directly in any template without
    # every route having to remember to pass it in explicitly.
    app.context_processor(lambda: {"csrf_token": get_csrf_token})

    @app.route("/")
    def index():
        return redirect(url_for("web.upload"))

    return app
