"""
Minimal, hand-rolled CSRF protection - a random per-session token
embedded in every form and checked on every POST. No extra dependency:
this tool has exactly two POST forms (login, upload), not enough
surface to justify pulling in a full CSRF library for.
"""

import secrets

from flask import abort, session


def get_csrf_token():
    """Returns this session's token, generating one on first use."""
    if "csrf_token" not in session:
        session["csrf_token"] = secrets.token_hex(32)
    return session["csrf_token"]


def validate_csrf_token(submitted_token):
    """Aborts the request with 400 if the submitted token doesn't match
    this session's token. secrets.compare_digest avoids leaking timing
    information about how much of the token matched."""
    expected = session.get("csrf_token")
    if not expected or not submitted_token or not secrets.compare_digest(expected, submitted_token):
        abort(400, description="Your session expired, or this form was submitted from somewhere unexpected. Please reload the page and try again.")
