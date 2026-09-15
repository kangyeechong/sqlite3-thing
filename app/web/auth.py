"""
Login and session helpers. Individual accounts, hashed passwords - no
shared password, per the original brief (this tool touches real payout
amounts, so who-processed-what matters).
"""

from functools import wraps

from flask import redirect, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash


def hash_password(plain_password):
    return generate_password_hash(plain_password)


def verify_password(plain_password, password_hash):
    if not password_hash:
        # users.password_hash is nullable in the schema - without this
        # guard, a row that somehow ended up with no hash would crash
        # check_password_hash() with an AttributeError instead of
        # simply failing the login attempt.
        return False
    return check_password_hash(password_hash, plain_password)


def find_user_by_email(conn, email):
    return conn.execute("SELECT * FROM users WHERE email = ?", (email,)).fetchone()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_email" not in session:
            return redirect(url_for("web.login"))
        return view(*args, **kwargs)
    return wrapped
