"""
Creates (or resets the password for) a login account.

Usage:
    python3 create_user.py --db ledger.db --email jane@xekl.com --name "Jane Tan"

You'll be prompted for a password interactively (not passed as a
command-line argument, so it never ends up in shell history).
"""

import argparse
import getpass

from app.db.connection import get_connection, init_db
from app.web.auth import hash_password


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--db", required=True, help="Path to the ledger database file")
    parser.add_argument("--email", required=True)
    parser.add_argument("--name", default=None, help="Display name (optional)")
    args = parser.parse_args()

    init_db(args.db)  # no-op if the tables already exist

    password = getpass.getpass("Password: ")
    confirm = getpass.getpass("Confirm password: ")
    if password != confirm:
        raise SystemExit("Passwords didn't match - nothing was saved.")
    if len(password) < 8:
        raise SystemExit("Password must be at least 8 characters - nothing was saved.")

    email = args.email.strip().lower()

    conn = get_connection(args.db)
    try:
        conn.execute(
            "INSERT INTO users (email, password_hash, display_name) VALUES (?, ?, ?) "
            "ON CONFLICT(email) DO UPDATE SET "
            "password_hash = excluded.password_hash, display_name = excluded.display_name",
            (email, hash_password(password), args.name),
        )
        conn.commit()
    finally:
        conn.close()

    print(f"User '{email}' created/updated.")


if __name__ == "__main__":
    main()
