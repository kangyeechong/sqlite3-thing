"""
Permanently removes one or more contracts (and every commission event
tied to them) from the ledger by PO number - a one-time cleanup tool
for data that should never have been imported at all (e.g. a sample/
test file accidentally uploaded into a real database).

This is NOT how to handle a real cancelled or withdrawn PO - those
stay in the ledger forever and show up beige in the report (see
docs/data_model.md). This script is only for rows that were never real
sales data in the first place.

Usage:
    python3 remove_contracts.py --db ledger.db --po 90101 90102 90103
    python3 remove_contracts.py --db ledger.db --customer-id-prefix TEST

Always prints exactly what will be removed and asks for typed
confirmation before deleting anything.
"""

import argparse

from app.db.connection import get_connection


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--db", required=True, help="Path to the ledger database file")
    parser.add_argument("--po", type=int, nargs="*", default=[], help="One or more PO numbers to remove")
    parser.add_argument(
        "--customer-id-prefix", default=None,
        help="Also remove every contract whose Customer ID starts with this prefix (e.g. TEST)",
    )
    args = parser.parse_args()

    if not args.po and not args.customer_id_prefix:
        raise SystemExit("Nothing to do - pass --po and/or --customer-id-prefix.")

    conn = get_connection(args.db)
    try:
        po_nos = set(args.po)
        if args.customer_id_prefix:
            matches = conn.execute(
                "SELECT po_no FROM contracts WHERE customer_id LIKE ?",
                (args.customer_id_prefix + "%",),
            ).fetchall()
            po_nos.update(row["po_no"] for row in matches)

        if not po_nos:
            print("No matching contracts found - nothing to remove.")
            return

        placeholders = ",".join("?" for _ in po_nos)
        rows = conn.execute(
            f"SELECT po_no, customer_id, agent_name, agency_code "
            f"FROM contracts WHERE po_no IN ({placeholders})",
            tuple(po_nos),
        ).fetchall()

        if not rows:
            print("No matching contracts found - nothing to remove.")
            return

        print(f"About to permanently remove {len(rows)} contract(s) and every commission event tied to them:")
        for r in rows:
            print(f"  PO {r['po_no']}  customer={r['customer_id']}  agent={r['agent_name']}  agency={r['agency_code']}")

        confirm = input("\nType 'yes' to permanently delete these: ").strip().lower()
        if confirm != "yes":
            print("Cancelled - nothing was deleted.")
            return

        conn.execute(f"DELETE FROM commission_events WHERE po_no IN ({placeholders})", tuple(po_nos))
        cursor = conn.execute(f"DELETE FROM contracts WHERE po_no IN ({placeholders})", tuple(po_nos))
        conn.commit()
        print(f"Removed {cursor.rowcount} contract(s).")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
