"""
Manually corrects one agency's settings (agency_group,
commission_split_type, splits_by_agent) - the importer only ever seeds
these the FIRST time it sees a new agency_code, then never touches
them again on later uploads (see app/importer.py), so a wrong or
missing value - e.g. from before a code's grouping was set up
correctly, or an agency row that predates a schema/rules change - has
to be fixed by hand here rather than by re-uploading.

Usage:
    python3 update_agency.py --db ledger.db --agency-code AC108-01 --group "AW Consultancy" --split-type agency_agent_split
    python3 update_agency.py --db ledger.db --agency-code AC200 --group ""   (clears the group)

Always prints the agency's current settings and the change about to be
made before applying it - nothing is changed unless at least one of
--group / --split-type / --splits-by-agent is passed.
"""

import argparse

from app.db.connection import get_connection


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--db", required=True, help="Path to the ledger database file")
    parser.add_argument("--agency-code", required=True)
    parser.add_argument(
        "--group", default=None,
        help="Agency group name (several agency_codes sharing one group get one combined sheet, "
             "e.g. \"AW Consultancy\"). Pass an empty string to clear it.",
    )
    parser.add_argument(
        "--split-type", default=None, choices=["flat", "agency_agent_split"],
        help="'flat' (one commission figure) or 'agency_agent_split' (separate agency/agent shares).",
    )
    parser.add_argument(
        "--splits-by-agent", default=None, choices=["true", "false"],
        help="Whether this agency's sheet also breaks down into one sheet per individual agent.",
    )
    args = parser.parse_args()

    if args.group is None and args.split_type is None and args.splits_by_agent is None:
        raise SystemExit("Nothing to do - pass at least one of --group / --split-type / --splits-by-agent.")

    conn = get_connection(args.db)
    try:
        row = conn.execute(
            "SELECT agency_code, agency_group, commission_split_type, splits_by_agent "
            "FROM agencies WHERE agency_code = ?",
            (args.agency_code,),
        ).fetchone()
        if row is None:
            raise SystemExit(
                f"No agency '{args.agency_code}' on file yet - it's only created the first time a "
                f"contract with that Agency Code is imported."
            )

        print(f"Agency {row['agency_code']} - current settings:")
        print(f"  agency_group        = {row['agency_group']!r}")
        print(f"  commission_split_type = {row['commission_split_type']!r}")
        print(f"  splits_by_agent     = {bool(row['splits_by_agent'])}")

        new_group = row["agency_group"] if args.group is None else (args.group or None)
        new_split_type = row["commission_split_type"] if args.split_type is None else args.split_type
        new_splits_by_agent = (
            row["splits_by_agent"] if args.splits_by_agent is None
            else (1 if args.splits_by_agent == "true" else 0)
        )

        print(f"\nWill update to:")
        print(f"  agency_group        = {new_group!r}")
        print(f"  commission_split_type = {new_split_type!r}")
        print(f"  splits_by_agent     = {bool(new_splits_by_agent)}")

        confirm = input("\nType 'yes' to apply this change: ").strip().lower()
        if confirm != "yes":
            print("Cancelled - nothing was changed.")
            return

        conn.execute(
            "UPDATE agencies SET agency_group = ?, commission_split_type = ?, splits_by_agent = ? "
            "WHERE agency_code = ?",
            (new_group, new_split_type, new_splits_by_agent, args.agency_code),
        )
        conn.commit()
        print("Updated.")
    finally:
        conn.close()


if __name__ == "__main__":
    main()
