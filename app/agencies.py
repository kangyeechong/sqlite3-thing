"""
Agency management - the small set of things staff configure once per
agency rather than every PO import re-deriving them: its human-readable
name, and which of three confirmed commission report "formats" it
follows (the actual percentages behind each format live in rules.py -
this module never computes commission itself).

Every agency already gets auto-seeded into the agencies table the
first time its code shows up on a Master report import (see
app.importer), defaulting to the plainest format unless it's one of
the small number of exceptions hardcoded in rules.py. This module is
what lets staff set the real format directly instead - either ahead of
time (so a brand-new agency imports correctly from day one) or after
the fact (a re-classification, which only ever affects future
commission detections - see commission._build_event; an
already-raised commission_events row is never recalculated).
"""

import sqlite3

# Every agency in this system follows exactly one of these three named
# formats - confirmed with the business, named after the real agency
# each one was first observed on. Each maps onto the two columns
# `agencies` already had (splits_by_agent, commission_split_type)
# rather than needing a schema change: "Lachesis-style" turned out to
# be exactly the combination a brand-new, not-yet-classified agency
# already defaults to (full commission paid to the agency, but still
# broken into per-agent sheets in the report) - this just gives staff
# a name to pick deliberately instead of only ever getting it by
# default.
AGENCY_FORMATS = {
    "xemp": {
        "label": "Single Table (XEMP-style)",
        "commission_split_type": "flat",
        "splits_by_agent": 0,
    },
    "lachesis": {
        "label": "Full Payout, Agent Breakdown (Lachesis-style)",
        "commission_split_type": "flat",
        "splits_by_agent": 1,
    },
    "aw": {
        "label": "Agency/Agent Split (AW Consultancy-style)",
        "commission_split_type": "agency_agent_split",
        "splits_by_agent": 1,
    },
}

# Reverse lookup - (commission_split_type, splits_by_agent) -> format
# key - used to pre-select an existing agency's format on the edit
# form, built once rather than searched linearly on every request.
_FORMAT_BY_COLUMNS = {
    (fmt["commission_split_type"], fmt["splits_by_agent"]): key
    for key, fmt in AGENCY_FORMATS.items()
}


def format_key_for_agency(agency_row):
    """
    Which AGENCY_FORMATS key an existing agencies row matches, or None
    if its (commission_split_type, splits_by_agent) combination doesn't
    correspond to any of the three named formats - shouldn't happen for
    anything created or edited through this module, but nothing stops
    an older row (or a direct database edit) from carrying a
    combination that predates these three names.
    """
    return _FORMAT_BY_COLUMNS.get((agency_row["commission_split_type"], agency_row["splits_by_agent"]))


def list_agencies(conn):
    """Every agency, alphabetical by code - what the Agencies page lists."""
    return conn.execute("SELECT * FROM agencies ORDER BY agency_code").fetchall()


def create_agency(conn, agency_code, name, format_key):
    """
    Adds a brand-new agency. Raises ValueError (never a raw
    sqlite3.IntegrityError) if the code's already taken - agency_code
    is this table's primary key, and every contract ties back to it, so
    a silent overwrite here would be a real data mistake, not a
    convenience - or if format_key isn't one of AGENCY_FORMATS.
    """
    if format_key not in AGENCY_FORMATS:
        raise ValueError(f"Unknown format {format_key!r}.")
    fmt = AGENCY_FORMATS[format_key]
    try:
        conn.execute(
            "INSERT INTO agencies (agency_code, name, splits_by_agent, commission_split_type) "
            "VALUES (?, ?, ?, ?)",
            (agency_code, name, fmt["splits_by_agent"], fmt["commission_split_type"]),
        )
    except sqlite3.IntegrityError:
        raise ValueError(f"Agency code {agency_code!r} already exists.")


def update_agency(conn, current_agency_code, new_agency_code, name, format_key):
    """
    Updates an existing agency's name and format, and - since staff
    explicitly need this to be possible - its code too. Checked
    directly against schema.sql: only contracts.agency_code references
    agencies.agency_code, so a rename is done as insert the new code /
    move every contract onto it / delete the old code, all within the
    caller's transaction - never a raw UPDATE of the primary key
    itself, which SQLite's foreign key enforcement (PRAGMA
    foreign_keys = ON - see db.connection.get_connection) would reject
    outright while any contract still points at the old code.

    Raises ValueError if current_agency_code doesn't exist, the new
    code collides with a different existing agency, or format_key isn't
    recognized.

    Only ever affects FUTURE commission detections - an already-raised
    commission_events row was computed once, at raise time, and is
    never recalculated (see commission._build_event). Re-classifying an
    agency here doesn't retroactively fix anything already confirmed or
    still pending review.
    """
    if format_key not in AGENCY_FORMATS:
        raise ValueError(f"Unknown format {format_key!r}.")
    current = conn.execute(
        "SELECT * FROM agencies WHERE agency_code = ?", (current_agency_code,)
    ).fetchone()
    if current is None:
        raise ValueError(f"No agency with code {current_agency_code!r}.")

    fmt = AGENCY_FORMATS[format_key]

    if new_agency_code != current_agency_code:
        collision = conn.execute(
            "SELECT 1 FROM agencies WHERE agency_code = ?", (new_agency_code,)
        ).fetchone()
        if collision is not None:
            raise ValueError(f"Agency code {new_agency_code!r} already exists.")
        # agency_group is carried over as-is (a free-text group label,
        # not itself tied to the code value) rather than exposed on
        # this form - see the planning discussion for why it's left out
        # for now.
        conn.execute(
            "INSERT INTO agencies (agency_code, name, splits_by_agent, commission_split_type, agency_group) "
            "VALUES (?, ?, ?, ?, ?)",
            (new_agency_code, name, fmt["splits_by_agent"], fmt["commission_split_type"], current["agency_group"]),
        )
        conn.execute(
            "UPDATE contracts SET agency_code = ? WHERE agency_code = ?",
            (new_agency_code, current_agency_code),
        )
        conn.execute("DELETE FROM agencies WHERE agency_code = ?", (current_agency_code,))
    else:
        conn.execute(
            "UPDATE agencies SET name = ?, splits_by_agent = ?, commission_split_type = ? WHERE agency_code = ?",
            (name, fmt["splits_by_agent"], fmt["commission_split_type"], current_agency_code),
        )
