"""
A single Jinja filter (`fmtdate`) that renders a stored ISO date or
timestamp string as DD/MM/YYYY - the business's own convention
everywhere else (confirmed against the real file's Remarks -
"Inurnment: 04/06/2026" - and its historical "As at DD/MM/YYYY"
summary rows; see app.report._DATE_FORMAT's comment for the same
reasoning applied to the downloaded Excel reports). Every date is
still stored as ISO (YYYY-MM-DD, or a full isoformat() timestamp) in
the database - this only changes how it's displayed on a page, never
how it's stored or compared.
"""

import datetime


def fmtdate(value):
    """
    Accepts whatever a template might actually pass: a plain ISO date
    string ("2026-08-27"), a full isoformat() timestamp
    ("2026-08-27T14:03:21.123456"), None, or an empty string - any of
    which show up here depending on which column it came from. Falls
    back to returning the value unchanged for anything that doesn't
    parse as either, rather than raising - a stale or blank value
    should never turn into a 500 error just because a page tried to
    pretty-print it.
    """
    if not value:
        return value
    try:
        return datetime.date.fromisoformat(value).strftime("%d/%m/%Y")
    except ValueError:
        pass
    try:
        return datetime.datetime.fromisoformat(value).strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return value
