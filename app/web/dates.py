"""
Two small helpers that keep DD/MM/YYYY - the business's own
convention everywhere else (confirmed against the real file's Remarks
- "Inurnment: 04/06/2026" - and its historical "As at DD/MM/YYYY"
summary rows; see app.report._DATE_FORMAT's comment for the same
reasoning applied to the downloaded Excel reports) - at the edges of
the web layer only:

  - fmtdate: a Jinja filter that renders a stored ISO date/timestamp
    string as DD/MM/YYYY for display.
  - parse_ddmmyyyy: converts a DD/MM/YYYY string typed into a form
    field back to the ISO YYYY-MM-DD every other layer of this app
    (app.aor, app.report, app.pipeline, the database itself) stores
    and compares dates as.

Both exist because the native HTML date-picker's displayed format
turns out to follow each browser's own language setting (confirmed via
a real user report - Chrome showing MM/DD/YYYY even on a Malaysia-
based machine, because Chrome's language list had "English (United
States)" ahead of a UK/Malaysia one), not anything this app's own code
can force. Every date field a person actually types into is a plain
text input with a DD/MM/YYYY placeholder instead, guaranteed
consistent on every machine regardless of any browser setting - the
trade-off being no native calendar pop-up.
"""

import datetime
import re


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


_DDMMYYYY = re.compile(r"(\d{2})/(\d{2})/(\d{4})")


def parse_ddmmyyyy(value):
    """
    Parses a DD/MM/YYYY string into the ISO YYYY-MM-DD string every
    other layer of this app expects - the conversion happens here,
    once, at the point a route first reads it out of the request, so
    nothing below the web layer needs to know DD/MM/YYYY text input
    exists at all.

    Raises ValueError, with a message safe to show directly to
    whoever typed it, if `value` isn't in that exact shape, OR is that
    shape but not a real calendar date (31/02/2026) - never silently
    rounds or reinterprets it.
    """
    match = _DDMMYYYY.fullmatch((value or "").strip())
    if match is None:
        raise ValueError(f"{value!r} isn't a valid date - use DD/MM/YYYY.")
    day, month, year = (int(part) for part in match.groups())
    try:
        return datetime.date(year, month, day).isoformat()
    except ValueError:
        raise ValueError(f"{value!r} isn't a valid date - use DD/MM/YYYY.")
