"""
Best-effort parsing of the free-text `Remarks` column from the Kenjin
Master report.

IMPORTANT LIMITATION: Remarks is text a staff member typed by hand, not
a structured field. Kenjin's export has no dedicated "status" or "case
type" column at all - Remarks is the only signal available for these.
The functions here match on keyword patterns confirmed against real
sample data ("Cancelled PO", "WITHDRAWAL NOTICE RECEIVED...", "At need
case"). They will not catch every possible phrasing a human might type.

Every function below documents which direction it fails in when it
doesn't recognize a phrasing, because that direction matters:
  - Failing toward "active"/"pre_need" (the defaults) is the SAFE
    direction for case type - it just means a cooling-off wait gets
    applied when it technically didn't need to be, delaying a
    commission rather than releasing one too early.
  - Failing toward "active" for status is the RISKY direction - a real
    cancellation that isn't recognized would incorrectly stay eligible
    for commission. That's why importer.py separately flags every
    detected status CHANGE in the review panel, so a human confirms it
    rather than trusting this module silently.
"""

import datetime
import re

_CANCELLED_KEYWORDS = ("cancelled", "cancel")
_WITHDRAWAL_KEYWORDS = ("withdrawal", "withdrawn", "withdraw")
_AT_NEED_KEYWORDS = ("at need",)

# Matches "Inurnment on 04/06/2026" or "Inurnment on\n04/06/2026" or
# "Inurnment: 04/06/2026", case-insensitive, DD/MM/YYYY.
_INURNMENT_DATE_PATTERN = re.compile(
    r"inurnment\s*(?:on)?\s*[:\-]?\s*[\r\n]*\s*(\d{1,2})/(\d{1,2})/(\d{4})",
    re.IGNORECASE,
)


def detect_status(remarks):
    """
    Returns 'cancelled', 'withdrawn', or 'active' based on keywords in
    Remarks. Never returns 'on_hold' - that has no signal in the
    Kenjin export and must be set manually within this tool.
    """
    if not remarks:
        return "active"
    lowered = remarks.lower()
    if any(kw in lowered for kw in _CANCELLED_KEYWORDS):
        return "cancelled"
    if any(kw in lowered for kw in _WITHDRAWAL_KEYWORDS):
        return "withdrawn"
    return "active"


def detect_case_type(remarks):
    """Returns 'at_need' or 'pre_need' based on keywords in Remarks."""
    if not remarks:
        return "pre_need"
    lowered = remarks.lower()
    if any(kw in lowered for kw in _AT_NEED_KEYWORDS):
        return "at_need"
    return "pre_need"


def extract_inurnment_date(remarks):
    """
    Pulls a DD/MM/YYYY date out of Remarks text and returns it as an
    ISO string (YYYY-MM-DD), or None if no date is found. Callers must
    treat None on an At-Need contract as something to flag, not
    silently accept - an At-Need contract without an inurnment date on
    file should not have its full-payment commission released yet.
    """
    if not remarks:
        return None
    match = _INURNMENT_DATE_PATTERN.search(remarks)
    if not match:
        return None
    day, month, year = (int(g) for g in match.groups())
    try:
        # Rejects garbage/placeholder text that merely looks like a
        # date (e.g. "99/99/9999" typed as a TBC placeholder) - such a
        # match must NOT be treated as a real inurnment date, since
        # that would incorrectly let an At-Need contract's full-payment
        # commission release immediately.
        return datetime.date(year, month, day).isoformat()
    except ValueError:
        return None
