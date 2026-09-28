"""
Tests for app.web.dates - the DD/MM/YYYY <-> ISO conversion helpers
used at the web layer's boundary (see that module's docstring for why
they exist: the native date-picker's displayed format turns out to
follow the browser's own language setting, not anything this app's
code can force, so every date field a person types into is now a
plain text field instead).
"""

import pytest

from app.web.dates import fmtdate, parse_ddmmyyyy


def test_fmtdate_formats_a_plain_iso_date():
    assert fmtdate("2026-08-27") == "27/08/2026"


def test_fmtdate_formats_a_full_iso_timestamp():
    assert fmtdate("2026-08-27T14:03:21.123456") == "27/08/2026 14:03"


def test_fmtdate_passes_through_none_and_blank_unchanged():
    assert fmtdate(None) is None
    assert fmtdate("") == ""


def test_fmtdate_passes_through_unparseable_text_unchanged():
    # Never crash on a stale/unexpected value - see the module docstring.
    assert fmtdate("not a date") == "not a date"


def test_parse_ddmmyyyy_converts_to_iso():
    assert parse_ddmmyyyy("27/08/2026") == "2026-08-27"


def test_parse_ddmmyyyy_strips_surrounding_whitespace():
    assert parse_ddmmyyyy("  27/08/2026  ") == "2026-08-27"


def test_parse_ddmmyyyy_rejects_the_wrong_shape():
    with pytest.raises(ValueError):
        parse_ddmmyyyy("2026-08-27")  # ISO, not DD/MM/YYYY
    with pytest.raises(ValueError):
        parse_ddmmyyyy("8/27/2026")  # single-digit month, US order
    with pytest.raises(ValueError):
        parse_ddmmyyyy("")


def test_parse_ddmmyyyy_rejects_a_shape_that_matches_but_isnt_a_real_date():
    with pytest.raises(ValueError):
        parse_ddmmyyyy("31/02/2026")  # February never has a 31st
    with pytest.raises(ValueError):
        parse_ddmmyyyy("00/01/2026")
