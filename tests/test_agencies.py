"""
Agency management: staff set an agency's name and commission report
format directly (see app/agencies.py for the three confirmed formats
and why "Lachesis-style" needs no schema change), instead of every new
agency needing a code change in rules.py.

Run with: pytest tests/test_agencies.py -v
"""

import pytest

from app.agencies import create_agency, format_key_for_agency, list_agencies, update_agency
from app.db.connection import get_connection, init_db


def _db_path(tmp_path):
    return str(tmp_path / "ledger.db")


def _conn(tmp_path):
    db_path = _db_path(tmp_path)
    init_db(db_path)
    return get_connection(db_path)


# --- format_key_for_agency: the reverse lookup -------------------------

@pytest.mark.parametrize("commission_split_type,splits_by_agent,expected", [
    ("flat", 0, "xemp"),
    ("flat", 1, "lachesis"),
    ("agency_agent_split", 1, "aw"),
    ("agency_agent_split", 0, None),  # not a real combination any format produces
])
def test_format_key_for_agency(commission_split_type, splits_by_agent, expected):
    row = {"commission_split_type": commission_split_type, "splits_by_agent": splits_by_agent}
    assert format_key_for_agency(row) == expected


# --- create_agency -------------------------------------------------------

def test_create_agency_sets_name_and_format_columns(tmp_path):
    conn = _conn(tmp_path)
    create_agency(conn, "AC300", "New Agency Sdn Bhd", "aw")
    conn.commit()

    row = conn.execute("SELECT * FROM agencies WHERE agency_code = 'AC300'").fetchone()
    assert row["name"] == "New Agency Sdn Bhd"
    assert row["commission_split_type"] == "agency_agent_split"
    assert row["splits_by_agent"] == 1
    conn.close()


def test_create_agency_lachesis_format(tmp_path):
    """The whole point: "Lachesis-style" is just flat + splits_by_agent
    on - no new commission_split_type value."""
    conn = _conn(tmp_path)
    create_agency(conn, "AC301", "Lachesis Marketing", "lachesis")
    conn.commit()

    row = conn.execute("SELECT * FROM agencies WHERE agency_code = 'AC301'").fetchone()
    assert row["commission_split_type"] == "flat"
    assert row["splits_by_agent"] == 1
    conn.close()


def test_create_agency_with_a_duplicate_code_raises_value_error(tmp_path):
    conn = _conn(tmp_path)
    create_agency(conn, "AC302", "First", "xemp")
    conn.commit()

    with pytest.raises(ValueError, match="already exists"):
        create_agency(conn, "AC302", "Second", "lachesis")
    conn.close()


def test_create_agency_with_an_unknown_format_raises_value_error(tmp_path):
    conn = _conn(tmp_path)
    with pytest.raises(ValueError, match="Unknown format"):
        create_agency(conn, "AC303", "Some Agency", "not-a-real-format")
    conn.close()


# --- update_agency ---------------------------------------------------------

def test_update_agency_changes_name_and_format_without_touching_code(tmp_path):
    conn = _conn(tmp_path)
    create_agency(conn, "AC304", "Old Name", "xemp")
    conn.commit()

    update_agency(conn, "AC304", "AC304", "New Name", "aw")
    conn.commit()

    row = conn.execute("SELECT * FROM agencies WHERE agency_code = 'AC304'").fetchone()
    assert row["name"] == "New Name"
    assert row["commission_split_type"] == "agency_agent_split"
    assert row["splits_by_agent"] == 1
    conn.close()


def test_update_agency_renaming_the_code_moves_every_contract(tmp_path):
    """The core safety property: an agency code is a foreign key every
    contract points at (PRAGMA foreign_keys = ON), so a rename can't
    just UPDATE the primary key in place - it has to move every
    referencing contract over first."""
    conn = _conn(tmp_path)
    create_agency(conn, "AC305", "Renaming Agency", "xemp")
    conn.execute(
        "INSERT INTO contracts (po_no, agency_code, net_price, case_type, status) "
        "VALUES (70500, 'AC305', 10000, 'pre_need', 'active')"
    )
    conn.execute(
        "INSERT INTO contracts (po_no, agency_code, net_price, case_type, status) "
        "VALUES (70501, 'AC305', 10000, 'pre_need', 'active')"
    )
    conn.commit()

    update_agency(conn, "AC305", "AC305-NEW", "Renaming Agency", "xemp")
    conn.commit()

    assert conn.execute("SELECT 1 FROM agencies WHERE agency_code = 'AC305'").fetchone() is None
    new_row = conn.execute("SELECT * FROM agencies WHERE agency_code = 'AC305-NEW'").fetchone()
    assert new_row is not None
    assert new_row["name"] == "Renaming Agency"

    moved_pos = {
        row["po_no"] for row in conn.execute("SELECT po_no FROM contracts WHERE agency_code = 'AC305-NEW'")
    }
    assert moved_pos == {70500, 70501}
    assert conn.execute("SELECT 1 FROM contracts WHERE agency_code = 'AC305'").fetchone() is None
    conn.close()


def test_update_agency_rename_preserves_agency_group(tmp_path):
    conn = _conn(tmp_path)
    conn.execute(
        "INSERT INTO agencies (agency_code, name, splits_by_agent, commission_split_type, agency_group) "
        "VALUES ('AC108-04', 'AW Consultancy Agent 4', 1, 'agency_agent_split', 'AW Consultancy')"
    )
    conn.commit()

    update_agency(conn, "AC108-04", "AC108-04-RENAMED", "AW Consultancy Agent 4", "aw")
    conn.commit()

    row = conn.execute("SELECT * FROM agencies WHERE agency_code = 'AC108-04-RENAMED'").fetchone()
    assert row["agency_group"] == "AW Consultancy"
    conn.close()


def test_update_agency_rename_colliding_with_another_agency_raises_value_error(tmp_path):
    conn = _conn(tmp_path)
    create_agency(conn, "AC306", "Agency A", "xemp")
    create_agency(conn, "AC307", "Agency B", "xemp")
    conn.commit()

    with pytest.raises(ValueError, match="already exists"):
        update_agency(conn, "AC306", "AC307", "Agency A", "xemp")
    conn.close()


def test_update_agency_with_a_nonexistent_current_code_raises_value_error(tmp_path):
    conn = _conn(tmp_path)
    with pytest.raises(ValueError, match="No agency"):
        update_agency(conn, "AC999999", "AC999999", "Doesn't Exist", "xemp")
    conn.close()


def test_update_agency_with_an_unknown_format_raises_value_error(tmp_path):
    conn = _conn(tmp_path)
    create_agency(conn, "AC308", "Agency C", "xemp")
    conn.commit()

    with pytest.raises(ValueError, match="Unknown format"):
        update_agency(conn, "AC308", "AC308", "Agency C", "not-a-real-format")
    conn.close()


# --- list_agencies ---------------------------------------------------------

def test_list_agencies_is_alphabetical_by_code(tmp_path):
    conn = _conn(tmp_path)
    create_agency(conn, "AC999", "Zed Agency", "xemp")
    create_agency(conn, "AC100", "Alpha Agency", "xemp")
    conn.commit()

    codes = [row["agency_code"] for row in list_agencies(conn)]
    assert codes.index("AC100") < codes.index("AC999")
    conn.close()
