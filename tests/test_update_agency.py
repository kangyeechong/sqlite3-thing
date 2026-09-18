"""
update_agency.py: the manual fix for an agency whose group/split-type
was never seeded correctly (or predates a rules.py change) - the
importer only ever seeds these on first sight of a new agency_code and
never touches them again (see app/importer.py), so a wrong value can't
self-heal from a later upload.

Run with: pytest tests/test_update_agency.py -v
"""

import subprocess
import sys

from app.db.connection import get_connection, init_db


def _seed(db_path):
    init_db(db_path)
    conn = get_connection(db_path)
    conn.execute(
        "INSERT INTO agencies (agency_code, splits_by_agent, commission_split_type, agency_group) "
        "VALUES ('AC108-01', 1, 'flat', NULL)"
    )
    conn.commit()
    conn.close()


def _run(db_path, args, confirm=True):
    return subprocess.run(
        [sys.executable, "update_agency.py", "--db", db_path, *args],
        input="yes\n" if confirm else "no\n",
        capture_output=True, text=True,
    )


def test_sets_group_and_split_type_on_an_existing_agency(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    _seed(db_path)

    result = _run(db_path, [
        "--agency-code", "AC108-01", "--group", "AW Consultancy",
        "--split-type", "agency_agent_split",
    ])
    assert result.returncode == 0
    assert "Updated." in result.stdout

    conn = get_connection(db_path)
    row = conn.execute("SELECT * FROM agencies WHERE agency_code = 'AC108-01'").fetchone()
    conn.close()
    assert row["agency_group"] == "AW Consultancy"
    assert row["commission_split_type"] == "agency_agent_split"


def test_declining_confirmation_changes_nothing(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    _seed(db_path)

    _run(db_path, ["--agency-code", "AC108-01", "--group", "AW Consultancy"], confirm=False)

    conn = get_connection(db_path)
    row = conn.execute("SELECT agency_group FROM agencies WHERE agency_code = 'AC108-01'").fetchone()
    conn.close()
    assert row["agency_group"] is None


def test_unknown_agency_code_errors_clearly(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    init_db(db_path)

    result = _run(db_path, ["--agency-code", "DOES-NOT-EXIST", "--group", "Something"])
    assert result.returncode != 0
    assert "No agency" in result.stderr
