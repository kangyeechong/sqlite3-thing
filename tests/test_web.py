"""
Step 4 verification: the web layer (login, upload, results, download).

Run with: pytest tests/test_web.py -v
"""

import datetime
import io
import os
import re

import openpyxl
import pytest

from app.db.connection import get_connection
from app.web import create_app
from app.web.auth import hash_password, verify_password
from tests.helpers import build_master_report


def test_verify_password_with_no_stored_hash_returns_false_not_a_crash():
    """
    users.password_hash is nullable in the schema - verify_password
    must fail closed on a missing hash, not raise.
    """
    assert verify_password("anything", None) is False
    assert verify_password("anything", "") is False


@pytest.fixture
def app(tmp_path):
    db_path = str(tmp_path / "ledger.db")
    application = create_app(db_path, secret_key="test-secret-not-for-real-use")
    application.config["TESTING"] = True

    conn = get_connection(db_path)
    conn.execute(
        "INSERT INTO users (email, password_hash, display_name) VALUES (?, ?, ?)",
        ("staff@xekl.com", hash_password("correct-horse-battery"), "Staff Member"),
    )
    conn.commit()
    conn.close()

    return application


@pytest.fixture
def client(app):
    return app.test_client()


def _csrf_token(client, get_path):
    """Fetches a page and pulls the csrf_token hidden field out of it -
    every real form submission needs a token from a page the same
    session actually rendered."""
    page = client.get(get_path)
    match = re.search(rb'name="csrf_token" value="([^"]+)"', page.data)
    assert match is not None, f"No csrf_token field found on {get_path}"
    return match.group(1).decode()


def _login(client, email="staff@xekl.com", password="correct-horse-battery"):
    token = _csrf_token(client, "/login")
    return client.post(
        "/login",
        data={"email": email, "password": password, "csrf_token": token},
        follow_redirects=True,
    )


def _upload(client, file_path, filename="upload.xlsx"):
    token = _csrf_token(client, "/upload")
    with open(file_path, "rb") as f:
        return client.post(
            "/upload",
            data={"report_file": (f, filename), "csrf_token": token},
            content_type="multipart/form-data",
        )


def test_upload_page_requires_login(client):
    response = client.get("/upload")
    assert response.status_code == 302
    assert "/login" in response.headers["Location"]


def test_wrong_password_is_rejected(client):
    token = _csrf_token(client, "/login")
    response = client.post(
        "/login",
        data={"email": "staff@xekl.com", "password": "wrong-password", "csrf_token": token},
    )
    assert response.status_code == 401
    assert b"Incorrect" in response.data


def test_unknown_email_is_rejected(client):
    token = _csrf_token(client, "/login")
    response = client.post(
        "/login",
        data={"email": "nobody@xekl.com", "password": "anything", "csrf_token": token},
    )
    assert response.status_code == 401


def test_login_without_a_valid_csrf_token_is_rejected(client):
    response = client.post(
        "/login",
        data={"email": "staff@xekl.com", "password": "correct-horse-battery", "csrf_token": "made-up-token"},
    )
    assert response.status_code == 400


def test_correct_login_reaches_upload_page(client):
    response = _login(client)
    assert response.status_code == 200
    assert b"Upload Commission Base Report" in response.data


def test_logout_blocks_further_access(client):
    _login(client)
    client.get("/logout")
    response = client.get("/upload")
    assert response.status_code == 302


def test_full_flow_upload_shows_results_and_download_produces_a_real_workbook(client, tmp_path):
    """
    End-to-end through the actual HTTP layer, not by calling the
    pipeline functions directly: log in, upload a fake file, read the
    rendered results page, follow the download link, and open the
    actual bytes that come back as a real Excel workbook.
    """
    _login(client)

    today = datetime.date.today()
    settlement_date = today - datetime.timedelta(days=6)
    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 50001, "Customer ID": "CUSTWEB1", "Customer Name": "Web Test Customer",
        "Niche/Tablet Price (RM)": 10000,
        "Full Settlement Paid Date": settlement_date,
        "Agency Code": "AC001",
    }])

    response = _upload(client, xlsx_path)

    assert response.status_code == 200
    assert b"1500.00" in response.data  # 15% of Net Price 10,000
    assert b"Full Payment" in response.data  # trigger_labels passed into the template
    assert b"Download Excel report" in response.data

    match = re.search(rb"/download/(\d+)", response.data)
    assert match is not None
    run_id = int(match.group(1))

    download_response = client.get(f"/download/{run_id}")
    assert download_response.status_code == 200
    assert download_response.mimetype == (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    workbook = openpyxl.load_workbook(io.BytesIO(download_response.data))
    assert "All" in workbook.sheetnames
    assert "AC001" in workbook.sheetnames


def test_upload_with_nothing_due_shows_message_and_no_download_link(client, tmp_path):
    _login(client)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 50002, "Customer ID": "CUSTWEB2", "Customer Name": "Web Test Customer 2",
        # nothing paid yet
    }])

    response = _upload(client, xlsx_path)

    assert response.status_code == 200
    assert b"Nothing newly due" in response.data
    assert b"Download Excel report" not in response.data


def test_upload_without_choosing_a_file_shows_a_clear_message(client):
    _login(client)
    token = _csrf_token(client, "/upload")
    response = client.post(
        "/upload", data={"csrf_token": token}, content_type="multipart/form-data",
    )
    assert response.status_code == 400
    assert b"Choose a Commission Base Report file" in response.data


def test_upload_without_a_valid_csrf_token_is_rejected(client, tmp_path):
    _login(client)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 50003, "Customer ID": "CUSTWEB3", "Customer Name": "Web Test Customer 3",
    }])

    with open(xlsx_path, "rb") as f:
        response = client.post(
            "/upload",
            data={"report_file": (f, "upload.xlsx"), "csrf_token": "made-up-token"},
            content_type="multipart/form-data",
        )
    assert response.status_code == 400


def test_malicious_filename_cannot_escape_the_temp_directory(client, tmp_path):
    """
    A filename crafted to look like a path-traversal or absolute-path
    attempt must be sanitized before it's ever used to build a file
    path on disk.
    """
    _login(client)

    xlsx_path = tmp_path / "upload.xlsx"
    build_master_report(xlsx_path, [{
        "No": 1, "PO No": 55001, "Customer ID": "CUSTSEC1", "Customer Name": "Security Test",
    }])

    response = _upload(client, xlsx_path, filename="../../../../tmp/evil_traversal_test.xlsx")

    # The point of this test is simply that the request completes
    # normally (the file gets processed inside the sandboxed temp
    # directory) rather than writing anywhere else on disk.
    assert response.status_code == 200
    assert not os.path.exists("/tmp/evil_traversal_test.xlsx")


def test_corrupted_upload_shows_a_friendly_message_not_a_server_error(client, tmp_path):
    not_really_excel = tmp_path / "fake.xlsx"
    not_really_excel.write_bytes(b"this is not a real xlsx file")

    _login(client)
    response = _upload(client, not_really_excel, filename="fake.xlsx")

    assert response.status_code == 400
    assert b"process this file" in response.data  # avoids the apostrophe, which Jinja2 HTML-escapes to &#39;


def test_downloading_a_nonexistent_run_id_returns_404(client):
    _login(client)
    response = client.get("/download/999999")
    assert response.status_code == 404


def test_upload_size_is_bounded(app):
    assert app.config["MAX_CONTENT_LENGTH"] == 20 * 1024 * 1024
