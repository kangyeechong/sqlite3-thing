-- Schema for the Agent Commission Tool, Phase 1.
-- See docs/data_model.md for the full design rationale behind every
-- table and column here - this file is just the DDL.

PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS users (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    email         TEXT NOT NULL UNIQUE,
    password_hash TEXT,
    display_name  TEXT
);

CREATE TABLE IF NOT EXISTS customers (
    customer_id TEXT PRIMARY KEY,
    name        TEXT
);

CREATE TABLE IF NOT EXISTS agencies (
    agency_code     TEXT PRIMARY KEY,
    name            TEXT,
    -- Whether the Excel export further breaks this agency's sheet down
    -- by individual agent. Defaults to true; AC001 (XEMP, in-house
    -- staff - not an external agency) is the one confirmed exception.
    -- A data flag, not a hardcoded "if AW Consultancy" check, so a
    -- future exception is a data change, not a code change.
    splits_by_agent INTEGER NOT NULL DEFAULT 1 CHECK (splits_by_agent IN (0, 1))
);

-- One row per PO. This is the core ledger table - its three
-- *_commission_flagged columns are the "memory" that makes re-running
-- an upload safe: a contract already flagged never gets flagged again.
CREATE TABLE IF NOT EXISTS contracts (
    po_no          INTEGER PRIMARY KEY,
    customer_id    TEXT REFERENCES customers(customer_id),
    agent_name     TEXT,
    agency_code    TEXT REFERENCES agencies(agency_code),
    lot_no         TEXT,
    po_date        TEXT,  -- ISO date string (YYYY-MM-DD)
    signature_date TEXT,

    niche_price NUMERIC,
    promotion   NUMERIC,
    discount    NUMERIC,
    -- Recomputed on every import as niche_price - promotion - discount,
    -- never trusted verbatim from the sheet (a derived value could be
    -- stale from a hand-edited cell).
    net_price   NUMERIC,

    -- 'pre_need' or 'at_need'. Detected from the Remarks column - see
    -- app/parsing.py for the (best-effort, keyword-based) detection
    -- logic and its documented failure direction.
    case_type      TEXT CHECK (case_type IN ('pre_need', 'at_need')) NOT NULL DEFAULT 'pre_need',
    -- Required for at_need contracts before full-payment commission can
    -- be released; extracted from Remarks text when present.
    inurnment_date TEXT,

    -- 'active' is the normal state. 'on_hold' is NOT auto-detected from
    -- the import (Kenjin's export has no "documents complete" signal) -
    -- it can only be set manually within this tool. 'cancelled' and
    -- 'withdrawn' ARE detected from Remarks keywords, and are
    -- re-derived fresh on every import (not a one-way lock) - so if a
    -- PO's remarks no longer say "cancelled", it naturally reverts to
    -- active on the next upload. Every detected status CHANGE gets
    -- surfaced in the import review panel for a human to confirm.
    status TEXT CHECK (status IN ('active', 'on_hold', 'cancelled', 'withdrawn')) NOT NULL DEFAULT 'active',

    full_settlement_paid_date    TEXT,
    first_installment_paid_date  TEXT,
    sixth_installment_paid_date  TEXT,

    -- Has this trigger already been raised as a commission_event, ever?
    -- Checking this flag before raising a new event is what makes "if
    -- nothing new was crossed, nothing happens" literal rather than
    -- something we have to re-derive by diffing old files.
    full_commission_flagged          INTEGER NOT NULL DEFAULT 0 CHECK (full_commission_flagged IN (0, 1)),
    installment_1_commission_flagged INTEGER NOT NULL DEFAULT 0 CHECK (installment_1_commission_flagged IN (0, 1)),
    installment_6_commission_flagged INTEGER NOT NULL DEFAULT 0 CHECK (installment_6_commission_flagged IN (0, 1)),

    remarks    TEXT,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS commission_runs (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    run_date         TEXT NOT NULL,
    source_filename  TEXT,
    created_by_user  TEXT
);

-- Append-only audit log. One row per commission-due determination,
-- ever. This is what answers "why was this paid" with a fact instead
-- of a reconstruction from old spreadsheets.
CREATE TABLE IF NOT EXISTS commission_events (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    po_no              INTEGER NOT NULL REFERENCES contracts(po_no),
    trigger_type       TEXT NOT NULL CHECK (trigger_type IN ('full_payment', 'installment_1', 'installment_6')),
    trigger_date       TEXT NOT NULL,
    amount             NUMERIC NOT NULL,
    detected_at        TEXT NOT NULL,
    detected_by_user   TEXT,
    commission_run_id  INTEGER REFERENCES commission_runs(id)
);
