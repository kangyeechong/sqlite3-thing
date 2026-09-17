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
    splits_by_agent INTEGER NOT NULL DEFAULT 1 CHECK (splits_by_agent IN (0, 1)),

    -- 'flat': XEKL pays this agency one commission figure and the
    -- agency handles paying their own agent internally - the default,
    -- and what every agency except AW Consultancy currently uses.
    -- 'agency_agent_split': XEKL pays the agency and the agent
    -- separately, at different percentages (see rules.py) - currently
    -- only AW Consultancy's codes. A data flag, seeded from a known
    -- list in rules.py on first sight of a new agency code, never
    -- overwritten afterward - so it's editable data, not hardcoded
    -- logic, the same reasoning as splits_by_agent above.
    commission_split_type TEXT NOT NULL DEFAULT 'flat'
        CHECK (commission_split_type IN ('flat', 'agency_agent_split')),

    -- Several agency_codes can belong to one real-world agency (AW
    -- Consultancy's AC108-01/-02/-03 are all the same agency, just
    -- different internal sub-codes). NULL means this code stands
    -- alone as its own group - the default, and what every agency
    -- except AW Consultancy currently uses. When set, the Excel
    -- export shows one combined sheet for the whole group first
    -- (every code's rows together, like the group's own flat view),
    -- THEN separate per-agent sheets underneath it - matching the
    -- real file's actual structure (an "AW Consultancy" sheet with
    -- everyone, followed by one sheet per named agent), not jumping
    -- straight from individual codes to individual agents with no
    -- combined view in between.
    agency_group TEXT
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

    -- When Accounts actually sent the money - filled in by hand on the
    -- real Master Report, never computed by this tool. Read back from
    -- the sheet on import (like every other date column) so a later
    -- upload, once Accounts has filled these in, carries "confirmed
    -- paid" status into the next report - this tool never writes to
    -- these itself.
    full_commission_paid_date          TEXT,
    installment_1_commission_paid_date TEXT,
    installment_6_commission_paid_date TEXT,

    -- Has this trigger already been raised as a commission_event, ever?
    -- Checking this flag before raising a new event is what makes "if
    -- nothing new was crossed, nothing happens" literal rather than
    -- something we have to re-derive by diffing old files.
    full_commission_flagged          INTEGER NOT NULL DEFAULT 0 CHECK (full_commission_flagged IN (0, 1)),
    installment_1_commission_flagged INTEGER NOT NULL DEFAULT 0 CHECK (installment_1_commission_flagged IN (0, 1)),
    installment_6_commission_flagged INTEGER NOT NULL DEFAULT 0 CHECK (installment_6_commission_flagged IN (0, 1)),

    -- Purely manual - there is no data signal for this anywhere in the
    -- Kenjin export. An agent tells staff verbally that a lead came
    -- from XEKL's own Facebook ads, and staff tick this themselves.
    -- Only meaningful for agencies with commission_split_type =
    -- 'agency_agent_split'; ignored otherwise.
    fb_lead_referred INTEGER NOT NULL DEFAULT 0 CHECK (fb_lead_referred IN (0, 1)),

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
    -- The total commission this trigger released. For a 'flat' agency
    -- this is the whole payout. For an 'agency_agent_split' agency,
    -- this still equals agency_amount + agent_amount (the FB-lead
    -- deduction reduces the total, not just the agency's share) - so
    -- every existing sum/total/report that only reads `amount` keeps
    -- working unchanged for both agency types.
    amount             NUMERIC NOT NULL,
    -- Only populated for 'agency_agent_split' agencies; NULL for
    -- 'flat' ones. agency_amount already has any FB-lead deduction
    -- applied - it is the actual amount payable to the agency, not
    -- the pre-deduction figure.
    agency_amount      NUMERIC,
    agent_amount       NUMERIC,
    detected_at        TEXT NOT NULL,
    detected_by_user   TEXT,
    commission_run_id  INTEGER REFERENCES commission_runs(id),

    -- 'pending': detected automatically, nothing sent to Accounts yet.
    -- 'confirmed': a staff member reviewed it and confirmed the payment
    -- is real - see app/web/routes.py's /review pages. Only confirmed
    -- events show up in a downloaded report or the Date Record summary
    -- - detection alone is never enough to treat something as due.
    -- The *_commission_flagged columns on contracts are still set the
    -- moment an event is detected (pending or not), so a candidate is
    -- never raised twice just because it hasn't been confirmed yet.
    status              TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'confirmed')),
    confirmed_at        TEXT,
    confirmed_by_user   TEXT
);

-- The real Master sheet carries its own trailing "Date Record"
-- summary table (below the PO rows) - the permanent history of every
-- processing cycle that happened before this tool existed. Read once
-- from the uploaded file (app/importer.py) and kept here forever, so
-- that history shows up in the Date Record table (app/report.py)
-- instead of the tool's own tracking silently starting from zero.
-- One row per "As at" date - date_record is UNIQUE so re-uploading
-- the same (or another) file that repeats the same historical rows
-- never duplicates them; the first import of a given date wins.
CREATE TABLE IF NOT EXISTS historical_summary_rows (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    date_record             TEXT NOT NULL UNIQUE,
    full_commission         NUMERIC NOT NULL DEFAULT 0,
    first_half_commission   NUMERIC NOT NULL DEFAULT 0,
    second_half_commission  NUMERIC NOT NULL DEFAULT 0,
    remarks                 TEXT,
    imported_at             TEXT NOT NULL,
    imported_by_user        TEXT,
    source_filename         TEXT
);
