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

    -- Manual for a routine Kenjin export - there's no data signal for
    -- this anywhere in it, an agent tells staff verbally that a lead
    -- came from XEKL's own Facebook ads, and staff tick this
    -- themselves. Auto-detected instead when onboarding a hand-
    -- maintained pre-existing file that already has its own AW
    -- Consultancy-style referral-fee breakdown sheet - see
    -- app.importer._read_referral_flagged_po_nos - sticky once set,
    -- never un-set by a later routine upload that doesn't carry that
    -- sheet. Only meaningful for agencies with commission_split_type =
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
    -- The total commission this trigger released - always the flat
    -- company-wide percentage of Net Price (see commission._build_
    -- event), for BOTH a 'flat' agency and an 'agency_agent_split' one
    -- (AW Consultancy). The FB-lead deduction never touches this
    -- figure - it's an internal AW Consultancy bookkeeping detail
    -- (how this SAME total divides between agency and agent), applied
    -- only to agency_amount below. So for an FB-lead-referred AW
    -- Consultancy sale, agency_amount + agent_amount is LESS than
    -- `amount` by exactly the deduction - that's expected, not a bug.
    amount             NUMERIC NOT NULL,
    -- Only populated for 'agency_agent_split' agencies; NULL for
    -- 'flat' ones. agency_amount already has any FB-lead deduction
    -- applied - it is the actual amount payable to the agency, not
    -- the pre-deduction figure. See the comment on `amount` above for
    -- why agency_amount + agent_amount can be less than `amount`.
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
    -- 'voided': a confirmed event that turned out to be wrong (bad
    -- price, mismatched receipt, confirmed by mistake) - see
    -- app.commission.void_commission_event. Never deleted, so the
    -- mistake and who corrected it stay visible forever (same "nothing
    -- hidden" standing-ledger philosophy as the rest of this app); just
    -- excluded from every report the same way a still-pending event
    -- already is.
    -- The *_commission_flagged columns on contracts are still set the
    -- moment an event is detected (pending or not), so a candidate is
    -- never raised twice just because it hasn't been confirmed yet -
    -- voiding an event clears its own flag back off so the corrected
    -- figure can be detected fresh.
    status              TEXT NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'confirmed', 'voided')),
    confirmed_at        TEXT,
    confirmed_by_user   TEXT,
    voided_at           TEXT,
    voided_by_user      TEXT,
    void_reason         TEXT
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

-- One row per AOR (Acknowledgment of Receipt) receipt ever imported -
-- see app/aor.py. acknowledgment_receipt_no is UNIQUE so re-uploading
-- an AOR export that overlaps a previous one (the real exports
-- routinely do - see app/aor.py's module docstring) never re-applies
-- the same receipt twice; the first import of a given receipt wins.
-- po_no is nullable: a receipt whose PO doesn't exist in the ledger
-- yet is still recorded as seen (so it's never silently reprocessed
-- once the PO does show up) but has nothing to attach a paid-date to.
-- aor_upload_id is nullable too (and NULL for every row written before
-- this column existed): purely a record of which upload first
-- introduced each receipt, not consulted by anything to decide what
-- gets shown - annotate_aor_file reads the uploaded file directly, so
-- a receipt already recorded by an earlier overlapping upload still
-- shows up in a later file's own annotated copy, exactly as it
-- genuinely appears in that file.
-- receipt_date/reference_text/payment_received/trigger_type (added
-- after the columns above, all nullable - NULL for every row written
-- before this existed) persist enough about each receipt to record
-- what it actually was, independent of whether it ended up changing
-- anything on the contract. trigger_type is NULL for a receipt
-- that was recognized but non-triggering (a Deposit, Stamp Duty, or
-- an installment number that isn't 1 or 6) or whose PO wasn't in the
-- ledger; otherwise it's a comma-joined list of whichever of
-- full_payment/installment_1/installment_6 this receipt's own
-- Reference No classified as - a receipt "being" that kind of payment
-- is a fact about the receipt itself, independent of whether its
-- specific write ended up changing anything on the contract (an
-- already-filled paid-date isn't overwritten, but the receipt still
-- genuinely was, say, an installment 1 payment).
CREATE TABLE IF NOT EXISTS aor_receipts (
    id                         INTEGER PRIMARY KEY AUTOINCREMENT,
    acknowledgment_receipt_no  TEXT NOT NULL UNIQUE,
    po_no                      INTEGER,
    imported_at                TEXT NOT NULL,
    imported_by_user           TEXT,
    source_filename            TEXT,
    aor_upload_id              INTEGER REFERENCES aor_uploads(id),
    receipt_date               TEXT,
    reference_text             TEXT,
    payment_received           NUMERIC,
    trigger_type               TEXT
);

-- The raw bytes of an uploaded AOR export, kept only so its
-- "downloadable annotated copy" (see app/aor.py's annotate_aor_file)
-- can be regenerated on demand from the /download-aor-annotated/<id>
-- route - the same "generate the download fresh from persisted state
-- every time" approach /download/<run_id> already uses for the
-- commission report, rather than caching generated bytes in memory,
-- which wouldn't survive a restart or a second worker process.
-- One row per AOR upload. Deliberately keyed by its own id, not
-- commission_run_id - process_commission_run only creates a
-- commission_runs row when something was actually newly detected
-- (see app/commission.py), but an AOR upload always has a file worth
-- annotating and downloading, whether or not anything happened to be
-- newly due that run. commission_run_id is carried along only when
-- one exists, so the results page can still link this upload back to
-- its review/confirm flow.
CREATE TABLE IF NOT EXISTS aor_uploads (
    id                 INTEGER PRIMARY KEY AUTOINCREMENT,
    commission_run_id  INTEGER REFERENCES commission_runs(id),
    filename           TEXT NOT NULL,
    file_bytes         BLOB NOT NULL,
    uploaded_at        TEXT NOT NULL,
    -- The period staff chose when uploading (see process_aor_upload) -
    -- NULL only for an upload that predates period filtering ever
    -- existing. annotate_aor_file uses NULL-ness here (not "zero
    -- receipts ended up linked to this upload") to tell "this upload
    -- never went through period-aware code at all, so an unscoped
    -- Filtered sheet is the right fallback" apart from "this upload
    -- genuinely matched nothing in its own chosen period, so an empty
    -- Filtered sheet is the CORRECT answer" - those look identical by
    -- receipt count alone but need opposite behavior.
    period_start       TEXT,
    period_end         TEXT
);

-- Marks which multi-statement migrations (see app/db/connection.py)
-- have fully completed, including any one-time data backfill - not
-- just which columns/tables exist. A column can exist the instant its
-- ALTER TABLE statement runs (SQLite commits DDL immediately, with no
-- way to roll it back), but the backfill statement that has to follow
-- it is separate DML that only becomes durable on a later commit. This
-- table is what lets a crash between the two be detected and safely
-- resumed on the next connection, instead of the backfill being
-- silently skipped forever just because the column already exists.
CREATE TABLE IF NOT EXISTS schema_migrations (
    key        TEXT PRIMARY KEY,
    applied_at TEXT NOT NULL
);
