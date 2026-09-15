# Data Model — Phase 1

Status: DRAFT — for review before any code is written.

This document describes the database schema and the state machine that
drives commission calculation. It's based on:

- `Comm. Documentation.pdf` (the written rules/workflow doc)
- A real (anonymized) Master report: `XEKL Commission Payout - 202606`
- A real (anonymized) AOR (Acknowledgment of Receipt) export
- Decisions made in conversation, listed inline below as "Decision:" notes

If anything here doesn't match how the business actually works, say so —
this is the checkpoint before writing a single line of application code.

---

## 1. Input scope

**Decision:** Phase 1 imports only the Kenjin **Master report** (one
Excel file per upload). The AOR file's `Reference No` column is too
inconsistent to parse automatically (~15% of rows are ambiguous —
`BALANCE PAYMENT`, `PARTIAL PAYMENT`, `STAMP DUTY`, `DEPOSIT`, etc. don't
map cleanly to "installment 1" or "installment 6"). Staff keep doing that
cross-check by hand, the same way they do today, and fill in the Master
report's paid-date columns. The tool picks up from there.

This means the entire system's picture of the world is: **"what does the
paid-date columns of the latest Master report say."** Nothing is inferred
from raw transaction text.

---

## 2. Real column headers (Master report)

Row 6 of the Master sheet, confirmed from the real file:

```
No | PO No | PO Date | Signature Date | Customer ID | Customer Name |
Lot No | Niche/Tablet Price (RM) | Promotion (RM) | Discount (RM) |
Nett Price (RM) | Cooling Off Period | Full Settlement Paid Date |
Full Payment Commission (RM) | Full Commission Paid Date |
First Instalment Paid Date | 1st Half Commission (RM) |
1st Half Commission Paid Date | Sixth Instalment Paid Date |
Balance Half Commission (RM) | Balance Half Commission Paid Date |
FCC/Agent | Agency Code | Remarks
```

Notes on specific columns:

- **`PO No`** is the unique contract identifier (confirmed: 59/59 unique
  in the sample file). This is our primary key for a contract.
- **`Customer ID`** identifies the person, and **can repeat** — a
  customer can have multiple POs (multiple lots/niches). So Customer is
  its own entity, one-to-many with Contract.
- **`Cooling Off Period`** is a free-text status staff type by hand
  today (`EXPIRED`, `WITHDRAWAL`, or blank) — not a number. Phase 1
  computes this automatically instead (see §4) and this column becomes
  informational/legacy on import, not something we read as truth.
- **`Full Settlement Paid Date`** vs **`Full Commission Paid Date`** are
  two different dates: when the customer paid vs. when the commission was
  actually released. We track both — the first is the state-machine
  trigger, the second is when it appeared in an actual commission run.
- **`Agency Code`** and **`FCC/Agent`** are reference fields — stored,
  shown in reports, but no split math computed on them in Phase 1 (see
  §7, out of scope).
- The sheet also contains a **summary table appended below the data
  rows** (Date Record / Full Commission / First Half Commission /
  Second Half Commission / Running Total / Remarks) with no blank-row
  separator the importer can rely on structurally — the importer detects
  the end of real data by checking `No` is a positive integer, not by
  row position.

---

## 3. Entities

### `customers`
| column | type | notes |
|---|---|---|
| customer_id | TEXT PK | Kenjin's Customer ID, e.g. `XEKL000858` |
| name | TEXT | |

### `agencies`
| column | type | notes |
|---|---|---|
| agency_code | TEXT PK | e.g. `AC001`, `AC108-02`. Nullable FK from contracts — blank Agency Code in the source file looks like a data-entry gap, not a real "no agency" case (per your read of the sample), so the importer **warns** on a blank rather than silently accepting it |
| name | TEXT | filled in as we learn agency names, not required for Phase 1 math |

### `contracts`
One row per PO — this is the core ledger table.

| column | type | notes |
|---|---|---|
| po_no | INTEGER PK | unique contract key |
| customer_id | TEXT FK → customers | |
| agent_name | TEXT | from `FCC/Agent` — single agent, no splits (Phase 1 scope) |
| agency_code | TEXT FK → agencies, nullable | reference only, no split math |
| lot_no | TEXT | |
| po_date | DATE | |
| signature_date | DATE | |
| niche_price | NUMERIC | |
| promotion | NUMERIC | |
| discount | NUMERIC | |
| net_price | NUMERIC | = niche_price − promotion − discount, computed on import, not trusted from the sheet (guards against a stale/hand-edited cell) |
| payment_type | TEXT | `full_payment` \| `installment` |
| case_type | TEXT | `pre_need` \| `at_need` — decision: Phase 1 supports both |
| inurnment_date | DATE, nullable | required when case_type = at_need |
| status | TEXT | `active` \| `on_hold` \| `cancelled` \| `withdrawn` \| `complete` |
| full_settlement_paid_date | DATE, nullable | |
| first_installment_paid_date | DATE, nullable | |
| sixth_installment_paid_date | DATE, nullable | |
| full_commission_flagged | BOOLEAN | has the 15% already been raised as due |
| installment_1_commission_flagged | BOOLEAN | has the first 7.5% already been raised |
| installment_6_commission_flagged | BOOLEAN | has the second 7.5% already been raised |

The three `*_flagged` booleans are the actual "memory" that makes this
tool worth building — see §5.

### `commission_events` (the audit log)
Append-only. One row per commission-due determination, ever.

| column | type | notes |
|---|---|---|
| id | INTEGER PK | |
| po_no | INTEGER FK → contracts | |
| trigger_type | TEXT | `full_payment` \| `installment_1` \| `installment_6` |
| trigger_date | DATE | the paid-date that caused this (e.g. Full Settlement Paid Date) |
| amount | NUMERIC | computed commission amount |
| detected_at | DATETIME | when the tool determined this was due (i.e. which upload/run) |
| detected_by_user | TEXT FK → users | who ran the import that surfaced this |
| commission_run_id | INTEGER FK → commission_runs | which run this shipped in |

This is what answers "why was this paid" — a plain row, not a
reconstruction from diffing spreadsheets.

### `commission_runs`
One row per twice-monthly cycle you actually process.

| column | type | notes |
|---|---|---|
| id | INTEGER PK | |
| run_date | DATE | |
| source_filename | TEXT | which Master report upload produced this |
| created_by_user | TEXT FK → users | |

### `users`
| column | type | notes |
|---|---|---|
| id | INTEGER PK | |
| email | TEXT UNIQUE | |
| password_hash | TEXT | |
| display_name | TEXT | |

Individual logins, per your original brief — no shared password.

---

## 4. State machine

Per contract, independent of everything else:

```
                 ┌─────────────────────────┐
                 │        ACTIVE           │
                 └─────────────────────────┘
                     │                 │
        payment_type=full_payment   payment_type=installment
                     │                 │
                     ▼                 ▼
        ┌─────────────────────┐   ┌─────────────────────────┐
        │ Full Settlement paid │   │ Not Started              │
        │ (date recorded)      │   └─────────────────────────┘
        └─────────────────────┘             │ 1st installment paid
                     │                       ▼
     Pre-Need: wait for cooling-off  ┌─────────────────────────┐
     gate (§4a). At-Need: no wait,   │ Installment 1 Paid       │
     inurnment_date required.        │ → 7.5% due               │
                     │               └─────────────────────────┘
                     ▼                       │ 6th installment paid
        ┌─────────────────────┐              ▼
        │ Full Commission Due  │   ┌─────────────────────────┐
        │ → 15% due             │   │ Installment 6 Paid       │
        └─────────────────────┘   │ → 7.5% due               │
                     │             └─────────────────────────┘
                     ▼                       │
        ┌─────────────────────┐              ▼
        │      COMPLETE        │◄──────────────
        └─────────────────────┘
```

At any point, a contract can instead move to `on_hold` (documents
incomplete — blocks any commission from being flagged, regardless of
payment progress), `cancelled`, or `withdrawn` (per your description:
withdrawn = customer switching product/lot, which shows up as a *new*
PO; cancelled = customer backing out entirely). Both `cancelled` and
`withdrawn` stop that PO from generating further commission events —
Phase 1 does not attempt to auto-link a withdrawn PO to its replacement
PO; the Remarks text already documents that for a human reader.

### 4a. Cooling-off gate (full payment, Pre-Need only)

**Decision:** commission becomes releasable **5 days after
`Full Settlement Paid Date`**. At-Need contracts skip this entirely but
require `inurnment_date` to be set.

This replaces the manual "look at the date, count days, type EXPIRED"
step entirely — the tool computes it from the date column on every
import.

---

## 5. The core import logic (what actually happens on upload)

For every PO in the uploaded Master report:

1. Look up the existing `contracts` row by `po_no` (or create it, if this
   is the first time this PO has appeared).
2. Update the plain fields (dates, prices, status, agent, agency) from
   the sheet.
3. Check each trigger, only flag if **not already flagged**:
   - `full_settlement_paid_date` is set, `case_type` allows it (Pre-Need
     past the 5-day gate, or At-Need with an inurnment date), and
     `full_commission_flagged` is still false → raise a `full_payment`
     event, set the flag true.
   - `first_installment_paid_date` is set and
     `installment_1_commission_flagged` is false → raise an
     `installment_1` event, set the flag true.
   - `sixth_installment_paid_date` is set and
     `installment_6_commission_flagged` is false → raise an
     `installment_6` event, set the flag true.
   - `status` is `on_hold` → skip all of the above regardless.
4. Everything raised in this pass belongs to one new `commission_run`.

If a PO shows up in an upload with a paid-date already flagged from a
previous run, nothing happens — this is the "silently carries forward"
behavior from your brief, now literal: the flag being `true` **is** the
carry-forward.

---

## 6. Rules engine (one place, not scattered)

A single module (`rules.py`) holds every number that could change:

```python
FULL_PAYMENT_COMMISSION_PCT = 0.15
INSTALLMENT_1_COMMISSION_PCT = 0.075
INSTALLMENT_6_COMMISSION_PCT = 0.075
COOLING_OFF_DAYS_BEFORE_RELEASE = 5
COOLING_OFF_TOTAL_DAYS = 10  # documented, not currently used in a calc
```

Every other part of the system calls into this module rather than
hardcoding a percentage — when the agency-split rules get built later,
they're new entries here, not new `if` statements buried in importer
code.

---

## 7. Explicitly out of scope for Phase 1 (schema left open, no logic built)

- Agency-split commission math (AW Consultancy's 7%/8% and 3.5%/4%
  split, FB-lead 3%/1.5% deductions from the agency's portion). The real
  numbers, confirmed from the doc, differ from the brief's original
  guess of 7.5%/7.5% — noted here so nobody builds against the wrong
  figures later.
- KPI bonus, agency development fund, agent incentive.
- Auto-importing/parsing the AOR file.
- Reversible name encryption (mentioned as a feature suggestion in the
  doc) — real names are stored as-is in the database; anonymization was
  only needed for getting sample data into this conversation safely.

---

## 8. Open items I'm assuming reasonable defaults for — flag if wrong

- **Non-AW agencies with blank Agency Code**: treated as a data-quality
  warning on import, not silently accepted.
- **`net_price` is recomputed on import**, not trusted verbatim from the
  sheet, since it's a derived value and the sheet could have a stale
  manually-typed number.
- **The 5th–8th / 15th–20th windows are not enforced** — a
  `commission_run` can be created any day; the dates are for your own
  reference, not a validation rule (per your "loose, depends on
  workload" note).
