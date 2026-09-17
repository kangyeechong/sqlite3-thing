"""
Commission rules for Phase 1 - the one place every percentage, trigger
point, and timing rule lives.

Nothing outside this file should hardcode a percentage or a day count.
When agency-split math, FB-lead deductions, or KPI bonuses get built in
a later phase, they're new entries here, not new `if` statements
scattered through importer or commission code.
"""

# --- Percentages -----------------------------------------------------
# Full payment: 15% released in one go.
FULL_PAYMENT_COMMISSION_PCT = 0.15

# Installment plans: split into two equal 7.5% halves, always triggered
# at installment #1 and installment #6 - confirmed against real
# Reference No. codes (INST 01/24, 06/24, 01/06, 06/06, 01/12, 06/12,
# 01/18, 06/18): the trigger point never moves even though plans run
# 6, 12, 18, or 24 months.
INSTALLMENT_1_COMMISSION_PCT = 0.075
INSTALLMENT_6_COMMISSION_PCT = 0.075
INSTALLMENT_TRIGGER_1 = 1
INSTALLMENT_TRIGGER_6 = 6

# --- Cooling-off gate (Pre-Need full payment only) --------------------
# Full-payment commission becomes releasable this many days after
# Full Settlement Paid Date. Confirmed with the business: this is
# deliberately less than the full cooling-off period - Accounts
# department processing is slow, so the commission team pre-flags at
# day 5 to give Accounts lead time; by the time they actually process
# it, the full 10 days will have elapsed anyway.
COOLING_OFF_DAYS_BEFORE_RELEASE = 5

# Documented for reference (this is the legal/contractual cooling-off
# window); not used directly in the release calculation above.
COOLING_OFF_TOTAL_DAYS = 10

# --- Export grouping exceptions ---------------------------------------
# Agencies that do NOT get a per-agent breakdown in the Excel export -
# see docs/data_model.md section 6a. AC001 (XEMP) is in-house sales
# staff, not an external agency, so there's no agency-vs-agent split to
# show. This only seeds a brand-new agency's `splits_by_agent` flag the
# first time it's ever seen; it is never used to overwrite an existing
# agency row, since that flag is meant to be data staff can maintain
# going forward (e.g. if another agency needs the same exception
# later), not something re-derived from the Excel file every import.
AGENCIES_WITHOUT_PER_AGENT_SPLIT = {"AC001"}

# --- AW Consultancy agency/agent commission split ----------------------
# Confirmed directly, and cross-checked against real rows: e.g. PO
# 20260266 (AC108-02, Net Price RM18,300) produces 3.5% = RM640.50 and
# 4% = RM732.00, exactly matching the real split table shown for that
# row. Every other agency stays 'flat' (full commission paid to the
# agency, who distributes to their own agents themselves) - AW
# Consultancy is the sole confirmed exception.
AW_AGENCY_FULL_PAYMENT_PCT = 0.07
AW_AGENT_FULL_PAYMENT_PCT = 0.08
AW_AGENCY_INSTALLMENT_PCT = 0.035
AW_AGENT_INSTALLMENT_PCT = 0.04

# Deducted from the AGENCY's share only (never the agent's) when a
# sale is manually flagged as FB-lead-referred (contracts.fb_lead_referred)
# - purely a manual entry, there is no data signal for this anywhere in
# the Kenjin export (an agent has to tell staff verbally).
AW_FB_LEAD_DEDUCTION_FULL_PAYMENT_PCT = 0.03
AW_FB_LEAD_DEDUCTION_INSTALLMENT_PCT = 0.015

# Agency codes confirmed (from the real sample file's "AW Consultancy"
# sheet) to use the agency/agent split above, seeded onto a brand-new
# agency's commission_split_type the first time it's seen - same
# never-overwrite-after-first-sight rule as
# AGENCIES_WITHOUT_PER_AGENT_SPLIT. AW Consultancy's individual agents
# each carry their own sub-code (AC108-01, -02, -03 confirmed so far);
# if another agent joins with a new code, add it here - a data change,
# not a code change. Flag to the user if a new AW Consultancy code
# shows up that isn't in this list yet, rather than silently treating
# it as a flat-commission agency.
AGENCIES_WITH_AGENCY_AGENT_SPLIT = {"AC108-01", "AC108-02", "AC108-03"}

# Which agency_codes are really sub-codes of one shared real-world
# agency - maps agency_code -> the group name the Excel export should
# show a combined sheet under (see agencies.agency_group in
# schema.sql). AW Consultancy's three sub-codes all point at the same
# "AW Consultancy" group; a code not listed here stands alone as its
# own group (the default for every other agency). Same
# never-overwrite-after-first-sight seeding rule as the flags above.
AGENCY_GROUPS = {
    "AC108-01": "AW Consultancy",
    "AC108-02": "AW Consultancy",
    "AC108-03": "AW Consultancy",
}
