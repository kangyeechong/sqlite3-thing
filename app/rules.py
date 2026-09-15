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
