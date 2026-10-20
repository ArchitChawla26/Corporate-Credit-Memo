#!/usr/bin/env python3
"""
pull_data.py — SEC EDGAR data pull for the CVS Health corporate credit memo.

Pulls CVS Health (CIK 0000064803) financials from the SEC EDGAR XBRL
company-facts API and the submissions index, extracts the exact series used
in CVS_credit_memo.md / CVS_credit_model.xlsx, and writes:

    data/companyfacts_<pull-date>.json    raw company-facts response
    data/submissions_<pull-date>.json     raw submissions response
    data/inputs.json                      extracted series + documented adjustments
    data/inputs.csv                       same, flat, in $ millions

Selection rules (applied uniformly, values in USD as filed):
  * Annual flows:  latest-filed 10-K entry with a ~365-day duration ending
    Dec 31 (fp=FY). Interim flows: latest-filed 10-Q entry with a ~181-day
    duration ending Jun 30 (six months year-to-date).
  * Instants:      latest-filed 10-K / 10-Q entry at the period-end date.
  * Interest expense uses us-gaap:InterestExpenseNonoperating (the tag CVS
    applies from FY2022 on); FY2021 falls back to us-gaap:InterestExpense.
    The two tags carry identical values in every overlap year.
  * Figures are therefore "as most recently reported" in comparative
    filings. Immaterial reclassifications vs. originally filed amounts
    (e.g. FY2022 operating income) are not adjusted.

Non-recurring adjustments are NOT consistently XBRL-tagged, so they are
carried in ADJUSTMENTS below, each tied to the printed statement or
reconciliation it was taken from. Goodwill impairment IS tagged and is
pulled from XBRL like everything else.

SEC fair-access: a descriptive User-Agent with contact is sent, per
https://www.sec.gov/os/accessing-edgar-data

Author: Archit Chawla — independent research project.
"""

import csv
import json
import sys
import time
import urllib.request
from datetime import date, datetime
from pathlib import Path

CIK = "0000064803"
UA = "ArchitChawla Research archit.26.chawla@gmail.com"
BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
PULL_DATE = date.today().isoformat()

FLOW_TAGS = {  # label -> us-gaap tag (annual + six-month durations)
    "revenues": "Revenues",
    "operating_income": "OperatingIncomeLoss",
    "depreciation_amortization": "DepreciationDepletionAndAmortization",
    "net_income": "NetIncomeLoss",
    "operating_cash_flow": "NetCashProvidedByUsedInOperatingActivities",
    "capex": "PaymentsToAcquirePropertyPlantAndEquipment",
    "dividends_paid": "PaymentsOfDividends",
    "diluted_shares": "WeightedAverageNumberOfDilutedSharesOutstanding",
    "goodwill_impairment": "GoodwillImpairmentLoss",
}
INTEREST_TAGS = ["InterestExpenseNonoperating", "InterestExpense"]  # splice order
INSTANT_TAGS = {  # label -> us-gaap tag (balance-sheet dates)
    "lt_debt_noncurrent": "LongTermDebtAndCapitalLeaseObligations",
    "lt_debt_current": "LongTermDebtAndCapitalLeaseObligationsCurrent",
    "short_term_borrowings": "ShortTermBorrowings",
    "cash": "CashAndCashEquivalentsAtCarryingValue",
    "op_lease_current": "OperatingLeaseLiabilityCurrent",
    "op_lease_noncurrent": "OperatingLeaseLiabilityNoncurrent",
    "fin_lease_current": "FinanceLeaseLiabilityCurrent",
    "fin_lease_noncurrent": "FinanceLeaseLiabilityNoncurrent",
    "total_assets": "Assets",
    "total_equity": "StockholdersEquity",
    "shares_outstanding": "CommonStockSharesOutstanding",
}
MATURITY_TAGS = {  # debt principal ladder, instant at Dec 31, 2025, FY2025 10-K
    "due_2026": "LongTermDebtMaturitiesRepaymentsOfPrincipalInNextTwelveMonths",
    "due_2027": "LongTermDebtMaturitiesRepaymentsOfPrincipalInYearTwo",
    "due_2028": "LongTermDebtMaturitiesRepaymentsOfPrincipalInYearThree",
    "due_2029": "LongTermDebtMaturitiesRepaymentsOfPrincipalInYearFour",
    "due_2030": "LongTermDebtMaturitiesRepaymentsOfPrincipalInYearFive",
    "due_thereafter": "LongTermDebtMaturitiesRepaymentsOfPrincipalAfterYearFive",
}

FY_YEARS = [2021, 2022, 2023, 2024, 2025]
BS_DATES = [f"{y}-12-31" for y in FY_YEARS] + ["2026-06-30"]

# Non-recurring items added back to reported EBITDA, $ millions.
# Each is a separate line item in the consolidated statements of operations
# or a discrete charge quantified in the filing's segment reconciliation /
# MD&A. Sources:
#   FY2021 goodwill impairment 431 — XBRL GoodwillImpairmentLoss (also pulled above)
#   FY2022 opioid litigation charges 5,803; loss on assets held for sale 2,533
#          (Omnicare LTC write-down) — FY2022 10-K consolidated results table
#   FY2023 restructuring 507; loss on assets held for sale 349
#          — FY2025 10-K consolidated results table (comparative columns)
#   FY2024 restructuring 1,179 — FY2025 10-K consolidated results table;
#          opioid accrual charge 100 — Q2 2025 10-Q segment reconciliation
#          (six months ended Jun 30, 2024; unchanged at nine months)
#   FY2025 goodwill impairment 5,725 (Q3 2025, Health Care Delivery reporting
#          unit) — XBRL; legacy litigation charges 1,220 (recorded H1 2025 —
#          Q2 2025 10-Q reconciliation) and opioid accrual charge 320
#          (Q3 2025) — FY2025 10-K MD&A and segment reconciliation
#   LTM to Jun 30, 2026: only the Q3 2025 items fall in the window
#          (5,725 goodwill + 320 opioid). The 1,220 legacy litigation charges
#          were recorded in H1 2025 and are swapped out of the LTM period.
ADJUSTMENTS = {
    "FY2021": {"goodwill_impairment": 431, "opioid_litigation": 0, "legacy_litigation": 0,
               "loss_assets_held_for_sale": 0, "restructuring": 0},
    "FY2022": {"goodwill_impairment": 0, "opioid_litigation": 5803, "legacy_litigation": 0,
               "loss_assets_held_for_sale": 2533, "restructuring": 0},
    "FY2023": {"goodwill_impairment": 0, "opioid_litigation": 0, "legacy_litigation": 0,
               "loss_assets_held_for_sale": 349, "restructuring": 507},
    "FY2024": {"goodwill_impairment": 0, "opioid_litigation": 100, "legacy_litigation": 0,
               "loss_assets_held_for_sale": 0, "restructuring": 1179},
    "FY2025": {"goodwill_impairment": 5725, "opioid_litigation": 320, "legacy_litigation": 1220,
               "loss_assets_held_for_sale": 0, "restructuring": 0},
    "LTM":    {"goodwill_impairment": 5725, "opioid_litigation": 320, "legacy_litigation": 0,
               "loss_assets_held_for_sale": 0, "restructuring": 0},
}


def fetch(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": UA,
                                                 "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as resp:
        return json.loads(resp.read().decode("utf-8"))


def d(s):
    return datetime.strptime(s, "%Y-%m-%d").date()


def pick_flow(entries, end, lo, hi, forms):
    """Latest-filed flow entry ending `end` with duration in [lo, hi] days."""
    cands = []
    for r in entries:
        if r.get("form") not in forms or r["end"] != end or "start" not in r:
            continue
        dur = (d(r["end"]) - d(r["start"])).days
        if lo <= dur <= hi:
            cands.append(r)
    if not cands:
        return None
    return max(cands, key=lambda r: r["filed"])


def pick_instant(entries, end, forms=("10-K", "10-Q")):
    cands = [r for r in entries
             if r.get("form") in forms and r["end"] == end and "start" not in r]
    if not cands:
        return None
    return max(cands, key=lambda r: r["filed"])


def main():
    DATA.mkdir(parents=True, exist_ok=True)
    facts = fetch(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json")
    (DATA / f"companyfacts_{PULL_DATE}.json").write_text(json.dumps(facts))
    time.sleep(0.2)
    subs = fetch(f"https://data.sec.gov/submissions/CIK{CIK}.json")
    (DATA / f"submissions_{PULL_DATE}.json").write_text(json.dumps(subs))

    gaap = facts["facts"]["us-gaap"]
    out = {"issuer": "CVS Health Corporation", "ticker": "CVS", "cik": CIK,
           "pull_date": PULL_DATE, "units": "USD as filed",
           "flows_annual": {}, "flows_h1": {}, "instants": {}, "maturity_ladder": {},
           "adjustments_musd": ADJUSTMENTS}

    def rows(tag, unit="USD"):
        return gaap.get(tag, {}).get("units", {}).get(unit, [])

    for label, tag in FLOW_TAGS.items():
        unit = "shares" if label == "diluted_shares" else "USD"
        out["flows_annual"][label] = {}
        out["flows_h1"][label] = {}
        for y in FY_YEARS:
            r = pick_flow(rows(tag, unit), f"{y}-12-31", 350, 380, ("10-K",))
            out["flows_annual"][label][f"FY{y}"] = r["val"] if r else None
        for y in (2025, 2026):
            r = pick_flow(rows(tag, unit), f"{y}-06-30", 170, 200, ("10-Q",))
            out["flows_h1"][label][f"H1-{y}"] = r["val"] if r else None

    # Interest expense: splice InterestExpenseNonoperating -> InterestExpense
    out["flows_annual"]["interest_expense"] = {}
    out["flows_h1"]["interest_expense"] = {}
    for y in FY_YEARS:
        val = None
        for tag in INTEREST_TAGS:
            r = pick_flow(rows(tag), f"{y}-12-31", 350, 380, ("10-K",))
            if r:
                val = r["val"]; break
        out["flows_annual"]["interest_expense"][f"FY{y}"] = val
    for y in (2025, 2026):
        val = None
        for tag in INTEREST_TAGS:
            r = pick_flow(rows(tag), f"{y}-06-30", 170, 200, ("10-Q",))
            if r:
                val = r["val"]; break
        out["flows_h1"]["interest_expense"][f"H1-{y}"] = val

    for label, tag in INSTANT_TAGS.items():
        unit = "shares" if label == "shares_outstanding" else "USD"
        out["instants"][label] = {}
        for end in BS_DATES:
            r = pick_instant(rows(tag, unit), end)
            out["instants"][label][end] = r["val"] if r else None
    # Goodwill impairment is only tagged when a charge exists; the six-month
    # periods had none (the FY2025 charge was recorded entirely in Q3 2025,
    # per the Q3 2025 10-Q). Untagged six-month values are nil.
    for per in ("H1-2025", "H1-2026"):
        if out["flows_h1"]["goodwill_impairment"][per] is None:
            out["flows_h1"]["goodwill_impairment"][per] = 0
    # CVS presented no short-term borrowings at Dec 31, 2021 (untagged = nil)
    if out["instants"]["short_term_borrowings"]["2021-12-31"] is None:
        out["instants"]["short_term_borrowings"]["2021-12-31"] = 0
    # None outstanding at Jun 30, 2026 — stated in the Q2 2026 10-Q
    if out["instants"]["short_term_borrowings"]["2026-06-30"] is None:
        out["instants"]["short_term_borrowings"]["2026-06-30"] = 0

    for label, tag in MATURITY_TAGS.items():
        r = pick_instant(rows(tag), "2025-12-31", forms=("10-K",))
        out["maturity_ladder"][label] = r["val"] if r else None

    (DATA / "inputs.json").write_text(json.dumps(out, indent=2))

    with open(DATA / "inputs.csv", "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["series", "period", "value_musd_or_shares_m"])
        for sect in ("flows_annual", "flows_h1"):
            for label, vals in out[sect].items():
                for per, v in vals.items():
                    w.writerow([f"{sect}:{label}", per,
                                None if v is None else v / 1e6])
        for label, vals in out["instants"].items():
            for per, v in vals.items():
                w.writerow([f"instants:{label}", per,
                            None if v is None else v / 1e6])
        for label, v in out["maturity_ladder"].items():
            w.writerow([f"maturity:{label}", "2025-12-31",
                        None if v is None else v / 1e6])

    missing = []
    for sect in ("flows_annual", "flows_h1", "instants"):
        for label, vals in out[sect].items():
            for per, v in vals.items():
                if v is None and not (sect == "instants" and
                                      label.startswith("fin_lease") and
                                      per == "2026-06-30"):
                    missing.append(f"{sect}:{label}:{per}")
    for label, v in out["maturity_ladder"].items():
        if v is None:
            missing.append(f"maturity:{label}")
    if missing:
        print("MISSING VALUES:", *missing, sep="\n  ")
        return 1
    print(f"Pull complete ({PULL_DATE}). Wrote data/inputs.json and data/inputs.csv")
    print("Total debt check ($M): " +
          ", ".join(f"{e}: {(out['instants']['lt_debt_noncurrent'][e] + out['instants']['lt_debt_current'][e] + out['instants']['short_term_borrowings'][e]) / 1e6:,.0f}"
                    for e in BS_DATES))
    return 0


if __name__ == "__main__":
    sys.exit(main())
