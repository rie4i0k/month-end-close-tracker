"""
generate_data.py
----------------
Creates the fictional sample data for the Month-End Close Tracker.

Company: Meridian Card Services (fictional).
All task names, owners, dates, balances and risks are FICTIONAL.

The only real-world input is public STRUCTURE from American Express's
FY2025 Form 10-K (filed 2026-02-06, SEC EDGAR):
  * reportable operating segment names, and the "Corporate & Other"
    category (which the 10-K says is NOT a reportable segment)
  * revenue line-item names (MD&A Table 2, "Total revenues net of
    interest expense summary")
See the "About & Sources" sheet / README for the citation.

Outputs:
  data/close_tasks.csv
  data/risk_log.csv
  data/reference_structure.csv
"""

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"

# ---------------------------------------------------------------------------
# Verified public structure (FY2025 10-K). Nothing here is a number.
# ---------------------------------------------------------------------------
SEGMENTS = [
    "U.S. Consumer Services",
    "Commercial Services",
    "International Card Services",
    "Global Merchant and Network Services",
]
CORPORATE = "Corporate & Other"  # not a reportable segment per the 10-K

REVENUE_LINES = [
    "Discount revenue",
    "Net card fees",
    "Service fees and other revenue",
    "Total interest income",
    "Total interest expense",
    "Net interest income",
]

USCS, CS, ICS, GMNS = SEGMENTS

# ---------------------------------------------------------------------------
# Fictional people
# ---------------------------------------------------------------------------
PREP = {
    "AP": "Ava Patel",
    "MR": "Marcus Reyes",
    "LC": "Lena Chen",
    "DO": "Daniel Okafor",
    "SM": "Sofia Martins",
    "JW": "James Whitaker",
    "NK": "Nora Kim",
}
REV = {
    "RM": "Rahul Mehta (Rev. Accounting Mgr)",
    "EB": "Elena Brooks (Balance Sheet Mgr)",
    "TH": "Tom Hughes (Asst. Controller)",
    "CT": "Claire Thompson (Controller)",
}

# ---------------------------------------------------------------------------
# 25 close tasks (status shown as of WD+2 of the fictional November close).
# Workday convention: WD0 = last business day of the month.
# Columns: id, name, segment, category, owner, reviewer, due, actual,
#          status, priority, dependency, control, comments
# ---------------------------------------------------------------------------
TASKS = [
    ("T01", "Subledger feed cutoff & completeness check (Oracle on-prem to Cloud bridge)", CORPORATE, "Systems & Data", "NK", "TH", -2, -2, "Complete", "High", "", "Preventive", "All 14 feeds received; 2 re-sent after cutover window."),
    ("T02", "Card Member billing cycle cutoff confirmation", USCS, "Revenue", "AP", "RM", -2, -2, "Complete", "High", "T01", "Preventive", ""),
    ("T03", "Net card fees revenue recognition – annual fee deferral & amortisation", USCS, "Revenue", "AP", "RM", 1, 1, "Complete", "High", "T02", "Detective", "Deferred fee rollforward tied to subledger."),
    ("T04", "Net card fees revenue recognition – corporate & small business cards", CS, "Revenue", "MR", "RM", 1, 1, "Complete", "High", "T02", "Detective", ""),
    ("T05", "Net card fees revenue recognition – international portfolios (multi-currency)", ICS, "Revenue", "LC", "RM", 1, None, "Delayed", "High", "T01", "Detective", "Month-end FX rate table not yet loaded into Cloud GL; revised ETA WD+3."),
    ("T06", "Discount revenue accrual – merchant settlement cutoff", GMNS, "Revenue", "DO", "RM", 1, 1, "Complete", "High", "T01", "Detective", ""),
    ("T07", "Network partner / interchange settlement reconciliation", GMNS, "Reconciliation", "DO", "EB", 2, None, "In Progress", "High", "T06", "Detective", "3 of 5 partner statements reconciled."),
    ("T08", "Card Member rewards liability accrual – consumer programme", USCS, "Accruals", "SM", "TH", 1, 2, "Complete", "High", "T02", "Detective", "Completed 1 day late; redemption-rate assumption refreshed."),
    ("T09", "Card Member rewards liability accrual – international programmes", ICS, "Accruals", "SM", "TH", 2, None, "In Progress", "High", "T05", "Detective", ""),
    ("T10", "Credit loss reserve true-up – consumer card receivables & loans", USCS, "Reserves", "JW", "TH", 2, None, "In Progress", "High", "T02", "Detective", "Awaiting final model sign-off from Credit Risk."),
    ("T11", "Credit loss reserve true-up – commercial card receivables", CS, "Reserves", "JW", "TH", 1, None, "Delayed", "High", "T04", "Detective", "Model run re-performed after data mapping issue in Cloud."),
    ("T12", "Service fees and other revenue – FX conversion & delinquency fee accrual", ICS, "Revenue", "LC", "RM", 2, 2, "Complete", "Medium", "T01", "Detective", ""),
    ("T13", "Interest income & interest expense accrual (net interest income)", CORPORATE, "Accruals", "MR", "EB", 1, 1, "Complete", "High", "T01", "Detective", ""),
    ("T14", "Intercompany reconciliation – GMNS vs card-issuing segments", CORPORATE, "Intercompany", "NK", "EB", 1, None, "Delayed", "High", "T06", "Detective", "Intercompany counterparty mapping gap in new Cloud module."),
    ("T15", "Intercompany elimination journal entries", CORPORATE, "Intercompany", "NK", "EB", 3, None, "Not Started", "High", "T14", "Preventive", "Blocked by T14."),
    ("T16", "Bank reconciliations – operating & payroll accounts", CS, "Reconciliation", "MR", "EB", 1, 1, "Complete", "Medium", "", "Detective", ""),
    ("T17", "Bank reconciliations – merchant settlement accounts", GMNS, "Reconciliation", "DO", "EB", 1, None, "Delayed", "High", "T06", "Detective", "Unreconciled break under investigation with Treasury (fictional)."),
    ("T18", "Prepaid expense amortisation – marketing & software licences", CS, "Amortisation", "MR", "EB", -1, -1, "Complete", "Low", "", "Detective", ""),
    ("T19", "Card Member services & marketing expense accrual", USCS, "Accruals", "SM", "TH", 1, 0, "Complete", "Medium", "", "Detective", ""),
    ("T20", "Manual journal entry review & approval (four-eyes check)", CORPORATE, "Journal Entries", "JW", "TH", 3, None, "In Progress", "High", "", "Preventive", "Manual JE volume up ~30% during Cloud cutover (fictional)."),
    ("T21", "Capitalised software depreciation run", GMNS, "Amortisation", "NK", "EB", -1, -1, "Complete", "Low", "", "Preventive", ""),
    ("T22", "Balance sheet account reconciliation certification", CORPORATE, "Reconciliation", "AP", "EB", 3, None, "In Progress", "Medium", "T16", "Detective", ""),
    ("T23", "Flux analysis – USCS & CS segment results vs prior month and forecast", CS, "Analytics", "LC", "TH", 4, None, "Not Started", "Medium", "T03", "Detective", ""),
    ("T24", "Flux analysis – ICS & GMNS segment results vs prior month and forecast", ICS, "Analytics", "LC", "TH", 4, None, "Not Started", "Medium", "T05", "Detective", ""),
    ("T25", "Controller review & close sign-off", CORPORATE, "Review & Sign-off", "JW", "CT", 5, None, "Not Started", "High", "T23", "Preventive", "Final sign-off pack due WD+5."),
]

RISKS = [
    ("R01", "Late subledger feeds from on-prem systems during Oracle Cloud cutover", "H", "H", "Nora Kim", "Daily feed-status call WD-3 to WD+1; fallback manual upload template pre-approved.", "Open"),
    ("R02", "Manual journal entry volume rises during transition, raising error risk", "H", "M", "James Whitaker", "Tiered JE approval thresholds; reviewer sample of 100% for entries above limit.", "Open"),
    ("R03", "Reconciliation breaks between Cloud GL and legacy subledgers", "M", "H", "Elena Brooks", "Parallel run of key reconciliations for two closes; break log with ageing.", "Open"),
    ("R04", "Oracle Cloud cutover window overlaps with WD-1 to WD+1 close activities", "H", "M", "Tom Hughes", "Cutover freeze calendar agreed with IT; go/no-go checkpoint at WD-5.", "Monitoring"),
    ("R05", "Month-end FX rate table loaded late, delaying international revenue entries", "M", "M", "Lena Chen", "Automate rate load; manual override procedure with Treasury sign-off.", "Open"),
    ("R06", "Credit loss reserve model inputs not final by WD+1", "H", "L", "James Whitaker", "Agreed model-run timetable with Credit Risk; preliminary estimate booked if late.", "Monitoring"),
    ("R07", "Key-person dependency on intercompany process owner", "M", "L", "Nora Kim", "Cross-train backup preparer; documented desk procedures.", "Open"),
    ("R08", "Rewards redemption-rate assumptions stale after programme change", "L", "M", "Sofia Martins", "Quarterly assumption review with Loyalty team; flux threshold alert.", "Closed"),
]

TASK_COLS = [
    "Task ID", "Task Name", "Segment", "Category", "Owner", "Reviewer",
    "Workday Due", "Actual Completion Day", "Status", "Priority",
    "Dependency", "Control Type", "Comments",
]
RISK_COLS = ["Risk ID", "Description", "Impact", "Likelihood", "Owner", "Mitigation", "Status"]


def build_tasks() -> pd.DataFrame:
    rows = []
    for (tid, name, seg, cat, own, rev, due, act, status, pri, dep, ctrl, com) in TASKS:
        rows.append([tid, name, seg, cat, PREP[own], REV[rev], due, act, status,
                     pri, dep, ctrl, com])
    df = pd.DataFrame(rows, columns=TASK_COLS)
    df["Actual Completion Day"] = df["Actual Completion Day"].astype("Int64")
    return df


def validate(tasks: pd.DataFrame, risks: pd.DataFrame) -> None:
    assert len(tasks) == 25, "expected 25 tasks"
    assert tasks["Task ID"].is_unique
    assert tasks["Workday Due"].between(-2, 5).all()
    assert set(tasks["Status"]) <= {"Not Started", "In Progress", "Complete", "Delayed"}
    assert set(tasks["Segment"]) <= set(SEGMENTS) | {CORPORATE}
    assert set(SEGMENTS) <= set(tasks["Segment"]), "every segment needs a task"
    deps = set(tasks["Dependency"]) - {""}
    assert deps <= set(tasks["Task ID"]), f"unknown dependency {deps - set(tasks['Task ID'])}"
    complete = tasks["Status"] == "Complete"
    assert tasks.loc[complete, "Actual Completion Day"].notna().all()
    assert tasks.loc[~complete, "Actual Completion Day"].isna().all()
    delayed_share = (tasks["Status"] == "Delayed").mean()
    assert 0.12 <= delayed_share <= 0.20, f"delayed share {delayed_share:.0%}"
    assert len(risks) == 8 and risks["Risk ID"].is_unique
    assert set(risks["Impact"]) | set(risks["Likelihood"]) <= {"H", "M", "L"}


def main() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    tasks = build_tasks()
    risks = pd.DataFrame(RISKS, columns=RISK_COLS)
    validate(tasks, risks)

    ref = pd.DataFrame(
        [("Reportable segment", s) for s in SEGMENTS]
        + [("Not a reportable segment", CORPORATE)]
        + [("Revenue line item", r) for r in REVENUE_LINES],
        columns=["Element", "Name (as written in 10-K)"],
    )

    tasks.to_csv(DATA_DIR / "close_tasks.csv", index=False)
    risks.to_csv(DATA_DIR / "risk_log.csv", index=False)
    ref.to_csv(DATA_DIR / "reference_structure.csv", index=False)

    delayed = (tasks["Status"] == "Delayed").sum()
    print(f"Wrote {len(tasks)} tasks ({delayed} delayed = {delayed/len(tasks):.0%}), "
          f"{len(risks)} risks to {DATA_DIR}")


if __name__ == "__main__":
    main()
