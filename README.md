# Month-End Close Tracker – Card Issuer Close Simulation

A Python-built Excel tracker that gives finance leadership one live view of two things at once:
- **the recurring month-end close:** 25 close tasks for a fictional card issuer, organised by reportable segment;
- **the one-time Oracle on-premises to Oracle Cloud migration programme:** 14 milestones from fit-gap to post go-live support.

It includes KPIs, RAG status, a risk log linked to migration milestones, and owner accountability, all driven by Excel formulas so the workbook updates as statuses and dates change.

![Leadership Summary](screenshots/leadership_summary.png)

> **Disclaimer:** Fictional company and data for portfolio purposes. Structure inspired by public information in American Express's Form 10-K. Not affiliated with or endorsed by American Express. Contains no confidential data.

---

## Key features

- **Leadership Summary:** close KPI tiles and an overall close RAG, plus a **Migration Programme Status** block (total, complete, on track, at risk, red/late, % complete). Charts: tasks by segment, cumulative planned vs actual completion (WD-2 to WD+5), and migration milestone RAG by phase. Prints on one landscape page.
- **Migration Milestones:** 14 milestones across 8 phases: Fit-Gap, Design, Data Migration, System Configuration, Testing (SIT), Testing (UAT), Go-Live and Post Go-Live Support. Each has a stakeholder (Process Owner / Product Owner / Technology / System Integrator), planned and forecast dates (DD-MMM-YYYY), status, a Days Variance formula, and a RAG formula (Red if Blocked or more than 5 days late; Green if Complete or not late; Amber if 1–5 days late).
- **Live formulas:** every KPI, RAG, segment count, owner count and chart series is a COUNTIFS, MINIFS or IF formula on the task list. No values are hard-coded.
- **Editable inputs:** the as-of workday and the Amber/Red RAG thresholds are yellow input cells. Changing them moves the close forward or tightens tolerance.
- **Close Calendar:** 25 tasks with frozen headers, filters on every column, dropdowns for Status/Priority/Control Type, and RAG conditional formatting.
- **Risk Log:** 8 ERP-transition risks scored Impact × Likelihood, with RAG formatting and a summary block. Three risks link to the migration milestone they threaten, and that milestone's RAG is looked up live.
- **Owner View:** open, delayed and high-priority open tasks per preparer, earliest open due day, owner RAG, and a chart.
- **Reproducible:** two scripts rebuild the data and the workbook from scratch. The data generator validates itself (task count, delay share, dependencies, status consistency).

## Headline results

| | |
|---|---|
| Close tasks | **25** across 4 reportable segments plus Corporate & Other |
| Live KPIs (as of WD+2) | **48%** complete · **4** delayed (16%) · **10** high-priority open |
| RAG logic | Overall close **Amber**. Red if % delayed ≥ 20%, Amber if ≥ 5% (both editable) |
| Migration programme (as of 13-Nov-2026) | **14** milestones · **6** complete (**43%**) · **5** on track · **3** at risk · **1** red (mock data load 2, 14 days late) |
| Risk log | **8** risks: 4 Red, 2 Amber, 2 Green · 3 linked to migration milestones |
| Formula integrity | **227** formulas, **0** errors after a full LibreOffice recalculation |
| Validation | Every close KPI, milestone count, % complete, per-row Days Variance and RAG, and RAG-by-phase count matches an independent pandas calculation. An edit test (Delayed → Complete) updated all dependent close KPIs |

![Migration Milestones](screenshots/migration_milestones.png)

<details>
<summary>Close Calendar and Risk Log screenshots</summary>

![Close Calendar](screenshots/close_calendar.png)

![Risk Log](screenshots/risk_log.png)

</details>

---

## How to run

Requires Python 3.11+.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

```bash
python src/generate_data.py
python src/build_tracker.py
```

Output: `output/Month_End_Close_Tracker.xlsx`. Open it in Excel. LibreOffice and Google Sheets also work.

**Try it:** on *Close Calendar*, change T05 from `Delayed` to `Complete` and enter `3` in its Actual Completion Day. The % complete, # delayed, overall RAG, segment chart, planned-vs-actual line and Owner View all update. On *Migration Milestones*, set M06's Forecast Date back to its Planned Date: its RAG turns Green, the Red / late tile drops to 0, and the Risk Log shows R01's linked milestone as Green.

### Folder structure

```
├── src/
│   ├── generate_data.py   # builds + validates the fictional sample data -> data/*.csv
│   └── build_tracker.py   # reads the CSVs, writes the formatted, formula-driven workbook
├── data/
│   ├── close_tasks.csv          # 25 close tasks
│   ├── risk_log.csv             # 8 close risks (3 linked to milestones)
│   ├── milestones.csv           # 14 Oracle Cloud migration milestones
│   └── reference_structure.csv  # verified 10-K segment & revenue-line names
├── output/
│   └── Month_End_Close_Tracker.xlsx
├── screenshots/
├── requirements.txt
└── README.md
```

---

## Sheet guide

| # | Sheet | What it shows |
|---|-------|---------------|
| 1 | **Leadership Summary** | Disclaimer banner. Editable inputs: as-of workday and RAG thresholds. Close KPI tiles: total tasks, % complete, # delayed, # high-priority open, and the **overall close RAG**. Beside them, the **Migration Programme Status** tiles: total milestones, complete, on track, at risk, red/late, % complete. Also a status-by-segment table and chart, a **planned vs actual completion** table and line chart, and a milestone **RAG-by-phase** table and chart. One landscape page. |
| 2 | **Close Calendar** | All 25 tasks: Task ID, Task Name, Segment, Category, Owner, Reviewer, Workday Due, Actual Completion Day, Status, Priority, Dependency, Control Type, Comments, plus a formula-driven RAG column. Red = Delayed, Amber = open and due on/before the as-of day, Green = complete or not yet due. |
| 3 | **Migration Milestones** | 14 Oracle on-prem to Cloud milestones: Milestone ID, Phase, Milestone, Stakeholder, Owner, Planned Date, Forecast Date, Status, Days Variance (`=Forecast - Planned`), RAG (formula) and Comments. Filters, frozen header, status dropdowns, and RAG conditional formatting. |
| 4 | **Risk Log** | 8 close risks. Risk Score = Impact × Likelihood (H=3, M=2, L=1). RAG: Red ≥ 6, Amber 3–4, Green ≤ 2. A **Linked Milestone** column ties R01 → M06 (late subledger feeds → mock data load 2), R03 → M10 (reconciliation breaks → SIT) and R04 → M12 (cutover overlap → cutover plan), and shows each milestone's live RAG. Includes a risk summary block. |
| 5 | **Owner View** | For each preparer: total assigned, complete, open, delayed, high-priority open, earliest open due day (MINIFS) and an owner RAG, plus an open-vs-delayed column chart. |
| 6 | **About & Sources** | Disclaimer, the 10-K citation (fiscal year + URL), and a breakdown of what is real (public structure) versus fictional (everything else). |

### Key formulas

- `% complete` → `=COUNTIFS(Status,"Complete")/Total`
- `# high-priority open` → `=COUNTIFS(Priority,"High",Status,"<>Complete")`
- Overall RAG → `=IF(Delayed/Total>=RedThreshold,"Red",IF(Delayed/Total>=AmberThreshold,"Amber","Green"))`
- Planned (cumulative) → `=COUNTIFS(WorkdayDue,"<="&WD)`; Actual → `=COUNTIFS(Status,"Complete",ActualDay,"<="&WD)`
- Risk score → nested `IF` mapping H/M/L to 3/2/1, multiplied together
- Milestone Days Variance → `=ForecastDate-PlannedDate`
- Milestone RAG → `=IF(Status="Blocked","Red",IF(OR(Status="Complete",DaysVariance<=0),"Green",IF(DaysVariance<=5,"Amber","Red")))`
- Programme tiles → `=COUNTIFS(MilestoneStatus,"Complete")` etc.; RAG by phase → `=COUNTIFS(Phase,p,RAG,"Red")`

Workdays are stored as integers (for example `-2`, `0`, `5`) and displayed as `WD-2`, `WD0`, `WD+5` using a custom number format. This keeps them sortable and usable in COUNTIFS.

---

## Skills demonstrated

- **Record-to-report / month-end close process design** for a card issuer: revenue recognition (card fee deferral, discount revenue), rewards liability, credit loss reserves, network settlement, intercompany, bank recs, flux analysis and controller sign-off. Each task has dependencies and preventive/detective control types.
- **Leadership reporting:** KPI design, RAG frameworks with documented thresholds, and segment and owner accountability views.
- **ERP programme tracking:** an Oracle Cloud migration plan from fit-gap through hypercare, with stakeholder ownership, schedule variance and risks linked to milestones.
- **Risk management:** an impact × likelihood scoring model with mitigations tied to ERP transition risks.
- **Excel engineering:** live COUNTIFS/MINIFS formulas, cross-sheet references, conditional formatting, data validation, frozen panes, AutoFilter and native Excel charts.
- **Python automation:** pandas for data generation and validation, openpyxl for workbook construction. The build is fully reproducible.
- **Source discipline:** public data was checked against the primary SEC filing and cited. Anything unverifiable was left out. Real and fictional content are clearly separated.

---

## How this was built

I designed the close process, KPIs, RAG logic, milestone tracking, risk scoring and validation approach. I used Claude (AI) to help write the Python code, then checked every metric against an independent pandas calculation and verified all public data against the SEC 10-K.

---

## Data sources & disclaimer

### Business context

Meridian Card Services (fictional) is moving its general ledger from Oracle on-premises (EBS) to Oracle Cloud. During the transition, subledger feeds arrive late, manual journal volume rises, reconciliations break between the new Cloud GL and legacy subledgers, and leadership loses the visibility the old close checklist gave. This tracker answers the four questions the Controller and CFO ask every day of the close: **Are we on track? What is late? Who is it waiting on? What could still derail sign-off?**

### Real vs fictional

**Real (public structure only)**, verified against American Express's **Form 10-K for the fiscal year ended December 31, 2025**, filed 6 February 2026:
<https://www.sec.gov/Archives/edgar/data/4962/000000496226000080/axp-20251231.htm>

- Reportable operating segments: *U.S. Consumer Services*, *Commercial Services*, *International Card Services*, *Global Merchant and Network Services*. *Corporate & Other* is used as a separate bucket. The 10-K says it is **not** a reportable segment.
- Revenue line names (MD&A Table 2): *Discount revenue*, *Net card fees*, *Service fees and other revenue*, *Total interest income*, *Total interest expense*, *Net interest income*.
- "Processed revenue" is intentionally excluded. The FY2025 10-K says it was folded into Service fees and other revenue starting in Q1 2025.

**Fictional:** the company, all task names, owners, reviewers, workdays, statuses, comments, risks, the Oracle migration scenario and all 14 migration milestones (names, owners, dates and statuses). No financial figures from the 10-K are used, and the workbook contains no real balances.

---

*Fictional company and data for portfolio purposes. Structure inspired by public information in American Express's Form 10-K. Not affiliated with or endorsed by American Express. Contains no confidential data.*
