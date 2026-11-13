"""
build_tracker.py
----------------
Builds output/Month_End_Close_Tracker.xlsx from the CSVs in data/.

Every KPI, RAG, segment/owner summary and chart series is an Excel
formula (COUNTIFS / IF / MINIFS) pointing at the Close Calendar and
Risk Log, so editing a Status cell updates the whole workbook.

Run src/generate_data.py first.
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUT = ROOT / "output" / "Month_End_Close_Tracker.xlsx"

COMPANY = "Meridian Card Services (fictional)"
PERIOD = "November 2026 month-end close (fictional)"
DISCLAIMER = (
    "Fictional company and data for portfolio purposes. Structure inspired by "
    "public information in American Express's Form 10-K. Not affiliated with or "
    "endorsed by American Express. Contains no confidential data."
)
TENK_URL = "https://www.sec.gov/Archives/edgar/data/4962/000000496226000080/axp-20251231.htm"
TENK_INDEX = "https://www.sec.gov/Archives/edgar/data/4962/000000496226000080/0000004962-26-000080-index.htm"

STATUSES = ["Not Started", "In Progress", "Complete", "Delayed"]
AS_OF_DEFAULT = 2  # WD+2

# --- neutral palette (deliberately slate / teal, no third-party branding) ---
INK = "1F2A30"
SLATE = "34495E"
TEAL = "2A7F78"
MUTED = "6B7780"
LINE = "D5DADD"
PANEL = "F3F5F6"
INPUT_FILL = "FFF2CC"
RAG = {
    "Red": ("F8D0CC", "9C1C10"),
    "Amber": ("FCE7B2", "7A4F00"),
    "Green": ("CDEBD3", "1E6B34"),
}
STATUS_COLOURS = {  # chart series colours
    "Not Started": "B8C1C7",
    "In Progress": "5B8DB8",
    "Complete": "3C9A5F",
    "Delayed": "C8553D",
}

FONT = "Arial"
F_BASE = Font(name=FONT, size=10, color=INK)
F_BOLD = Font(name=FONT, size=10, bold=True, color=INK)
F_HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=18, bold=True, color=INK)
F_SUB = Font(name=FONT, size=11, color=MUTED)
F_NOTE = Font(name=FONT, size=9, italic=True, color=MUTED)
F_INPUT = Font(name=FONT, size=10, bold=True, color="0000FF")  # blue = editable input
F_SECTION = Font(name=FONT, size=12, bold=True, color=SLATE)

FILL_HDR = PatternFill("solid", fgColor=SLATE)
FILL_PANEL = PatternFill("solid", fgColor=PANEL)
FILL_INPUT = PatternFill("solid", fgColor=INPUT_FILL)
THIN = Side(style="thin", color=LINE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

WD_FMT = '"WD+"0;"WD-"0;"WD0"'

# Close Calendar geometry (used by every cross-sheet formula)
CAL = "'Close Calendar'"
CAL_HDR_ROW = 4
CAL_FIRST = 5


def rng(col: str, last: int) -> str:
    return f"{CAL}!${col}${CAL_FIRST}:${col}${last}"


def style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = F_HDR
        cell.fill = FILL_HDR
        cell.alignment = CENTER
        cell.border = BORDER


def style_body(ws, r1, r2, c1, c2, wrap_cols=()):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = F_BASE
            cell.border = BORDER
            cell.alignment = WRAP if c in wrap_cols else Alignment(vertical="top")


def title_block(ws, title, subtitle, width_cols):
    ws["A1"] = title
    ws["A1"].font = F_TITLE
    ws["A2"] = subtitle
    ws["A2"].font = F_SUB
    ws.row_dimensions[1].height = 28
    ws.sheet_view.showGridLines = False


def rag_rules(ws, ref, anchor):
    """Conditional fill for a cell range whose text is Red/Amber/Green."""
    for name, (bg, fg) in RAG.items():
        ws.conditional_formatting.add(
            ref,
            FormulaRule(
                formula=[f'UPPER({anchor})="{name.upper()}"'],
                fill=PatternFill("solid", fgColor=bg, bgColor=bg),
                font=Font(name=FONT, bold=True, color=fg),
                stopIfTrue=True,
            ),
        )


def fit_landscape(ws, tall=0):
    """Print landscape, one page wide (tall=0 lets rows flow onto more pages)."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = tall
    ws.sheet_properties.pageSetUpPr.fitToPage = True


def finish_chart(chart, x_title=None, y_title=None):
    chart.roundedCorners = False
    chart.legend.position = "b"
    chart.legend.overlay = False
    if chart.title is not None:
        chart.title.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if x_title:
        chart.x_axis.title = x_title
        chart.x_axis.title.overlay = False
    if y_title:
        chart.y_axis.title = y_title
        chart.y_axis.title.overlay = False
    chart.y_axis.majorGridlines = None
    chart.y_axis.numFmt = "0"


# ---------------------------------------------------------------------------
# Sheet 2: Close Calendar
# ---------------------------------------------------------------------------
def build_calendar(ws, tasks: pd.DataFrame, as_of_ref: str) -> int:
    title_block(ws, "Close Calendar", f"{COMPANY} · {PERIOD} · all 25 close tasks", 14)
    ws["A3"] = ("Edit Status (col I) and Actual Completion Day (col H) — every KPI, chart and "
                "owner count recalculates. RAG: Red = Delayed · Amber = open and due on/before "
                "the as-of workday · Green = Complete or not yet due.")
    ws["A3"].font = F_NOTE

    cols = list(tasks.columns) + ["RAG"]
    for c, name in enumerate(cols, 1):
        ws.cell(row=CAL_HDR_ROW, column=c, value=name)
    style_header(ws, CAL_HDR_ROW, 1, len(cols))
    ws.row_dimensions[CAL_HDR_ROW].height = 44

    for i, row in enumerate(tasks.itertuples(index=False)):
        r = CAL_FIRST + i
        for c, val in enumerate(row, 1):
            if pd.isna(val) or val == "":
                val = None
            elif hasattr(val, "item"):
                val = val.item()  # numpy -> python
            ws.cell(row=r, column=c, value=val)
        ws.cell(row=r, column=14, value=(
            f'=IF(I{r}="Delayed","Red",IF(I{r}="Complete","Green",'
            f'IF(G{r}<={as_of_ref},"Amber","Green")))'
        ))
    last = CAL_FIRST + len(tasks) - 1

    style_body(ws, CAL_FIRST, last, 1, len(cols), wrap_cols=(2, 3, 6, 13))
    for r in range(CAL_FIRST, last + 1):
        for c in (7, 8):
            ws.cell(row=r, column=c).number_format = WD_FMT
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="top")
        for c in (1, 9, 10, 11, 12, 14):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=r, column=8).font = F_INPUT
        ws.cell(row=r, column=9).font = F_INPUT
        ws.row_dimensions[r].height = 30

    widths = [8, 46, 22, 15, 15, 24, 10, 12, 12, 9, 13, 11, 40, 9]
    for c, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.freeze_panes = ws.cell(row=CAL_FIRST, column=3)  # header + Task ID/Name stay visible
    ws.auto_filter.ref = f"A{CAL_HDR_ROW}:N{last}"

    # dropdowns so edits stay consistent with the COUNTIFS criteria
    dv_status = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
    dv_pri = DataValidation(type="list", formula1='"High,Medium,Low"')
    dv_ctrl = DataValidation(type="list", formula1='"Preventive,Detective"')
    dv_wd = DataValidation(type="whole", operator="between", formula1="-2", formula2="10", allow_blank=True)
    for dv, col in ((dv_status, "I"), (dv_pri, "J"), (dv_ctrl, "L"), (dv_wd, "H")):
        ws.add_data_validation(dv)
        dv.add(f"{col}{CAL_FIRST}:{col}{last}")

    # RAG colouring: the RAG column and the Status column
    rag_rules(ws, f"N{CAL_FIRST}:N{last}", f"$N{CAL_FIRST}")
    status_map = {"Delayed": "Red", "In Progress": "Amber", "Complete": "Green"}
    for status, rag in status_map.items():
        bg, fg = RAG[rag]
        ws.conditional_formatting.add(
            f"I{CAL_FIRST}:I{last}",
            CellIsRule(operator="equal", formula=[f'"{status}"'],
                       fill=PatternFill("solid", fgColor=bg, bgColor=bg),
                       font=Font(name=FONT, color=fg, bold=True)),
        )
    ws.print_title_rows = f"{CAL_HDR_ROW}:{CAL_HDR_ROW}"
    fit_landscape(ws, tall=0)
    return last


# ---------------------------------------------------------------------------
# Sheet 1: Leadership Summary
# ---------------------------------------------------------------------------
def build_summary(ws, segments, last):
    title_block(ws, "Month-End Close Tracker – Leadership Summary",
                f"{COMPANY} · {PERIOD} · Card issuer close simulation", 12)
    ws.merge_cells("A3:L3")
    ws["A3"] = DISCLAIMER
    ws["A3"].font = Font(name=FONT, size=9, italic=True, bold=True, color="9C1C10")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[3].height = 28

    STATUS = rng("I", last)
    PRI = rng("J", last)
    DUE = rng("G", last)
    ACT = rng("H", last)
    SEG = rng("C", last)
    IDS = rng("A", last)

    # --- inputs -----------------------------------------------------------
    ws["A5"] = "Inputs (edit yellow cells)"
    ws["A5"].font = F_SECTION
    inputs = [
        (6, "As-of workday", AS_OF_DEFAULT, WD_FMT, "Current day of the close. Drives Amber RAG for open tasks that are due."),
        (7, "Amber if % delayed ≥", 0.05, "0%", "Assumption: leadership tolerance for delay (portfolio choice)."),
        (8, "Red if % delayed ≥", 0.20, "0%", "Assumption: leadership tolerance for delay (portfolio choice)."),
    ]
    for r, label, val, fmt, note in inputs:
        ws.cell(row=r, column=1, value=label).font = F_BOLD
        cell = ws.cell(row=r, column=3, value=val)
        cell.font = F_INPUT
        cell.fill = FILL_INPUT
        cell.number_format = fmt
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=4, value=note).font = F_NOTE

    # --- KPI tiles (row 10 label, row 11 value) ----------------------------
    ws["A10"] = "Close KPIs"
    ws["A10"].font = F_SECTION
    kpis = [
        ("A", "B", "Total tasks", f"=COUNTA({IDS})", "0"),
        ("C", "D", "% complete", f'=IFERROR(COUNTIFS({STATUS},"Complete")/A12,0)', "0%"),
        ("E", "F", "# delayed", f'=COUNTIFS({STATUS},"Delayed")', "0"),
        ("G", "H", "# high-priority open", f'=COUNTIFS({PRI},"High",{STATUS},"<>Complete")', "0"),
        ("I", "J", "Overall close RAG",
         '=IF(A12=0,"Green",IF(E12/A12>=$C$8,"Red",IF(E12/A12>=$C$7,"Amber","Green")))', "@"),
    ]
    for c1, c2, label, formula, fmt in kpis:
        ws.merge_cells(f"{c1}11:{c2}11")
        ws.merge_cells(f"{c1}12:{c2}12")
        ws[f"{c1}11"] = label
        ws[f"{c1}11"].font = Font(name=FONT, size=9, bold=True, color=MUTED)
        ws[f"{c1}11"].alignment = CENTER
        ws[f"{c1}12"] = formula
        ws[f"{c1}12"].font = Font(name=FONT, size=22, bold=True, color=INK)
        ws[f"{c1}12"].alignment = CENTER
        ws[f"{c1}12"].number_format = fmt
        for col in (c1, c2):
            for r in (11, 12):
                ws[f"{col}{r}"].fill = FILL_PANEL
                ws[f"{col}{r}"].border = Border(
                    left=THIN if col == c1 else None, right=THIN if col == c2 else None,
                    top=THIN if r == 11 else None, bottom=THIN if r == 12 else None)
    ws.row_dimensions[12].height = 40
    rag_rules(ws, "I12:J12", "$I$12")
    ws["A13"] = ("RAG rule: Red if % delayed ≥ C8; Amber if ≥ C7; otherwise Green. "
                 "High-priority open = High priority and Status ≠ Complete.")
    ws["A13"].font = F_NOTE
    ws["L12"] = f'=COUNTIFS({PRI},"High",{STATUS},"Delayed")'
    ws["L11"] = "of which high-priority delayed"
    ws["L11"].font = Font(name=FONT, size=9, bold=True, color=MUTED)
    ws["L11"].alignment = CENTER
    ws["L12"].font = Font(name=FONT, size=16, bold=True, color="9C1C10")
    ws["L12"].alignment = CENTER

    # --- status by segment table --------------------------------------------
    t1 = 15
    ws.cell(row=t1, column=1, value="Status by segment").font = F_SECTION
    hdr = t1 + 1
    headers = ["Segment", ""] + STATUSES + ["Total", "% complete"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=hdr, column=c, value=h)
    ws.merge_cells(start_row=hdr, start_column=1, end_row=hdr, end_column=2)
    style_header(ws, hdr, 1, len(headers))
    for i, seg in enumerate(segments):
        r = hdr + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(row=r, column=1, value=seg)
        for j, st in enumerate(STATUSES):
            col = 3 + j
            ws.cell(row=r, column=col, value=f'=COUNTIFS({SEG},$A{r},{STATUS},{get_column_letter(col)}${hdr})')
        ws.cell(row=r, column=7, value=f"=SUM(C{r}:F{r})")
        ws.cell(row=r, column=8, value=f"=IFERROR(E{r}/G{r},0)").number_format = "0%"
    seg_last = hdr + len(segments)
    tot = seg_last + 1
    ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=2)
    ws.cell(row=tot, column=1, value="Total")
    for col in range(3, 8):
        L = get_column_letter(col)
        ws.cell(row=tot, column=col, value=f"=SUM({L}{hdr+1}:{L}{seg_last})")
    ws.cell(row=tot, column=8, value=f"=IFERROR(E{tot}/G{tot},0)").number_format = "0%"
    style_body(ws, hdr + 1, tot, 1, 8)
    for c in range(1, 9):
        ws.cell(row=tot, column=c).font = F_BOLD
    for r in range(hdr + 1, tot + 1):
        for c in range(3, 9):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    ws.cell(row=tot + 1, column=1,
            value="Corporate & Other is shown separately: the 10-K states it is not a reportable segment.").font = F_NOTE

    # --- planned vs actual by workday --------------------------------------
    t2 = tot + 3
    ws.cell(row=t2, column=1, value="Planned vs actual completion (cumulative tasks)").font = F_SECTION
    h2 = t2 + 1
    for c, h in enumerate(["Workday", "", "Planned (due by WD)", "Actual (completed by WD)", "Gap"], 1):
        ws.cell(row=h2, column=c, value=h)
    ws.merge_cells(start_row=h2, start_column=1, end_row=h2, end_column=2)
    style_header(ws, h2, 1, 5)
    wds = list(range(-2, 6))
    for i, wd in enumerate(wds):
        r = h2 + 1 + i
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(row=r, column=1, value=wd).number_format = WD_FMT
        ws.cell(row=r, column=3, value=f'=COUNTIFS({DUE},"<="&$A{r})')
        ws.cell(row=r, column=4, value=f'=COUNTIFS({STATUS},"Complete",{ACT},"<="&$A{r})')
        ws.cell(row=r, column=5, value=f"=D{r}-C{r}").number_format = "0;(0);-"
    wd_last = h2 + len(wds)
    style_body(ws, h2 + 1, wd_last, 1, 5)
    for r in range(h2 + 1, wd_last + 1):
        for c in range(1, 6):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    ws.cell(row=wd_last + 1, column=1,
            value="Actual line flattens after the as-of workday because later days have not happened yet.").font = F_NOTE

    # --- charts --------------------------------------------------------------
    bar = BarChart()
    bar.type = "bar"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.gapWidth = 60
    bar.title = "Task status by segment"
    data = Reference(ws, min_col=3, max_col=6, min_row=hdr, max_row=seg_last)
    cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=seg_last)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    for s, st in zip(bar.series, STATUSES):
        s.graphicalProperties.solidFill = STATUS_COLOURS[st]
        s.graphicalProperties.line.solidFill = STATUS_COLOURS[st]
    finish_chart(bar, y_title="Tasks")
    bar.x_axis.scaling.orientation = "maxMin"  # segment order top-to-bottom
    bar.y_axis.crosses = "max"  # keep value axis at the bottom when reversed
    bar.height, bar.width = 8.5, 17
    ws.add_chart(bar, "J15")

    line = LineChart()
    line.title = "Planned vs actual completion by workday"
    data = Reference(ws, min_col=3, max_col=4, min_row=h2, max_row=wd_last)
    cats = Reference(ws, min_col=1, min_row=h2 + 1, max_row=wd_last)
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    for s, colour, dash in zip(line.series, ("8A969E", TEAL), ("dash", "solid")):
        s.graphicalProperties.line.solidFill = colour
        s.graphicalProperties.line.width = 28575
        s.graphicalProperties.line.dashStyle = dash
        s.marker.symbol = "circle"
        s.marker.size = 6
        s.marker.graphicalProperties.solidFill = colour
        s.marker.graphicalProperties.line.solidFill = colour
        s.smooth = False
    finish_chart(line, x_title="Close workday", y_title="Cumulative tasks")
    line.x_axis.number_format = WD_FMT
    line.height, line.width = 8.5, 17
    ws.add_chart(line, "J33")

    widths = {"A": 16, "B": 16, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13, "H": 13,
              "I": 13, "J": 13, "K": 4, "L": 16}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    ws.freeze_panes = "A5"
    fit_landscape(ws, tall=1)
    return {"seg_hdr": hdr, "seg_last": seg_last, "wd_hdr": h2, "wd_last": wd_last}


# ---------------------------------------------------------------------------
# Sheet 3: Risk Log
# ---------------------------------------------------------------------------
def build_risks(ws, risks: pd.DataFrame):
    title_block(ws, "Close Risk Log", f"{COMPANY} · risks to the {PERIOD}", 9)
    ws["A3"] = "Score = Impact × Likelihood (H = 3, M = 2, L = 1). RAG: Red ≥ 6 · Amber 3–4 · Green ≤ 2. Edit blue cells (H/M/L)."
    ws["A3"].font = F_NOTE
    cols = ["Risk ID", "Description", "Impact", "Likelihood", "Risk Score", "Risk RAG",
            "Owner", "Mitigation", "Status"]
    hdr = 4
    for c, h in enumerate(cols, 1):
        ws.cell(row=hdr, column=c, value=h)
    style_header(ws, hdr, 1, len(cols))
    for i, rk in enumerate(risks.itertuples(index=False)):
        r = hdr + 1 + i
        ws.cell(row=r, column=1, value=rk[0])
        ws.cell(row=r, column=2, value=rk[1])
        ws.cell(row=r, column=3, value=rk[2])
        ws.cell(row=r, column=4, value=rk[3])
        ws.cell(row=r, column=5, value=(
            f'=IF(C{r}="H",3,IF(C{r}="M",2,IF(C{r}="L",1,0)))'
            f'*IF(D{r}="H",3,IF(D{r}="M",2,IF(D{r}="L",1,0)))'))
        ws.cell(row=r, column=6, value=f'=IF(E{r}>=6,"Red",IF(E{r}>=3,"Amber","Green"))')
        ws.cell(row=r, column=7, value=rk[4])
        ws.cell(row=r, column=8, value=rk[5])
        ws.cell(row=r, column=9, value=rk[6])
    last = hdr + len(risks)
    style_body(ws, hdr + 1, last, 1, len(cols), wrap_cols=(2, 8))
    for r in range(hdr + 1, last + 1):
        ws.row_dimensions[r].height = 42
        for c in (1, 3, 4, 5, 6, 9):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=r, column=3).font = F_INPUT
        ws.cell(row=r, column=4).font = F_INPUT
    rag_rules(ws, f"F{hdr+1}:F{last}", f"$F{hdr+1}")
    rag_rules(ws, f"E{hdr+1}:E{last}", f"$F{hdr+1}")
    dv = DataValidation(type="list", formula1='"H,M,L"')
    ws.add_data_validation(dv)
    dv.add(f"C{hdr+1}:D{last}")
    dvs = DataValidation(type="list", formula1='"Open,Monitoring,Closed"')
    ws.add_data_validation(dvs)
    dvs.add(f"I{hdr+1}:I{last}")

    # summary block
    s = last + 2
    ws.cell(row=s, column=1, value="Risk summary").font = F_SECTION
    rows = [("Red risks", f'=COUNTIFS(F{hdr+1}:F{last},"Red")'),
            ("Amber risks", f'=COUNTIFS(F{hdr+1}:F{last},"Amber")'),
            ("Green risks", f'=COUNTIFS(F{hdr+1}:F{last},"Green")'),
            ("Open (not Closed)", f'=COUNTIFS(I{hdr+1}:I{last},"<>Closed")'),
            ("Average score", f"=AVERAGE(E{hdr+1}:E{last})")]
    for i, (label, f) in enumerate(rows):
        r = s + 1 + i
        ws.cell(row=r, column=1, value=label).font = F_BOLD
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        c = ws.cell(row=r, column=3, value=f)
        c.font = F_BOLD
        c.number_format = "0.0" if "Average" in label else "0"
        c.alignment = Alignment(horizontal="center")

    for c, w in enumerate([8, 50, 9, 11, 10, 10, 16, 52, 12], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "C5"
    fit_landscape(ws, tall=1)
    ws.auto_filter.ref = f"A{hdr}:I{last}"
    return hdr + 1, last


# ---------------------------------------------------------------------------
# Sheet 4: Owner View
# ---------------------------------------------------------------------------
def build_owner(ws, owners, last):
    title_block(ws, "Owner View", f"{COMPANY} · open and delayed close tasks by preparer", 8)
    ws["A3"] = "All counts are live COUNTIFS on the Close Calendar. Open = any status other than Complete."
    ws["A3"].font = F_NOTE
    OWN = rng("E", last)
    STATUS = rng("I", last)
    PRI = rng("J", last)
    DUE = rng("G", last)
    cols = ["Owner", "Total assigned", "Complete", "Open", "of which Delayed",
            "High-priority open", "Earliest open due", "Owner RAG"]
    hdr = 5
    for c, h in enumerate(cols, 1):
        ws.cell(row=hdr, column=c, value=h)
    style_header(ws, hdr, 1, len(cols))
    for i, own in enumerate(owners):
        r = hdr + 1 + i
        ws.cell(row=r, column=1, value=own)
        ws.cell(row=r, column=2, value=f"=COUNTIFS({OWN},$A{r})")
        ws.cell(row=r, column=3, value=f'=COUNTIFS({OWN},$A{r},{STATUS},"Complete")')
        ws.cell(row=r, column=4, value=f'=COUNTIFS({OWN},$A{r},{STATUS},"<>Complete")')
        ws.cell(row=r, column=5, value=f'=COUNTIFS({OWN},$A{r},{STATUS},"Delayed")')
        ws.cell(row=r, column=6, value=f'=COUNTIFS({OWN},$A{r},{STATUS},"<>Complete",{PRI},"High")')
        g = ws.cell(row=r, column=7, value=(
            f'=IF(D{r}=0,"—",_xlfn.MINIFS({DUE},{OWN},$A{r},{STATUS},"<>Complete"))'))
        g.number_format = WD_FMT
        ws.cell(row=r, column=8, value=(
            f"=IF(E{r}>0,\"Red\",IF(AND(D{r}>0,G{r}<='Leadership Summary'!$C$6),\"Amber\",\"Green\"))"))
    olast = hdr + len(owners)
    tot = olast + 1
    ws.cell(row=tot, column=1, value="Total")
    for col in range(2, 7):
        L = get_column_letter(col)
        ws.cell(row=tot, column=col, value=f"=SUM({L}{hdr+1}:{L}{olast})")
    style_body(ws, hdr + 1, tot, 1, len(cols))
    for r in range(hdr + 1, tot + 1):
        for c in range(2, 9):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")
    for c in range(1, 9):
        ws.cell(row=tot, column=c).font = F_BOLD
    rag_rules(ws, f"H{hdr+1}:H{olast}", f"$H{hdr+1}")
    ws.cell(row=tot + 1, column=1,
            value="Owner RAG: Red = has a delayed task · Amber = open task due on/before the as-of workday · Green = otherwise.").font = F_NOTE

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Open vs delayed tasks by owner"
    data = Reference(ws, min_col=4, max_col=5, min_row=hdr, max_row=olast)
    cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=olast)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    for s, colour in zip(chart.series, (STATUS_COLOURS["In Progress"], STATUS_COLOURS["Delayed"])):
        s.graphicalProperties.solidFill = colour
        s.graphicalProperties.line.solidFill = colour
    finish_chart(chart, y_title="Tasks")
    chart.height, chart.width = 8, 18
    ws.add_chart(chart, f"A{tot + 3}")

    for c, w in enumerate([20, 13, 11, 9, 13, 14, 15, 11], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "B6"
    fit_landscape(ws, tall=1)


# ---------------------------------------------------------------------------
# Sheet 5: About & Sources
# ---------------------------------------------------------------------------
def build_about(ws, ref: pd.DataFrame):
    title_block(ws, "About & Sources", "What is real, what is fictional, and where the real parts come from", 4)
    ws.merge_cells("A4:D4")
    ws["A4"] = DISCLAIMER
    ws["A4"].font = Font(name=FONT, size=11, bold=True, color="9C1C10")
    ws["A4"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[4].height = 36

    r = 6
    ws.cell(row=r, column=1, value="Source citation").font = F_SECTION
    cite = [
        ("Document", "American Express Company, Annual Report on Form 10-K"),
        ("Fiscal year", "Fiscal year ended December 31, 2025 (FY2025)"),
        ("Filed", "February 6, 2026 (SEC accession no. 0000004962-26-000080)"),
        ("URL", TENK_URL),
        ("Filing index", TENK_INDEX),
        ("Sections used", "MD&A – Business Introduction (segment names); MD&A Table 2, "
                          "'Total revenues net of interest expense summary' (revenue line names)."),
        ("Checked", "November 2026, against the filing on SEC EDGAR. No financial figures were taken."),
    ]
    for i, (k, v) in enumerate(cite):
        rr = r + 1 + i
        ws.cell(row=rr, column=1, value=k).font = F_BOLD
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        c = ws.cell(row=rr, column=2, value=v)
        c.font = F_BASE
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if v.startswith("http"):
            c.hyperlink = v
            c.font = Font(name=FONT, size=10, color="1F5F8B", underline="single")
    r = r + len(cite) + 2

    ws.cell(row=r, column=1, value="Real (public structure only, as named in the 10-K)").font = F_SECTION
    r += 1
    for c, h in enumerate(["Element", "Name", "", "Used in this workbook as"], 1):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, r, 1, 4)
    use = {
        "Reportable segment": "Segment column values; status-by-segment table and chart",
        "Not a reportable segment": "Bucket for corporate / consolidation tasks (kept separate from segments)",
        "Revenue line item": "Names in revenue / accrual task descriptions only; no amounts",
    }
    first = r + 1
    for row in ref.itertuples(index=False):
        r += 1
        ws.cell(row=r, column=1, value=row[0])
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(row=r, column=2, value=row[1])
        ws.cell(row=r, column=4, value=use[row[0]])
    style_body(ws, first, r, 1, 4, wrap_cols=(4,))
    r += 1
    notes = [
        "Verification notes:",
        "• The 10-K names four reportable operating segments: U.S. Consumer Services (USCS), Commercial Services (CS), "
        "International Card Services (ICS) and Global Merchant and Network Services (GMNS). Corporate functions are "
        "'included in Corporate & Other', which is not a segment.",
        "• 'Processed revenue' is deliberately NOT used: the FY2025 10-K says it was consolidated within "
        "'Service fees and other revenue' from Q1 2025 and renamed network partnership revenue.",
        "• Revenue line names are taken from MD&A Table 2. Anything that could not be verified in the filing was left out.",
    ]
    for n in notes:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        c = ws.cell(row=r, column=1, value=n)
        c.font = F_BOLD if n.endswith(":") else F_BASE
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 15 if n.endswith(":") else 30

    r += 2
    ws.cell(row=r, column=1, value="Fictional (everything else)").font = F_SECTION
    r += 1
    fict = [
        ("Company", "Meridian Card Services is invented. No logo, colours or branding of any real company are used."),
        ("Close tasks", "All 25 task names, categories, dependencies, control types and comments."),
        ("People", "All owners, reviewers and risk owners are invented names."),
        ("Dates & status", "Workdays, actual completion days, statuses and the November 2026 close period."),
        ("Risks", "All 8 risks, ratings, mitigations and statuses."),
        ("Balances", "The workbook contains no real balances or amounts. Any amounts referenced in comments are fictional."),
        ("Systems context", "The Oracle on-prem to Oracle Cloud transition is a hypothetical scenario for this portfolio piece."),
    ]
    for c, h in enumerate(["Area", "What is fictional", "", ""], 1):
        ws.cell(row=r, column=c, value=h)
    style_header(ws, r, 1, 4)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    first = r + 1
    for k, v in fict:
        r += 1
        ws.cell(row=r, column=1, value=k)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(row=r, column=2, value=v)
    style_body(ws, first, r, 1, 4, wrap_cols=(2,))

    r += 2
    ws.cell(row=r, column=1, value="How to use").font = F_SECTION
    for line in [
        "• Change a task's Status on Close Calendar (dropdown) and its Actual Completion Day: KPIs, RAG, charts and Owner View update.",
        "• Change the as-of workday on Leadership Summary (C6) to move the close forward.",
        "• Blue text = editable input; yellow fill = key assumption. Black = formula.",
        "• Rebuild from scratch: python src/generate_data.py && python src/build_tracker.py",
    ]:
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        c = ws.cell(row=r, column=1, value=line)
        c.font = F_BASE
        c.alignment = Alignment(wrap_text=True)
    for c, w in enumerate([24, 34, 20, 60], 1):
        ws.column_dimensions[get_column_letter(c)].width = w


def main() -> None:
    tasks = pd.read_csv(DATA_DIR / "close_tasks.csv", keep_default_na=False,
                        na_values={"Actual Completion Day": [""]})
    tasks["Actual Completion Day"] = tasks["Actual Completion Day"].astype("Int64")
    risks = pd.read_csv(DATA_DIR / "risk_log.csv")
    ref = pd.read_csv(DATA_DIR / "reference_structure.csv")

    segments = ref.loc[ref["Element"] == "Reportable segment", "Name (as written in 10-K)"].tolist()
    segments += ref.loc[ref["Element"] == "Not a reportable segment", "Name (as written in 10-K)"].tolist()
    owners = sorted(tasks["Owner"].unique())

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Leadership Summary"
    ws_cal = wb.create_sheet("Close Calendar")
    ws_risk = wb.create_sheet("Risk Log")
    ws_own = wb.create_sheet("Owner View")
    ws_about = wb.create_sheet("About & Sources")

    last = build_calendar(ws_cal, tasks, "'Leadership Summary'!$C$6")
    build_summary(ws_sum, segments, last)
    build_risks(ws_risk, risks)
    build_owner(ws_own, owners, last)
    build_about(ws_about, ref)

    ws_sum["C6"].comment = Comment("Input: current close workday (WD0 = last business day). "
                                   "Default WD+2 for this simulation.", "Close Tracker")
    for ws, colour in ((ws_sum, SLATE), (ws_cal, TEAL), (ws_risk, "9C1C10"),
                       (ws_own, TEAL), (ws_about, MUTED)):
        ws.sheet_properties.tabColor = colour

    wb.properties.title = "Month-End Close Tracker – Card Issuer Close Simulation"
    wb.properties.creator = "Portfolio project"
    wb.properties.description = DISCLAIMER
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
