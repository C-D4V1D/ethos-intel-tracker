"""Rebuild the formatted Excel workbook from data/*.csv.

Everything a reader sees is produced here: tab order, headers, colours,
dropdowns, formulas (priority score, freshness, dashboard totals), conditional
formatting and charts. Formulas stay live, so edits in Excel recalculate.
"""
from __future__ import annotations

import datetime as dt
import math
from urllib.parse import urlparse

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from . import layout as LY
from .common import OUTPUT, WORKBOOK_NAME, load_config, log, now, read_table

# ---------------------------------------------------------------- style kit
FONT = "Arial"
NAVY, TEAL, INK, MUTED = "0F2537", "00707C", "1A2530", "5B6B78"
PALE, TEAL_PALE, LINE = "F2F6F8", "E8F4F5", "D5DEE4"
GREEN_F, AMBER_F, RED_F = "DDF1E4", "FFF1D6", "FBE0DD"
GREEN_T, AMBER_T, RED_T = "1E6B3A", "8A5A00", "A12A1C"

def F(size=10, bold=False, color=INK, italic=False, underline=None):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic, underline=underline)

def fill(hex_):
    return PatternFill("solid", start_color=hex_, end_color=hex_)

THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOTTOM = Border(bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
FIRST = LY.HEADER_ROW + 1
MAXR = 400  # formula ranges reach this far so new rows are picked up


def to_date(s: str):
    try:
        return dt.date.fromisoformat(s[:10])
    except (ValueError, TypeError):
        return None


def to_num(s: str):
    try:
        x = float(str(s).replace(",", ""))
        return int(x) if x.is_integer() else x
    except (ValueError, TypeError):
        return None


def title_block(ws, title: str, purpose: str, key: str = "", span: int = 10):
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = F(16, True, NAVY)
    ws["A2"] = purpose
    ws["A2"].font = F(10, color=MUTED)
    if key:
        ws["A3"] = key
        ws["A3"].font = F(9, italic=True, color=MUTED)
    ws.row_dimensions[1].height = 26
    for r in (1, 2, 3):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(span, 2))


def header_row(ws, cols):
    for i, (_k, label, width, kind) in enumerate(cols, start=1):
        c = ws.cell(LY.HEADER_ROW, i, label)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(TEAL if kind == "formula" else NAVY)
        c.alignment = CENTER
        c.border = BOX
        ws.column_dimensions[L(i)].width = width
    ws.row_dimensions[LY.HEADER_ROW].height = 44
    ws.freeze_panes = ws.cell(FIRST, 3)


def est_height(values_widths) -> float:
    lines = 1
    for text, width in values_widths:
        if text:
            per = max(int(width * 1.15), 4)
            n = sum(math.ceil(max(len(part), 1) / per) for part in str(text).split("\n"))
            lines = max(lines, n)
    return min(15 * lines + 3, 150)


def write_rows(ws, cols, rows, formulas: dict, lists_ref: dict):
    """Write data rows; formula columns use formulas[key](row_number)."""
    for i, rec in enumerate(rows):
        r = FIRST + i
        hw = []
        for ci, (k, _label, width, kind) in enumerate(cols, start=1):
            c = ws.cell(r, ci)
            c.font = F(10)
            c.alignment = WRAP_TOP
            c.border = BOX
            if kind == "formula":
                c.value = formulas[k](r)
                c.alignment = Alignment(horizontal="center", vertical="top")
                c.fill = fill(TEAL_PALE)
                continue
            v = rec.get(k, "")
            if kind == "num":
                n = to_num(v)
                c.value = n if n is not None else (v or None)
                c.number_format = "#,##0.##"
                c.alignment = Alignment(horizontal="right", vertical="top")
            elif kind == "date":
                d = to_date(v)
                c.value = d or (v or None)
                c.number_format = "dd mmm yyyy"
            elif kind == "url" and v:
                host = urlparse(v).netloc.replace("www.", "") or v
                c.value = host
                c.hyperlink = v
                c.font = F(10, color=TEAL, underline="single")
            else:
                c.value = v or None
            hw.append((v, width))
        ws.row_dimensions[r].height = est_height(hw)
    last = FIRST + max(len(rows), 1) - 1
    ws.auto_filter.ref = f"A{LY.HEADER_ROW}:{L(len(cols))}{last}"
    # dropdowns
    for ci, (_k, _l, _w, kind) in enumerate(cols, start=1):
        if kind.startswith("list:"):
            ref = lists_ref[kind[5:]]
            dv = DataValidation(type="list", formula1=ref, allow_blank=True,
                                showErrorMessage=False)
            dv.add(f"{L(ci)}{FIRST}:{L(ci)}{MAXR}")
            ws.add_data_validation(dv)
    return last


def col(cols, key) -> str:
    return L([k for k, *_ in cols].index(key) + 1)


# ---------------------------------------------------------------- sheets
def sheet_lists(wb):
    ws = wb.create_sheet("Lists")
    title_block(ws, "Lists & scoring rules",
                "Dropdown values and the points behind the Priority score. Change a number here and every score updates.",
                span=8)
    refs, c = {}, 1
    for name, values in LY.LISTS.items():
        ws.cell(LY.HEADER_ROW, c, name.title()).font = F(10, True, "FFFFFF")
        ws.cell(LY.HEADER_ROW, c).fill = fill(NAVY)
        for i, v in enumerate(values):
            ws.cell(FIRST + i, c, v).font = F(10)
        refs[name] = f"=Lists!${L(c)}${FIRST}:${L(c)}${FIRST + len(values) - 1}"
        ws.column_dimensions[L(c)].width = 18
        c += 1
    # scoring tables below
    r0 = FIRST + 18
    ws.cell(r0, 1, "Priority score = stage points + relationship points + confidence points + size points (max 12)").font = F(10, True, NAVY)
    tables = [("Stage", LY.STAGE_POINTS), ("Ethos relationship", LY.REL_POINTS),
              ("Confidence", LY.CONF_POINTS)]
    score_refs = {}
    cc = 1
    for name, pts in tables:
        ws.cell(r0 + 1, cc, name).font = F(10, True, "FFFFFF")
        ws.cell(r0 + 1, cc).fill = fill(TEAL)
        ws.cell(r0 + 1, cc + 1, "Points").font = F(10, True, "FFFFFF")
        ws.cell(r0 + 1, cc + 1).fill = fill(TEAL)
        for i, (k, v) in enumerate(pts.items()):
            ws.cell(r0 + 2 + i, cc, k).font = F(10)
            ws.cell(r0 + 2 + i, cc + 1, v).font = F(10, color="0000FF")
        a, b = r0 + 2, r0 + 1 + len(pts)
        score_refs[name] = (f"Lists!${L(cc)}${a}:${L(cc)}${b}", f"Lists!${L(cc+1)}${a}:${L(cc+1)}${b}")
        cc += 3
    ws.cell(r0 + 9, 1, "Size points: 3 if 200 MW or more, 2 if 50-199 MW, 1 if under 50 MW, 0 if unknown").font = F(10, color=MUTED)
    ws.cell(r0 + 10, 1, "Priority band: High = 9+, Medium = 6-8, Low = under 6").font = F(10, color=MUTED)
    ws.cell(r0 + 12, 1, "Spain decree threshold (%)").font = F(10, True)
    ws.cell(r0 + 12, 3, load_config()["grid_data"]["decree_threshold_pct"]).font = F(10, color="0000FF")
    score_refs["threshold"] = f"Lists!$C${r0 + 12}"
    ws.sheet_properties.tabColor = "9AA7B1"
    return refs, score_refs


def sheet_pipeline(wb, rows, refs, sc):
    ws = wb.create_sheet("Pipeline")
    cols = LY.PIPELINE
    title_block(ws, "Project & opportunity pipeline",
                "Named European data centre projects where Ethos could win work. White cells are editable; teal cells are formulas.",
                "Priority = stage + relationship + confidence + size (rules on the Lists tab). Days since check turns amber at 90, red at 180.",
                span=14)
    header_row(ws, cols)
    st, rel, conf, mw, chk = (col(cols, k) for k in ("stage", "ethos_relationship", "confidence", "capacity_mw", "last_checked"))
    pr = col(cols, "_priority")
    fx = {
        "_priority": lambda r: (
            f"=IFERROR(INDEX({sc['Stage'][1]},MATCH({st}{r},{sc['Stage'][0]},0)),0)"
            f"+IFERROR(INDEX({sc['Ethos relationship'][1]},MATCH({rel}{r},{sc['Ethos relationship'][0]},0)),0)"
            f"+IFERROR(INDEX({sc['Confidence'][1]},MATCH({conf}{r},{sc['Confidence'][0]},0)),0)"
            f"+IF({mw}{r}>=200,3,IF({mw}{r}>=50,2,IF(N({mw}{r})>0,1,0)))"),
        "_tier": lambda r: f'=IF({pr}{r}>=9,"High",IF({pr}{r}>=6,"Medium","Low"))',
        "_age": lambda r: f'=IF({chk}{r}="","",TODAY()-{chk}{r})',
    }
    last = write_rows(ws, cols, rows, fx, refs)
    # hidden sort key for the dashboard's top-10 (priority + tiny row tie-break)
    kc = len(cols) + 1
    ws.cell(LY.HEADER_ROW, kc, "_sortkey")
    for r in range(FIRST, last + 1):
        ws.cell(r, kc, f"={pr}{r}+(1000-ROW())/100000")
    ws.column_dimensions[L(kc)].hidden = True
    add_age_rules(ws, col(cols, "_age"), last)
    tier = col(cols, "_tier")
    for txt, f_, t in (("High", GREEN_F, GREEN_T), ("Medium", AMBER_F, AMBER_T), ("Low", PALE, MUTED)):
        ws.conditional_formatting.add(f"{tier}{FIRST}:{tier}{MAXR}",
            CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(f_), font=F(10, True, t)))
    for c in (col(cols, "id"), col(cols, "project")):
        for r in range(FIRST, last + 1):
            ws[f"{c}{r}"].font = F(10, True, NAVY)
    ws.sheet_properties.tabColor = NAVY
    return last, kc


def add_age_rules(ws, c, last):
    rng = f"{c}{FIRST}:{c}{MAXR}"
    cfg = load_config()["freshness"]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({c}{FIRST}),{c}{FIRST}>={cfg["red_days"]})'],
                                                   fill=fill(RED_F), font=F(10, True, RED_T)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({c}{FIRST}),{c}{FIRST}>={cfg["amber_days"]})'],
                                                   fill=fill(AMBER_F), font=F(10, True, AMBER_T)))


def sheet_simple(wb, name, cols, rows, title, purpose, key="", fx=None, tab=TEAL, span=12, refs=None):
    ws = wb.create_sheet(name)
    title_block(ws, title, purpose, key, span)
    header_row(ws, cols)
    last = write_rows(ws, cols, rows, fx or {}, refs or {})
    ws.sheet_properties.tabColor = tab
    return ws, last


def rag_rules(ws, c):
    rng = f"{c}{FIRST}:{c}{MAXR}"
    for txt, f_, t in (("Red", RED_F, RED_T), ("Amber", AMBER_F, AMBER_T), ("Green", GREEN_F, GREEN_T)):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                                      fill=fill(f_), font=F(10, True, t)))


# ---------------------------------------------------------------- dashboard
def tile(ws, r, c, label, formula, fmt="#,##0", note=""):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
    ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
    ws.merge_cells(start_row=r + 2, start_column=c, end_row=r + 2, end_column=c + 1)
    for rr in (r, r + 1, r + 2):
        for cc in (c, c + 1):
            ws.cell(rr, cc).fill = fill(PALE)
    top = ws.cell(r, c, label)
    top.font = F(9, True, MUTED)
    top.alignment = Alignment(horizontal="left", vertical="bottom", indent=1)
    v = ws.cell(r + 1, c, formula)
    v.font = F(20, True, NAVY)
    v.number_format = fmt
    v.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    n = ws.cell(r + 2, c, note)
    n.font = F(8, color=MUTED)
    n.alignment = Alignment(horizontal="left", vertical="top", indent=1, wrap_text=True)
    ws.cell(r - 1, c).border = Border(bottom=Side(style="thick", color=TEAL))
    ws.cell(r - 1, c + 1).border = Border(bottom=Side(style="thick", color=TEAL))


def section(ws, r, c, text, span=6):
    ws.cell(r, c, text.upper()).font = F(10, True, TEAL)
    for cc in range(c, c + span):
        ws.cell(r, cc).border = BOTTOM


def sheet_dashboard(wb, pipe_last, sortcol, countries, cfg):
    ws = wb["Dashboard"]
    ws.sheet_view.showGridLines = False
    for i, w in enumerate([2, 17, 11, 17, 15, 15, 19, 15, 11, 17, 13, 2], start=1):
        ws.column_dimensions[L(i)].width = w
    ws.merge_cells("B1:K1")
    ws["B1"] = cfg["workbook"]["title"]
    ws["B1"].font = F(18, True, NAVY)
    ws.merge_cells("B2:K2")
    ws["B2"] = (f"Where Ethos can win data centre work in Europe, how fresh each fact is, and what changed. "
                f"Refreshed automatically each week. Last build: {now()}.")
    ws["B2"].font = F(10, color=MUTED)
    ws.row_dimensions[1].height = 30
    P = "Pipeline!"
    pc = {k: col(LY.PIPELINE, k) for k in ("project", "country", "capacity_mw", "stage", "ethos_relationship",
                                           "_tier", "_priority", "developer", "next_action", "ethos_fit")}
    rng = lambda k: f"{P}${pc[k]}${FIRST}:${pc[k]}${MAXR}"
    tile(ws, 5, 2, "PROJECTS TRACKED", f"=COUNTA({rng('project')})", note="Named projects across target markets")
    tile(ws, 5, 4, "MW TRACKED", f"=SUM({rng('capacity_mw')})", note="Sum of published capacity (IT or grid)")
    tile(ws, 5, 6, "IN THE CONSULTANT WINDOW",
         f'=COUNTIF({rng("stage")},"Announced")+COUNTIF({rng("stage")},"Permitting")+COUNTIF({rng("stage")},"Permitted")',
         note="Announced, permitting or permitted")
    tile(ws, 5, 8, "WITH EXISTING CLIENTS", f'=COUNTIF({rng("ethos_relationship")},"Existing client")',
         note="Projects of accounts Ethos already serves")
    tile(ws, 5, 10, "SPAIN RENEWABLE SHARE",
         f"=IF(COUNT('Live Data'!$B${FIRST}:$B${MAXR})=0,\"n/a\",INDEX('Live Data'!$B${FIRST}:$B${MAXR},COUNT('Live Data'!$B${FIRST}:$B${MAXR}))/100)",
         fmt="0.0%", note="Latest month (Red Eléctrica). Decree eases at 90%")
    # health row
    tile(ws, 10, 2, "HIGH-PRIORITY PROJECTS", f'=COUNTIF({rng("_tier")},"High")', note="Priority score 9+")
    tile(ws, 10, 4, "SOURCES DUE A RE-CHECK",
         f'=COUNTIF(Sources!$H${FIRST}:$H${MAXR},">="&{cfg["freshness"]["amber_days"]})',
         note=f"Not checked for {cfg['freshness']['amber_days']}+ days")
    tile(ws, 10, 6, "BROKEN / BLOCKED LINKS",
         f'=COUNTIF(Sources!$I${FIRST}:$I${MAXR},"BROKEN")+COUNTIF(Sources!$I${FIRST}:$I${MAXR},"BLOCKED")',
         note="From the weekly link check")
    tile(ws, 10, 8, "AI FLAGS TO REVIEW", f"=COUNTIF('AI Scan'!$D${FIRST}:$D${MAXR},\"Needs review\")",
         note="Google AI Mode answers that disagree")
    tile(ws, 10, 10, "OPEN INTELLIGENCE GAPS", f'=COUNTIF(Gaps!$F${FIRST}:$F${MAXR},"Open")',
         note="Questions to close with clients")

    # by country
    r = 15
    section(ws, r, 2, "Pipeline by market", 4)
    for i, h in enumerate(["Market", "Projects", "MW", "Score /25"]):
        c = ws.cell(r + 1, 2 + i, h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = CENTER
    for j, ctry in enumerate(countries):
        rr = r + 2 + j
        ws.cell(rr, 2, ctry).font = F(10, True, NAVY)
        ws.cell(rr, 3, f'=COUNTIF({rng("country")},B{rr})').font = F(10)
        ws.cell(rr, 4, f'=SUMIF({rng("country")},B{rr},{rng("capacity_mw")})').font = F(10)
        ws.cell(rr, 4).number_format = "#,##0"
        ws.cell(rr, 5, f"=IFERROR(INDEX(Markets!$G${FIRST}:$G${MAXR},MATCH(B{rr},Markets!$A${FIRST}:$A${MAXR},0)),\"\")").font = F(10)
        for cc in range(2, 6):
            ws.cell(rr, cc).border = BOX
    end_c = r + 1 + len(countries)

    # by stage
    section(ws, r, 7, "Pipeline by stage", 4)
    for i, h in enumerate(["Stage", "Projects", "MW"]):
        c = ws.cell(r + 1, 7 + i, h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = CENTER
    stages = LY.LISTS["stages"]
    for j, s in enumerate(stages):
        rr = r + 2 + j
        ws.cell(rr, 7, s).font = F(10, True, NAVY)
        ws.cell(rr, 8, f'=COUNTIF({rng("stage")},G{rr})').font = F(10)
        ws.cell(rr, 9, f'=SUMIF({rng("stage")},G{rr},{rng("capacity_mw")})').font = F(10)
        ws.cell(rr, 9).number_format = "#,##0"
        for cc in range(7, 10):
            ws.cell(rr, cc).border = BOX
    # by service
    rs = r + 3 + len(stages)
    section(ws, rs, 7, "Projects by Ethos service", 4)
    services = [("M&E design", "*M&E*"), ("Energy / BESS / grid", "*nergy*"), ("Cooling / CFD", "*ooling*"),
                ("Civil & structural", "*ivil*"), ("Sustainability", "*ustainab*")]
    for j, (lab, pat) in enumerate(services):
        rr = rs + 1 + j
        ws.cell(rr, 7, lab).font = F(10, True, NAVY)
        ws.cell(rr, 8, f'=COUNTIF({rng("ethos_fit")},"{pat}")').font = F(10)
        for cc in (7, 8):
            ws.cell(rr, cc).border = BOX

    # charts
    ch = BarChart()
    ch.type = "bar"
    ch.style = 10
    ch.title = "MW tracked by market"
    ch.y_axis.title = None
    ch.add_data(Reference(ws, min_col=4, min_row=r + 1, max_row=end_c), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=r + 2, max_row=end_c))
    ch.legend = None
    ch.height, ch.width = 7.5, 12
    ch.series[0].graphicalProperties.solidFill = TEAL
    ch.series[0].graphicalProperties.line.solidFill = TEAL
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.dataLabels.showSerName = False
    ch.dataLabels.showCatName = False
    ch.dataLabels.showLegendKey = False
    style_chart(ch)
    ws.add_chart(ch, f"B{max(end_c, rs + 6) + 2}")

    ch2 = BarChart()
    ch2.style = 10
    ch2.title = "Projects by stage"
    ch2.add_data(Reference(ws, min_col=8, min_row=r + 1, max_row=r + 1 + len(stages)), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=7, min_row=r + 2, max_row=r + 1 + len(stages)))
    ch2.legend = None
    ch2.height, ch2.width = 7.5, 12
    ch2.series[0].graphicalProperties.solidFill = NAVY
    ch2.dataLabels = DataLabelList()
    ch2.dataLabels.showVal = True
    ch2.dataLabels.showSerName = False
    ch2.dataLabels.showCatName = False
    ch2.dataLabels.showLegendKey = False
    style_chart(ch2)
    ws.add_chart(ch2, f"G{max(end_c, rs + 6) + 2}")

    # top 10
    rt = max(end_c, rs + 6) + 19
    section(ws, rt, 2, "Top 10 opportunities (by priority score)", 10)
    hdr = ["#", "Project", "Developer", "Market", "Stage", "MW", "Score", "Next action"]
    spans = [(2, 2), (3, 4), (5, 5), (6, 6), (7, 7), (8, 8), (9, 9), (10, 11)]
    for (a, b), h in zip(spans, hdr):
        if b > a:
            ws.merge_cells(start_row=rt + 1, start_column=a, end_row=rt + 1, end_column=b)
        c = ws.cell(rt + 1, a, h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = CENTER
    key = f"{P}${L(sortcol)}${FIRST}:${L(sortcol)}${MAXR}"
    for n in range(1, 11):
        rr = rt + 1 + n
        m = f"MATCH(LARGE({key},{n}),{key},0)"
        pick = lambda k: f'=IFERROR(INDEX({rng(k)},{m}),"")'
        vals = [n, pick("project"), pick("developer"), pick("country"), pick("stage"),
                pick("capacity_mw"), pick("_priority"), pick("next_action")]
        for (a, b), v in zip(spans, vals):
            if b > a:
                ws.merge_cells(start_row=rr, start_column=a, end_row=rr, end_column=b)
            c = ws.cell(rr, a, v)
            c.font = F(10, a == 3, NAVY if a == 3 else INK)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            for cc in range(a, b + 1):
                ws.cell(rr, cc).border = BOX
        ws.row_dimensions[rr].height = 42
        ws.cell(rr, 2).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(rr, 8).number_format = "#,##0"

    # latest changes
    rc = rt + 14
    section(ws, rc, 2, "Latest changes (newest first)", 10)
    hdr2 = [("Date", 2, 2), ("Item", 3, 5), ("What changed", 6, 7), ("Why it matters", 8, 11)]
    for h, a, b in hdr2:
        ws.merge_cells(start_row=rc + 1, start_column=a, end_row=rc + 1, end_column=b)
        c = ws.cell(rc + 1, a, h)
        c.font = F(9, True, "FFFFFF")
        c.fill = fill(NAVY)
        c.alignment = CENTER
    for n in range(6):
        rr = rc + 2 + n
        src = FIRST + n
        vals = [(f'=IF(Changes!A{src}="","",Changes!A{src})', 2, 2),
                (f'=IF(Changes!D{src}="","",Changes!D{src})', 3, 5),
                (f'=IF(Changes!E{src}="","",Changes!E{src}&": "&Changes!F{src}&" → "&Changes!G{src})', 6, 7),
                (f'=IF(Changes!H{src}="","",Changes!H{src})', 8, 11)]
        for v, a, b in vals:
            ws.merge_cells(start_row=rr, start_column=a, end_row=rr, end_column=b)
            c = ws.cell(rr, a, v)
            c.font = F(10)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            for cc in range(a, b + 1):
                ws.cell(rr, cc).border = BOX
        ws.cell(rr, 2).number_format = "dd mmm yyyy"
        ws.row_dimensions[rr].height = 30

    # Spain line chart
    rl = rc + 9
    section(ws, rl, 2, "Spain: renewable share of generation vs the decree's 90% level", 10)
    lc = LineChart()
    lc.title = None
    lc.style = 12
    lc.add_data(Reference(wb["Live Data"], min_col=2, min_row=LY.HEADER_ROW, max_row=FIRST + 30), titles_from_data=True)
    lc.add_data(Reference(wb["Live Data"], min_col=7, min_row=LY.HEADER_ROW, max_row=FIRST + 30), titles_from_data=True)
    lc.set_categories(Reference(wb["Live Data"], min_col=1, min_row=FIRST, max_row=FIRST + 30))
    lc.height, lc.width = 7.5, 24
    lc.y_axis.scaling.min = 0
    lc.y_axis.scaling.max = 100
    lc.series[0].graphicalProperties.line.solidFill = TEAL
    lc.series[0].graphicalProperties.line.width = 28000
    lc.series[1].graphicalProperties.line.solidFill = "C0392B"
    lc.series[1].graphicalProperties.line.dashStyle = "dash"
    style_chart(lc)
    lc.legend.position = "b"
    for _s in lc.series:
        _s.smooth = False
    ws.add_chart(lc, f"B{rl + 1}")
    ws.print_area = f"A1:L{rl + 17}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = TEAL


def style_chart(ch):
    ch.roundedCorners = False
    if ch.legend is not None:
        ch.legend.overlay = False
    if ch.title is not None and hasattr(ch.title, "overlay"):
        ch.title.overlay = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.y_axis.majorGridlines = None


# ---------------------------------------------------------------- read me
def sheet_readme(wb, cfg, counts):
    ws = wb.create_sheet("Read me", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 95
    ws["B1"] = cfg["workbook"]["title"]
    ws["B1"].font = F(18, True, NAVY)
    ws["B2"] = ("Work sample for Ethos Engineering's Market Intelligence & Research Executive role. "
                f"Prepared by {cfg['workbook']['prepared_by']}. Public sources only.")
    ws["B2"].font = F(10, color=MUTED)
    ws.row_dimensions[1].height = 30
    blocks = [
        ("What this is", [
            ("Purpose", "One place to see where Ethos can win data centre work in Europe: named projects, the markets behind them, the rules that change them, and the accounts to call."),
            ("Who it is for", "Client Partners preparing meetings, Marketing planning campaigns, and leadership choosing which markets to prioritise."),
            ("How current", f"Rebuilt automatically every Monday. This copy was built {now()}."),
        ]),
        ("The tabs", [
            ("Dashboard", "Headline numbers, top 10 opportunities, latest changes, charts. Read this first."),
            ("Pipeline", "Every tracked project with stage, size, Ethos fit, priority score, sources and freshness."),
            ("Markets", "Why these countries: each scored 0-5 on pipeline, investment, grid access, policy and Ethos fit."),
            ("Grid & Policy", "Grid connection conditions and the latest rule change per market, with a Red/Amber/Green risk."),
            ("Accounts", "Developers and operators to target, linked to their projects, with the next step."),
            ("Sources", "Every source with publisher, date, primary/secondary label and whether the link still works."),
            ("Changes", "Dated log of what moved since the last refresh, and why it matters."),
            ("Gaps", "What we don't know yet and how to find out (often a question for the client)."),
            ("News", "Latest headlines per project, collected weekly via Google News (SerpApi)."),
            ("AI Scan", "Google AI Mode's latest status answer per project, with automatic flags where it disagrees with the tracker. Unverified until a person reviews it."),
            ("Live Data", "Spain's monthly renewable share from Red Eléctrica's public API, against the 90% level in Real Decreto 1032/2026."),
            ("Lists", "Dropdown values and the scoring rules. Change a number here and every score updates."),
        ]),
        ("How to read it", [
            ("Colours", "Navy headers = data you can edit. Teal headers and pale-teal cells = formulas, don't type over them. Green / amber / red = good / watch / act."),
            ("Freshness", f"'Days since check' turns amber at {cfg['freshness']['amber_days']} days and red at {cfg['freshness']['red_days']} days."),
            ("P / S", "P = primary source (company, regulator, grid operator, planning file). S = secondary (press, trackers)."),
            ("Confidence", "High = primary source confirms the figure. Medium = reputable secondary source. Low = single tracker or partial data; verify before quoting."),
            ("UNVERIFIED / blank", "Left blank rather than guessed. Every blank is a research task, listed on the Gaps tab where it matters."),
        ]),
        ("Automation", [
            ("What runs weekly", "1) keep any edits made in this workbook  2) re-open every source link  3) pull Red Eléctrica grid data  4) Google News headlines per project  5) Google AI Mode status scan per project  6) log what changed  7) rebuild this workbook."),
            ("Where it runs", "GitHub Actions (free, in the cloud). Nothing needs to run on your computer."),
            ("Code & run history", cfg["workbook"]["repo_url"]),
            ("Latest workbook", cfg["workbook"]["latest_download"]),
            ("AI guard-rail", "AI answers never overwrite the data. They are logged on the AI Scan tab and flagged for a person to check against a real source."),
            ("SerpApi credits", f"Capped at {cfg['serpapi']['monthly_budget']} searches a month (free plan = 250)."),
        ]),
        ("This build", [(k, str(v)) for k, v in counts.items()]),
    ]
    r = 4
    for head, items in blocks:
        ws.cell(r, 2, head.upper()).font = F(10, True, TEAL)
        ws.cell(r, 2).border = BOTTOM
        ws.cell(r, 3).border = BOTTOM
        r += 1
        for k, v in items:
            a = ws.cell(r, 2, k)
            a.font = F(10, True, NAVY)
            a.alignment = WRAP_TOP
            b = ws.cell(r, 3, v)
            b.font = F(10)
            b.alignment = WRAP_TOP
            if v.startswith("http"):
                b.hyperlink = v
                if v.endswith(".xlsx"):
                    b.value = "Download the latest version of this workbook (rebuilt every Monday)"
                b.font = F(10, color=TEAL, underline="single")
            ws.row_dimensions[r].height = est_height([(v, 95)])
            r += 1
        r += 1
    ws.sheet_properties.tabColor = NAVY


# ---------------------------------------------------------------- main
def run() -> str:
    cfg = load_config()
    T = {k: read_table(k) for k in ("projects", "accounts", "markets", "policy", "sources",
                                    "changes", "gaps", "news", "ai_scan", "grid_es")}
    T["changes"].sort(key=lambda r: r["date"], reverse=True)
    T["news"].sort(key=lambda r: r["date_found"], reverse=True)
    T["ai_scan"].sort(key=lambda r: (r["review_status"] != "Needs review", r["date"]), reverse=False)
    T["markets"].sort(key=lambda r: -sum(to_num(r[k]) or 0 for k in
                                         ("pipeline_score", "capex_score", "grid_score", "policy_score", "fit_score")))

    wb = Workbook()
    wb.remove(wb.active)
    wb.create_sheet("Dashboard")
    pipe_cols_count = None
    refs, sc = sheet_lists(wb)
    pipe_last, sortcol = sheet_pipeline(wb, T["projects"], refs, sc)

    pc = lambda k: col(LY.PIPELINE, k)
    prng = lambda k: f"Pipeline!${pc(k)}${FIRST}:${pc(k)}${MAXR}"
    mc = lambda k: col(LY.MARKETS, k)
    ws, last = sheet_simple(
        wb, "Markets", LY.MARKETS, T["markets"], "Market selection: where the deals are",
        "Each candidate market scored 0-5 on five tests; the top markets get the pipeline research. Scores are judgements backed by the evidence column.",
        "Pipeline = MW under construction + planned. Investment = 2025-26 announcements. Grid access = how fast large loads connect. Policy = how supportive the rules are. Ethos fit = match with Ethos services.",
        fx={"_total": lambda r: f"=SUM({mc('pipeline_score')}{r}:{mc('fit_score')}{r})",
            "_rank": lambda r: f"=RANK({mc('_total')}{r},${mc('_total')}${FIRST}:${mc('_total')}${FIRST + len(T['markets']) - 1})",
            "_projects": lambda r: f"=COUNTIF({prng('country')},A{r})",
            "_mw": lambda r: f"=SUMIF({prng('country')},A{r},{prng('capacity_mw')})"},
        tab="2E8B57", span=11)
    for r in range(FIRST, last + 1):
        ws[f"A{r}"].font = F(10, True, NAVY)
    ws.conditional_formatting.add(f"{mc('_total')}{FIRST}:{mc('_total')}{MAXR}",
        CellIsRule(operator="greaterThanOrEqual", formula=["16"], fill=fill(GREEN_F), font=F(10, True, GREEN_T)))

    ws, last = sheet_simple(wb, "Grid & Policy", LY.POLICY, T["policy"], "Grid & policy watchlist",
        "What it takes to connect a large data centre in each market, the latest rule change, and the next date to watch.",
        "Risk: Red = rules or grid make new projects hard now; Amber = possible with engineering effort; Green = supportive.",
        tab="C0392B", span=10, refs=refs)
    rag_rules(ws, col(LY.POLICY, "rag"))
    add_age_rules(ws, col(LY.POLICY, "last_checked"), last) if False else None

    ac = lambda k: col(LY.ACCOUNTS, k)
    ws, last = sheet_simple(wb, "Accounts", LY.ACCOUNTS, T["accounts"], "Target accounts",
        "Developers and operators with European pipelines that fit Ethos. 'Projects tracked' and 'MW tracked' count matching rows on the Pipeline tab.",
        fx={"_projects": lambda r: f'=COUNTIF({prng("developer")},"*"&{ac("account")}{r}&"*")',
            "_mw": lambda r: f'=SUMIF({prng("developer")},"*"&{ac("account")}{r}&"*",{prng("capacity_mw")})'},
        tab="6C5CE7", span=10, refs=refs)

    sc_ = lambda k: col(LY.SOURCES, k)
    ws, last = sheet_simple(wb, "Sources", LY.SOURCES, T["sources"], "Source register",
        "Every source behind the tracker. Click a link to open it. The weekly check re-opens each link and stamps the date.",
        "Link status: OK = opens; BLOCKED = site refuses robots, check by hand; BROKEN = gone, find a replacement.",
        fx={"_age": lambda r: f'=IF({sc_("last_checked")}{r}="","",TODAY()-{sc_("last_checked")}{r})'},
        tab="7F8C8D", span=8, refs=refs)
    add_age_rules(ws, sc_("_age"), last)
    ls = sc_("link_status")
    for txt, f_, t in (("BROKEN", RED_F, RED_T), ("BLOCKED", AMBER_F, AMBER_T), ("OK", GREEN_F, GREEN_T)):
        ws.conditional_formatting.add(f"{ls}{FIRST}:{ls}{MAXR}",
            CellIsRule(operator="equal", formula=[f'"{txt}"'], fill=fill(f_), font=F(10, True, t)))

    sheet_simple(wb, "Changes", LY.CHANGES, T["changes"], "Change log",
                 "What changed, when, and why it matters. 'auto' rows come from the weekly comparison; add your own as you learn things.",
                 tab="E67E22", span=9)
    ws, _ = sheet_simple(wb, "Gaps", LY.GAPS, T["gaps"], "Intelligence gaps",
                         "What we don't know yet, why it matters, and how to close it.", tab="E67E22", span=6, refs=refs)
    sheet_simple(wb, "News", LY.NEWS, T["news"], "News watch",
                 "Headlines found by the weekly Google News search (SerpApi) for each project. Newest first.",
                 tab="95A5A6", span=6)
    ws, last = sheet_simple(wb, "AI Scan", LY.AI_SCAN, T["ai_scan"], "Google AI Mode status scan",
                            "AI answers are leads, not facts. A flagged row means the answer disagrees with the Pipeline; confirm against a real source before changing anything.",
                            tab="95A5A6", span=6, refs=refs)
    rv = col(LY.AI_SCAN, "review_status")
    ws.conditional_formatting.add(f"{rv}{FIRST}:{rv}{MAXR}",
        CellIsRule(operator="equal", formula=['"Needs review"'], fill=fill(AMBER_F), font=F(10, True, AMBER_T)))

    gc = lambda k: col(LY.GRID_ES, k)
    ws, last = sheet_simple(wb, "Live Data", LY.GRID_ES, T["grid_es"], "Live data: Spain generation mix",
                            "Monthly renewable share of Spanish generation from Red Eléctrica's REData API (free, no key). Refreshed weekly.",
                            "Real Decreto 1032/2026 requires 80% hourly renewable matching for 10 MW+ data centres until renewables exceed 90% of the mix.",
                            fx={"_gap": lambda r: f"=IF({gc('renewable_share_pct')}{r}=\"\",\"\",{sc['threshold']}-{gc('renewable_share_pct')}{r})"},
                            tab="16A085", span=6)
    ws.cell(LY.HEADER_ROW, 7, "Decree level (%)").font = F(10, True, "FFFFFF")
    ws.cell(LY.HEADER_ROW, 7).fill = fill(TEAL)
    ws.column_dimensions["G"].width = 14
    for r in range(FIRST, max(last, FIRST + 30) + 1):
        if r <= last and T["grid_es"]:
            ws.cell(r, 7, f"={sc['threshold']}").font = F(10)

    countries = [m["country"] for m in T["markets"]][:8] or LY.LISTS["countries"][:7]
    sheet_dashboard(wb, pipe_last, sortcol, countries, cfg)
    counts = {
        "Projects": len(T["projects"]), "Accounts": len(T["accounts"]),
        "Markets scored": len(T["markets"]), "Sources": len(T["sources"]),
        "Change-log entries": len(T["changes"]), "News items": len(T["news"]),
        "AI scans": len(T["ai_scan"]), "Grid data months": len(T["grid_es"]),
    }
    sheet_readme(wb, cfg, counts)
    order = ["Read me", "Dashboard", "Pipeline", "Markets", "Grid & Policy", "Accounts", "Sources",
             "Changes", "Gaps", "News", "AI Scan", "Live Data", "Lists"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 1
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
        if ws.title not in ("Dashboard", "Read me"):
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.print_title_rows = f"{LY.HEADER_ROW}:{LY.HEADER_ROW}"
    OUTPUT.mkdir(exist_ok=True)
    path = OUTPUT / WORKBOOK_NAME
    for _ws in wb.worksheets:  # print/PDF: landscape, one page wide
        _ws.page_setup.orientation = 'landscape'
        _ws.page_setup.paperSize = _ws.PAPERSIZE_A4
        _ws.sheet_properties.pageSetUpPr.fitToPage = True
        _ws.page_setup.fitToWidth = 1
        _ws.page_setup.fitToHeight = 0
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)  # Excel recalculates on open
    wb.save(path)
    log("build_tracker", "ok", f"{path.name}: {counts}")
    return str(path)


if __name__ == "__main__":
    run()
