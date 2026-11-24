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
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from . import docs as DOCS
from . import layout as LY
from .common import OUTPUT, WORKBOOK_NAME, load_config, log, now, read_table

# ---------------------------------------------------------------- style kit
FONT = "Arial"
NAVY, TEAL, INK, MUTED = "0F2537", "00707C", "1A2530", "5B6B78"
PALE, TEAL_PALE, LINE = "F2F6F8", "E8F4F5", "D5DEE4"
GREEN_F, AMBER_F, RED_F = "DDF1E4", "FFF1D6", "FBE0DD"
GREEN_T, AMBER_T, RED_T = "1E6B3A", "8A5A00", "A12A1C"
AUTO_H, AUTO_F, BAND = "6C5CE7", "F3F0FD", "F7F9FA"
BLUE_F, BLUE_T = "E3EEFA", "1F4E79"

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


def tip(cell, text, w=260, h=110):
    if text:
        cm = Comment(text, "Tracker guide")
        cm.width, cm.height = w, h
        cell.comment = cm


def header_row(ws, cols):
    for i, (k, label, width, kind) in enumerate(cols, start=1):
        c = ws.cell(LY.HEADER_ROW, i, label)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(TEAL if kind == "formula" else AUTO_H if kind == "auto" else NAVY)
        tip(c, DOCS.TIPS.get(f"{ws.title}:{k}") or DOCS.TIPS.get(k, ""))
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
            if kind == "auto":
                c.fill = fill(AUTO_F)
                c.font = F(10, italic=True, color="3D3570")
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
    polish(ws, cols, last)
    # dropdowns
    for ci, (_k, _l, _w, kind) in enumerate(cols, start=1):
        if kind.startswith("list:"):
            ref = lists_ref[kind[5:]]
            dv = DataValidation(type="list", formula1=ref, allow_blank=True,
                                showErrorMessage=False)
            dv.add(f"{L(ci)}{FIRST}:{L(ci)}{MAXR}")
            ws.add_data_validation(dv)
    return last


def chips(ws, c, mapping):
    rng = f"{c}{FIRST}:{c}{MAXR}"
    for txt, (f_, t) in mapping.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{txt}"'],
                                                      fill=fill(f_), font=F(10, True, t)))


STAGE_CHIPS = {"Announced": (BLUE_F, BLUE_T), "Permitting": (GREEN_F, GREEN_T), "Permitted": (GREEN_F, GREEN_T),
               "Under construction": (AMBER_F, AMBER_T), "Operational": (PALE, MUTED),
               "Cancelled / paused": (RED_F, RED_T)}


def polish(ws, cols, last):
    """Conditional formatting by column meaning, then zebra banding on data columns."""
    keys = [k for k, *_ in cols]
    kinds = {k: kind for k, _l, _w, kind in cols}
    C = lambda k: col(cols, k)
    end = max(last, FIRST + 30)
    if "stage" in keys:
        chips(ws, C("stage"), STAGE_CHIPS)
    if "ethos_relationship" in keys:
        chips(ws, C("ethos_relationship"), {"Existing client": (GREEN_F, GREEN_T), "Target": (BLUE_F, BLUE_T),
                                             "Competitor-held": (PALE, MUTED)})
    if "confidence" in keys:
        chips(ws, C("confidence"), {"High": (GREEN_F, GREEN_T), "Medium": (AMBER_F, AMBER_T), "Low": (RED_F, RED_T)})
    if "status" in keys and ws.title == "Gaps":
        chips(ws, C("status"), {"Open": (AMBER_F, AMBER_T), "In progress": (BLUE_F, BLUE_T), "Closed": (GREEN_F, GREEN_T)})
    if "type" in keys and ws.title == "Sources":
        chips(ws, C("type"), {"P": (GREEN_F, GREEN_T), "S": (PALE, MUTED)})
    if "design_team" in keys:
        c = C("design_team")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{MAXR}", FormulaRule(
            formula=[f'LEFT({c}{FIRST},3)="Not"'], font=F(10, True, BLUE_T), fill=fill(BLUE_F)))
    if "capacity_mw" in keys:
        c = C("capacity_mw")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{last}", DataBarRule(
            start_type="num", start_value=0, end_type="max", color="7FB8BE", showValue=True))
    if "_priority" in keys:
        c = C("_priority")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{last}", ColorScaleRule(
            start_type="num", start_value=3, start_color="FFFFFF", mid_type="num", mid_value=7,
            mid_color="FFE7A8", end_type="num", end_value=12, end_color="63BE7B"))
    for k in ("pipeline_score", "capex_score", "grid_score", "policy_score", "fit_score"):
        if k in keys:
            c = C(k)
            ws.conditional_formatting.add(f"{c}{FIRST}:{c}{last}", ColorScaleRule(
                start_type="num", start_value=1, start_color="F8696B", mid_type="num", mid_value=3,
                mid_color="FFEB84", end_type="num", end_value=5, end_color="63BE7B"))
    if "_total" in keys:
        c = C("_total")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{last}", DataBarRule(
            start_type="num", start_value=0, end_type="num", end_value=25, color="00707C", showValue=True))
    if "milestone_date" in keys:
        c = C("milestone_date")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{MAXR}", FormulaRule(
            formula=[f"AND(ISNUMBER({c}{FIRST}),{c}{FIRST}>=TODAY(),{c}{FIRST}-TODAY()<=60)"],
            fill=fill(AMBER_F), font=F(10, True, AMBER_T)))
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{MAXR}", FormulaRule(
            formula=[f"AND(ISNUMBER({c}{FIRST}),{c}{FIRST}<TODAY())"], font=F(10, color=MUTED, italic=True)))
    if "deadline" in keys:
        c = C("deadline")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{MAXR}", FormulaRule(
            formula=[f"AND(ISNUMBER({c}{FIRST}),{c}{FIRST}>=TODAY(),{c}{FIRST}-TODAY()<=21)"],
            fill=fill(RED_F), font=F(10, True, RED_T)))
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{MAXR}", FormulaRule(
            formula=[f"AND(ISNUMBER({c}{FIRST}),{c}{FIRST}<TODAY())"], font=F(10, color=MUTED, italic=True)))
    if "_gap" in keys:
        c = C("_gap")
        ws.conditional_formatting.add(f"{c}{FIRST}:{c}{last}", ColorScaleRule(
            start_type="num", start_value=0, start_color="63BE7B", end_type="max", end_color="F8696B"))
    if ws.title == "Changes" and "field" in keys:
        c = C("field")
        ws.conditional_formatting.add(f"A{FIRST}:{L(len(cols))}{MAXR}", FormulaRule(
            formula=[f'OR(${c}{FIRST}="stage",${c}{FIRST}="rag",${c}{FIRST}="design_team")'],
            fill=fill(AMBER_F)))
    # zebra banding on plain data columns only (formula/auto columns keep their tint)
    blocks, start = [], None
    for i, k in enumerate(keys + [None], start=1):
        plain = k is not None and kinds[k] not in ("formula", "auto")
        if plain and start is None:
            start = i
        if not plain and start is not None:
            blocks.append((start, i - 1)); start = None
    ws._band = (blocks, end)   # added last (lowest priority) in add_banding()


def add_banding(wb):
    for ws in wb.worksheets:
        blocks, end = getattr(ws, "_band", ([], 0))
        for a, b in blocks:
            ws.conditional_formatting.add(f"{L(a)}{FIRST}:{L(b)}{end}", FormulaRule(
                formula=[f'AND(MOD(ROW(),2)=0,$A{FIRST}<>"")'], fill=fill(BAND)))


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
TILE_TIPS = {
    "PROJECTS TRACKED": "Count of rows on the Pipeline tab.",
    "MW TRACKED": "Sum of 'Capacity (MW)' on the Pipeline tab. Mixes IT and grid MW where sources differ; see the MW basis column.",
    "IN THE CONSULTANT WINDOW": "Projects at Announced, Permitting or Permitted: the stages when design consultants are usually appointed.",
    "WITH EXISTING CLIENTS": "Projects whose developer already works with Ethos.",
    "SPAIN RENEWABLE SHARE": "Latest monthly renewable share of Spanish generation (Red Eléctrica API). Spain's decree requirement eases once this passes 90%.",
    "HIGH-PRIORITY PROJECTS": "Projects scoring 9 or more out of 12. See the Scoring guide tab.",
    "SOURCES DUE A RE-CHECK": "Sources not confirmed for 90+ days.",
    "BROKEN / BLOCKED LINKS": "Sources the weekly link check could not open. Blocked sites usually open fine in a browser.",
    "AI FLAGS TO REVIEW": "Google AI Mode answers that disagree with the Pipeline and need a person to check.",
    "OPEN INTELLIGENCE GAPS": "Open questions on the Gaps tab.",
    "PUBLIC TENDERS (90 DAYS)": "Public notices mentioning data centres found in EU TED and UK Find a Tender in the last 90 days.",
    "HEADLINES (14 DAYS)": "New headlines found by the free Google News watch in the last 14 days.",
    "DESIGN TEAM NOT PUBLIC": "Projects with no published design team: openings to ask clients about.",
    "MW WITH EXISTING CLIENTS": "Total MW of projects whose developer is already an Ethos client.",
    "NEXT POLICY MILESTONE": "Earliest upcoming date on the Grid & Policy tab.",
}


def tile(ws, r, c, label, formula, fmt="#,##0", note=""):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + 1)
    ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + 1)
    ws.merge_cells(start_row=r + 2, start_column=c, end_row=r + 2, end_column=c + 1)
    for rr in (r, r + 1, r + 2):
        for cc in (c, c + 1):
            ws.cell(rr, cc).fill = fill(PALE)
    top = ws.cell(r, c, label)
    top.font = F(9, True, MUTED)
    tip(top, TILE_TIPS.get(label, ""))
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
                                           "_tier", "_priority", "developer", "next_action", "ethos_fit",
                                           "suggested_action")}
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
    pd = f"Pipeline!${col(LY.PIPELINE, 'design_team')}${FIRST}:${col(LY.PIPELINE, 'design_team')}${MAXR}"
    ms = f"'Grid & Policy'!${col(LY.POLICY, 'milestone_date')}${FIRST}:${col(LY.POLICY, 'milestone_date')}${MAXR}"
    tile(ws, 15, 2, "PUBLIC TENDERS (90 DAYS)", f'=COUNTIF(Tenders!$A${FIRST}:$A${MAXR},">="&(TODAY()-90))',
         note="EU TED + UK Find a Tender")
    tile(ws, 15, 4, "HEADLINES (14 DAYS)", f'=COUNTIF(News!$A${FIRST}:$A${MAXR},">="&(TODAY()-14))',
         note="Free Google News watch")
    tile(ws, 15, 6, "DESIGN TEAM NOT PUBLIC", f'=COUNTIF({pd},"Not*")', note="Openings to ask clients about")
    tile(ws, 15, 8, "MW WITH EXISTING CLIENTS",
         f'=SUMIF({rng("ethos_relationship")},"Existing client",{rng("capacity_mw")})', note="Warm-door capacity")
    tile(ws, 15, 10, "NEXT POLICY MILESTONE",
         f'=IFERROR(SMALL({ms},COUNTIF({ms},"<"&TODAY())+1),"none set")', fmt="dd mmm yyyy",
         note="Earliest upcoming date, Grid & Policy tab")

    # by country
    r = 20
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
    # by relationship (feeds a chart)
    rrel = r + 3 + len(stages) + len(services) + 1
    section(ws, rrel, 2, "MW by Ethos relationship", 4)
    for i, h in enumerate(["Relationship", "Projects", "MW"]):
        c = ws.cell(rrel + 1, 2 + i, h)
        c.font = F(9, True, "FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER
    for j, rel in enumerate(LY.LISTS["relationship"]):
        rr = rrel + 2 + j
        ws.cell(rr, 2, rel).font = F(10, True, NAVY)
        ws.cell(rr, 3, f'=COUNTIF({rng("ethos_relationship")},B{rr})').font = F(10)
        ws.cell(rr, 4, f'=SUMIF({rng("ethos_relationship")},B{rr},{rng("capacity_mw")})').font = F(10)
        ws.cell(rr, 4).number_format = "#,##0"
        for cc in range(2, 5):
            ws.cell(rr, cc).border = BOX
    end_c = max(end_c, rrel + 1 + len(LY.LISTS["relationship"]))

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
                pick("capacity_mw"), pick("_priority"),
                f'=IFERROR(IF(INDEX({rng("next_action")},{m})<>"",INDEX({rng("next_action")},{m}),"Suggested: "&INDEX({rng("suggested_action")},{m})),"")']
        for (a, b), v in zip(spans, vals):
            if b > a:
                ws.merge_cells(start_row=rr, start_column=a, end_row=rr, end_column=b)
            c = ws.cell(rr, a, v)
            c.font = F(10, a == 3, NAVY if a == 3 else INK)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            for cc in range(a, b + 1):
                ws.cell(rr, cc).border = BOX
        ws.row_dimensions[rr].height = 58
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

    rm = rl + 18
    section(ws, rm, 2, "Why these markets: score breakdown (1-5 per test) and MW by relationship", 10)
    mk = wb["Markets"]
    n_m = len(countries)
    sb = BarChart()
    sb.type, sb.grouping, sb.overlap = "bar", "stacked", 100
    sb.style = 10
    sb.title = "Market score breakdown (out of 25)"
    sb.add_data(Reference(mk, min_col=2, max_col=6, min_row=LY.HEADER_ROW, max_row=LY.HEADER_ROW + n_m),
                titles_from_data=True)
    sb.set_categories(Reference(mk, min_col=1, min_row=FIRST, max_row=FIRST + n_m - 1))
    for sr, colr in zip(sb.series, ("0F2537", "00707C", "7FB8BE", "E3A33B", "6C5CE7")):
        sr.graphicalProperties.solidFill = colr
        sr.graphicalProperties.line.solidFill = colr
    sb.x_axis.scaling.orientation = "maxMin"
    sb.y_axis.scaling.min, sb.y_axis.scaling.max = 0, 25
    sb.height, sb.width = 8.5, 14
    style_chart(sb)
    sb.legend.position = "b"
    ws.add_chart(sb, f"B{rm + 1}")
    rc3 = BarChart()
    rc3.style = 10
    rc3.title = "MW by Ethos relationship"
    rc3.add_data(Reference(ws, min_col=4, min_row=rrel + 1, max_row=rrel + 1 + len(LY.LISTS["relationship"])),
                 titles_from_data=True)
    rc3.set_categories(Reference(ws, min_col=2, min_row=rrel + 2, max_row=rrel + 1 + len(LY.LISTS["relationship"])))
    rc3.legend = None
    rc3.height, rc3.width = 8.5, 10
    rc3.series[0].graphicalProperties.solidFill = TEAL
    rc3.dataLabels = DataLabelList(); rc3.dataLabels.showVal = True
    rc3.dataLabels.showSerName = rc3.dataLabels.showCatName = rc3.dataLabels.showLegendKey = False
    style_chart(rc3)
    ws.add_chart(rc3, f"H{rm + 1}")
    ws.print_area = f"A1:L{rm + 19}"
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


# ---------------------------------------------------------------- guides
def _head(ws, r, labels, widths_from=2):
    for i, h in enumerate(labels):
        c = ws.cell(r, widths_from + i, h)
        c.font = F(10, True, "FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER; c.border = BOX


def _row(ws, r, vals, start=2, bold_first=True, h=None):
    for i, v in enumerate(vals):
        c = ws.cell(r, start + i, v)
        c.font = F(10, bold_first and i == 0, NAVY if bold_first and i == 0 else INK)
        c.alignment = WRAP_TOP; c.border = BOX
    ws.row_dimensions[r].height = h or est_height([(str(v), 34) for v in vals])


def sheet_scoring(wb):
    ws = wb.create_sheet("Scoring guide")
    ws.sheet_view.showGridLines = False
    for c, w in zip("ABCDEFG", (2, 26, 34, 34, 34, 34, 34)):
        ws.column_dimensions[c].width = w
    ws["B1"] = "How the scores work"; ws["B1"].font = F(16, True, NAVY)
    ws["B2"] = ("Two scores drive the tracker: a Priority score for each project (what to chase now) and a Market score "
                "for each country (where to look). Both are transparent: every point is explained below.")
    ws["B2"].font = F(10, color=MUTED); ws.merge_cells("B2:G2"); ws.row_dimensions[2].height = 28
    ws["B2"].alignment = WRAP_TOP
    r = 4
    section(ws, r, 2, "1. Priority score (per project, 0 to 12)", 6); r += 1
    ws.cell(r, 2, "Priority = stage points + relationship points + confidence points + size points. "
            "Points are live: edit them on the Lists tab and every score updates.").font = F(10)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7); r += 2
    _head(ws, r, ["Part", "Value", "Points", "Why"]); r += 1
    why = {"Announced": "Design team usually not chosen yet: the best time to get in.",
           "Permitting": "Design is being developed for the planning application: consultants are being appointed now.",
           "Permitted": "Consent won; detailed design and contractor appointments follow.",
           "Under construction": "Design largely fixed; only fit-out, commissioning or later phases remain.",
           "Operational": "Built. Only expansions or retrofits.",
           "Cancelled / paused": "No work to win.",
           "Existing client": "A warm door: Ethos already knows the buyer.",
           "Target": "Reachable, but needs an introduction.",
           "Competitor-held": "Design held by another consultant; only specialist or later-phase work.",
           "High": "Figures confirmed by a primary source.", "Medium": "Reputable secondary source.",
           "Low": "Partial or single-tracker data."}
    lists = wb["Lists"]
    pts_cells = {}
    for rr in range(1, 60):
        for cc in range(1, 12):
            v = lists.cell(rr, cc).value
            if isinstance(v, str) and v in why and lists.cell(rr, cc + 1).value is not None \
                    and isinstance(lists.cell(rr, cc + 1).value, (int, float)):
                pts_cells[v] = f"=Lists!{L(cc + 1)}{rr}"
    for part, d in (("Stage", LY.STAGE_POINTS), ("Ethos relationship", LY.REL_POINTS), ("Confidence", LY.CONF_POINTS)):
        for k in d:
            _row(ws, r, [part, k, pts_cells.get(k, d[k]), why.get(k, "")], h=30); r += 1
    for k, v, w in (("200 MW or more", 3, "Large campus: multi-year design fees."),
                    ("50 to 199 MW", 2, "Mid-size project."), ("Under 50 MW", 1, "Small or single building."),
                    ("MW unknown", 0, "Not scored until confirmed (see Gaps).")):
        _row(ws, r, ["Size", k, v, w], h=30); r += 1
    r += 1
    ws.cell(r, 2, "Priority band:").font = F(10, True, NAVY)
    ws.cell(r, 3, "High = 9 or more   ·   Medium = 6 to 8   ·   Low = under 6").font = F(10); r += 2

    section(ws, r, 2, "Worked example: the top project on the Pipeline tab", 6); r += 1
    pc = lambda k: col(LY.PIPELINE, k)
    ex = [("Project", f"=Pipeline!{pc('project')}{FIRST}"),
          ("Stage", f"=Pipeline!{pc('stage')}{FIRST}"),
          ("Ethos relationship", f"=Pipeline!{pc('ethos_relationship')}{FIRST}"),
          ("Confidence", f"=Pipeline!{pc('confidence')}{FIRST}"),
          ("Capacity (MW)", f"=Pipeline!{pc('capacity_mw')}{FIRST}"),
          ("Priority score", f"=Pipeline!{pc('_priority')}{FIRST}")]
    for k, f_ in ex:
        _row(ws, r, [k, f_], h=18); r += 1
    r += 1

    section(ws, r, 2, "2. Market score (per country, 1 to 5 on five tests, total out of 25)", 6); r += 1
    ws.cell(r, 2, "Each market is scored against the descriptions below, using the evidence and sources on the Markets tab. "
            "The five tests are weighted equally; a different weighting is a one-line change.").font = F(10)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7); r += 2
    _head(ws, r, ["Score", "5", "4", "3", "2", "1"][:1] + [t for t, *_ in DOCS.MARKET_RUBRIC]); r += 1
    for lvl in range(5):
        _row(ws, r, [str(5 - lvl)] + [lv[lvl] for _t, _k, lv in DOCS.MARKET_RUBRIC], h=46); r += 1
    r += 1
    section(ws, r, 2, "3. Other ratings", 6); r += 1
    for k, v in (("Risk (Grid & Policy)", "Red = rules or grid make new projects hard now. Amber = possible with engineering effort. Green = supportive."),
                 ("Confidence", "High = primary source confirms key figures. Medium = reputable secondary source. Low = partial or single-tracker data."),
                 ("Freshness", "Days since last check: amber at 90, red at 180 (set in config.yaml)."),
                 ("AI Scan flags", "An AI answer that mentions a different stage, MW (±10%) or target year. Always checked by a person.")):
        _row(ws, r, [k, v], h=32)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7); r += 1
    ws.sheet_properties.tabColor = "9AA7B1"


def sheet_glossary(wb):
    ws = wb.create_sheet("Glossary")
    ws.sheet_view.showGridLines = False
    for c, w in zip("ABC", (2, 32, 100)):
        ws.column_dimensions[c].width = w
    ws["B1"] = "Glossary"; ws["B1"].font = F(16, True, NAVY)
    ws["B2"] = "Plain-English meanings of the terms used in this workbook. Hover over any column header for its own explanation."
    ws["B2"].font = F(10, color=MUTED)
    r = 4
    for group, items in DOCS.GLOSSARY:
        section(ws, r, 2, group, 2); r += 1
        for term, meaning in items:
            _row(ws, r, [term, meaning], h=est_height([(meaning, 100)])); r += 1
        r += 1
    r += 1
    section(ws, r, 2, "Free tools and data behind the automation", 2); r += 1
    for name, what, status in DOCS.FREE_TOOLS:
        _row(ws, r, [name, f"{what}  [{status}]"], h=18); r += 1
    ws.sheet_properties.tabColor = "9AA7B1"


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
            ("News", "Latest headlines per project, collected weekly from Google News (free RSS feed)."),
            ("Tenders", "Public tenders mentioning data centres in the target markets (EU TED and UK Find a Tender)."),
            ("AI Scan", "Google AI Mode's latest status answer per project, with automatic flags where it disagrees with the tracker. Unverified until a person reviews it."),
            ("Live Data", "Spain's monthly renewable share from Red Eléctrica's public API, against the 90% level in Real Decreto 1032/2026."),
            ("Scoring guide", "How the Priority score and the Market score are worked out, point by point, with a live worked example."),
            ("Glossary", "Plain-English meaning of every technical term and abbreviation, plus the free tools behind the automation."),
            ("Lists", "Dropdown values and the scoring points. Change a number here and every score updates."),
        ]),
        ("How to use it", [
            ("Hover for help", "Every column header has a note (small red corner). Hover over it to see what the column means and how it is filled."),
            ("Sort & filter", "Click the arrow in any header to sort (e.g. Priority score, largest first) or filter (e.g. Country = Spain). Pipeline is already sorted by priority."),
            ("Colours", "Navy headers = data you can edit. Teal headers = formulas, don't type over them. Purple headers = written automatically each week. Green / amber / red = good / watch / act."),
            ("Next actions", "'Next action (agreed)' is yours. 'Suggested next action (auto)' is rebuilt weekly from stage, relationship, news, AI flags and policy (and improved by Gemini when enabled). Copy a suggestion across when you agree with it."),
        ]),
        ("How to read it", [
            ("Freshness", f"'Days since check' turns amber at {cfg['freshness']['amber_days']} days and red at {cfg['freshness']['red_days']} days."),
            ("P / S", "P = primary source (company, regulator, grid operator, planning file). S = secondary (press, trackers)."),
            ("Confidence", "High = primary source confirms the figure. Medium = reputable secondary source. Low = single tracker or partial data; verify before quoting."),
            ("UNVERIFIED / blank", "Left blank rather than guessed. Every blank is a research task, listed on the Gaps tab where it matters."),
        ]),
        ("Automation", [
            ("What runs weekly", "1) keep edits made in this workbook  2) re-open every source link  3) pull Red Eléctrica grid data  4) Google News headlines per project (free RSS)  5) public tenders from EU TED and UK Find a Tender  6) Google AI Mode status check per project  7) suggested next action per project (rules + Gemini)  8) log what changed  9) rebuild this workbook."),
            ("Where it runs", "GitHub Actions (free, in the cloud). Nothing needs to run on your computer."),
            ("Code & run history", cfg["workbook"]["repo_url"]),
            ("Latest workbook", cfg["workbook"]["latest_download"]),
            ("AI guard-rail", "AI answers never overwrite the data. They are logged on the AI Scan tab and flagged for a person to check against a real source."),
            ("SerpApi credits", f"Only the AI Mode check uses SerpApi; capped at {cfg['serpapi']['monthly_budget']} searches a month (free plan = 250)."),
            ("Cost", "All services used are free tiers or free public APIs. The full list is on the Glossary tab."),
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
    def prio(p):
        mw = to_num(p["capacity_mw"]) or 0
        return (LY.STAGE_POINTS.get(p["stage"], 0) + LY.REL_POINTS.get(p["ethos_relationship"], 0)
                + LY.CONF_POINTS.get(p["confidence"], 0) + (3 if mw >= 200 else 2 if mw >= 50 else 1 if mw > 0 else 0))
    T["projects"].sort(key=lambda p: (-prio(p), -(to_num(p["capacity_mw"]) or 0)))
    T["accounts"].sort(key=lambda a: (a["ethos_relationship"] != "Existing client", a["account"]))
    T["policy"].sort(key=lambda r: ({"Red": 0, "Amber": 1, "Green": 2}.get(r["rag"], 3), r["country"]))
    T["gaps"].sort(key=lambda g: ({"Open": 0, "In progress": 1, "Closed": 2}.get(g["status"], 3), g["id"]))
    T["tenders"] = read_table("tenders")
    T["tenders"].sort(key=lambda t: t["published"], reverse=True)
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
    for txt, f_, t in (("BROKEN", RED_F, RED_T), ("BLOCKED", AMBER_F, AMBER_T), ("ERROR", AMBER_F, AMBER_T), ("OK", GREEN_F, GREEN_T), ("OK (reader)", GREEN_F, GREEN_T)):
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

    ws, last = sheet_simple(wb, "Tenders", LY.TENDERS, T["tenders"], "Public tender watch",
        "Public notices mentioning data centres in the target markets, found weekly in the EU's TED and the UK's Find a Tender (both free, official). Newest first.",
        "Deadline turns red when it is 21 days away or less, grey once passed. Most private campuses never go to public tender; these show public-sector and enabling work.",
        tab="2C3E50", span=8)

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

    sheet_scoring(wb)
    sheet_glossary(wb)
    countries = [m["country"] for m in T["markets"]][:8] or LY.LISTS["countries"][:7]
    sheet_dashboard(wb, pipe_last, sortcol, countries, cfg)
    counts = {
        "Projects": len(T["projects"]), "Accounts": len(T["accounts"]),
        "Markets scored": len(T["markets"]), "Sources": len(T["sources"]),
        "Change-log entries": len(T["changes"]), "News items": len(T["news"]),
        "AI scans": len(T["ai_scan"]), "Grid data months": len(T["grid_es"]),
        "Public tenders": len(T["tenders"]),
    }
    sheet_readme(wb, cfg, counts)
    order = ["Read me", "Dashboard", "Pipeline", "Markets", "Grid & Policy", "Accounts", "Sources",
             "Changes", "Gaps", "News", "Tenders", "AI Scan", "Live Data", "Scoring guide", "Glossary", "Lists"]
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
    add_banding(wb)
    from openpyxl.workbook.properties import CalcProperties
    wb.calculation = CalcProperties(fullCalcOnLoad=True)  # Excel recalculates on open
    wb.save(path)
    log("build_tracker", "ok", f"{path.name}: {counts}")
    return str(path)


if __name__ == "__main__":
    run()
