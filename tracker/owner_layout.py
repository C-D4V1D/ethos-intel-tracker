"""Keep the owner's layout changes made in the OneDrive copy.

capture(): compares the owner's OneDrive copy with the exact file that was uploaded
(the baseline) and records what the owner changed: column widths, hidden columns,
row heights, frozen panes, zoom, cell formatting, and formula/text edits outside the
data tables. Saved in data/owner_layout.json.

apply(): re-applies those changes to every newly built workbook, so the owner never
loses them. Data tables are matched by column name and row ID (not cell address),
so the changes survive new rows and reordered columns. A text/formula edit is only
re-applied while the tracker still writes the same original there.
"""
from __future__ import annotations

import copy
import json

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .common import DATA, log
from .layout import HEADER_ROW

STORE = "owner_layout.json"
DATA_SHEETS = {"Pipeline", "Accounts", "Markets", "Grid & Policy", "Sources", "Changes", "Gaps", "News",
               "Tenders", "Hiring signals", "Events", "Review", "AI Scan", "Live Data", "Grid by market"}
FIRST = HEADER_ROW + 1


def _path():
    from . import common
    return common.DATA / STORE


def load() -> dict:
    p = _path()
    return json.loads(p.read_text()) if p.exists() else {}


def _rgb(c):
    try:
        v = c.rgb
        return v[-6:] if isinstance(v, str) else None
    except AttributeError:
        return None


def _style(c) -> dict:
    f, fl, a = c.font, c.fill, c.alignment
    return {"b": f.b, "i": f.i, "sz": f.sz, "color": _rgb(f.color) if f.color else None, "name": f.name,
            "fill": _rgb(fl.fgColor) if fl and fl.fill_type == "solid" else None,
            "h": a.horizontal, "v": a.vertical, "wrap": a.wrap_text, "fmt": c.number_format}


def _widths(ws) -> dict:
    out = {}
    for dim in ws.column_dimensions.values():
        if dim.width is None:
            continue
        lo, hi = dim.min or 0, dim.max or 0
        if not lo:
            continue
        for i in range(lo, min(hi, 200) + 1):
            out[get_column_letter(i)] = (round(dim.width, 2), bool(dim.hidden))
    return out


def _headers(ws) -> dict:
    return {c.column_letter: str(c.value).strip() for c in ws[HEADER_ROW] if c.value not in (None, "")}


def _norm(v) -> str:
    import datetime as _dt
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    s = "" if v is None else str(v).strip()
    return s[:10] if len(s) >= 19 and s[4] == "-" and s[10] == " " else s


def _rowkey(ws, r) -> str:
    return f"{_norm(ws.cell(r, 1).value)}|{_norm(ws.cell(r, 2).value)}"


def _diff_style(a: dict, b: dict) -> dict:
    return {k: b[k] for k in b if a.get(k) != b[k]}


def capture(baseline_path, owner_path) -> int:
    """Record what the owner changed (owner file vs the file we uploaded)."""
    base, own = load_workbook(baseline_path), load_workbook(owner_path)
    store = load()
    n = 0
    for name in own.sheetnames:
        if name not in base.sheetnames:
            continue
        a, b = base[name], own[name]
        s = store.setdefault(name, {})
        data = name in DATA_SHEETS
        hdr = _headers(a) if data else {}
        # sheet settings
        for key, va, vb in (("freeze", a.freeze_panes, b.freeze_panes),
                            ("zoom", a.sheet_view.zoomScale, b.sheet_view.zoomScale)):
            if va != vb:
                s[key] = vb
                n += 1
        # columns
        wa, wb_ = _widths(a), _widths(b)
        for col in set(wa) | set(wb_):
            va, vb = wa.get(col, (None, False)), wb_.get(col, (None, False))
            if vb[0] is None:
                continue
            if va[0] is None or abs((va[0] or 0) - vb[0]) > 0.6 or va[1] != vb[1]:
                k = hdr.get(col, col) if data else col
                s.setdefault("cols", {})[k] = {"w": vb[0], "hidden": vb[1]}
                n += 1
        # rows
        for r, dim in b.row_dimensions.items():
            hb, ha = dim.height, a.row_dimensions[r].height if r in a.row_dimensions else None
            if hb is None or (ha is not None and abs(ha - hb) <= 1):
                continue
            k = _rowkey(b, r) if data and r >= FIRST else str(r)
            s.setdefault("rows", {})[k] = hb
            n += 1
        # cells
        colstyle = {}
        for row in b.iter_rows(min_row=1, max_row=b.max_row, max_col=b.max_column):
            for c in row:
                ca = a[c.coordinate]
                d = _diff_style(_style(ca), _style(c))
                if d:
                    if data and c.row >= FIRST and c.column_letter in hdr:
                        colstyle.setdefault(hdr[c.column_letter], []).append(json.dumps(d, sort_keys=True))
                    else:
                        s.setdefault("cells", {})[c.coordinate] = {**s.get("cells", {}).get(c.coordinate, {}), **d}
                        n += 1
                if not data and str(ca.value) != str(c.value) and c.value is not None:
                    s.setdefault("values", {})[c.coordinate] = {"from": str(ca.value), "to": c.value}
                    n += 1
        rows_n = max(1, a.max_row - HEADER_ROW)
        for label, diffs in colstyle.items():
            top = max(set(diffs), key=diffs.count)
            if diffs.count(top) >= max(1, rows_n // 2):  # changed for (most of) the column
                s.setdefault("colstyle", {})[label] = {**s.get("colstyle", {}).get(label, {}), **json.loads(top)}
                n += 1
        if not s:
            store.pop(name)
    _path().write_text(json.dumps(store, indent=1, ensure_ascii=False))
    return n


def _restyle(c, d: dict) -> None:
    if any(k in d for k in ("b", "i", "sz", "color", "name")):
        f = copy.copy(c.font)
        c.font = Font(name=d.get("name", f.name), sz=d.get("sz", f.sz), b=d.get("b", f.b), i=d.get("i", f.i),
                      color=d.get("color", f.color.rgb if f.color and isinstance(f.color.rgb, str) else None),
                      u=f.u, strike=f.strike, vertAlign=f.vertAlign)
    if "fill" in d:
        c.fill = PatternFill("solid", fgColor=d["fill"]) if d["fill"] else PatternFill(fill_type=None)
    if any(k in d for k in ("h", "v", "wrap")):
        a = c.alignment
        c.alignment = Alignment(horizontal=d.get("h", a.horizontal), vertical=d.get("v", a.vertical),
                                wrap_text=d.get("wrap", a.wrap_text), indent=a.indent, text_rotation=a.text_rotation)
    if "fmt" in d:
        c.number_format = d["fmt"]


def _autofit_rows(ws) -> int:
    """Approximate Excel's 'AutoFit row height' for wrapped text, using the current column widths."""
    import math
    widths = {}
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        widths[col] = ws.column_dimensions[letter].width or 13
    n = 0
    for r in range(FIRST, ws.max_row + 1):
        lines = 1
        for c in ws[r]:
            if c.value is None or (isinstance(c.value, str) and c.value.startswith("=")):
                continue
            text = str(c.value)
            chars = max(1.0, widths.get(c.column, 13) * 1.15)
            need = sum(max(1, math.ceil(len(part) / chars)) for part in text.split("\n"))
            lines = max(lines, need)
        ws.row_dimensions[r].height = min(409, round(lines * 12 + 9.75, 2))  # Excel: 2 lines = 33.75
        n += 1
    return n


def apply(wb) -> int:
    store = load()
    n = 0
    for name, s in store.items():
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        data = name in DATA_SHEETS
        letter = {v: k for k, v in _headers(ws).items()} if data else {}
        if "freeze" in s:
            ws.freeze_panes = s["freeze"]
        if s.get("zoom"):
            ws.sheet_view.zoomScale = s["zoom"]
        for k, v in s.get("cols", {}).items():
            col = letter.get(k, k if not data or len(k) <= 3 and k.isupper() else None)
            if col:
                ws.column_dimensions[col].width = v["w"]
                ws.column_dimensions[col].hidden = v.get("hidden", False)
                n += 1
        rows = s.get("rows", {})
        if rows:
            if data and any("|" in k for k in rows):
                n += _autofit_rows(ws)  # the owner auto-fitted rows: do the same for today's content
            for k, h in rows.items():
                if k.isdigit() and (not data or int(k) < FIRST):
                    ws.row_dimensions[int(k)].height = h
                    n += 1
        for coord, d in s.get("cells", {}).items():
            _restyle(ws[coord], d)
            n += 1
        for label, d in s.get("colstyle", {}).items():
            col = letter.get(label)
            if col:
                for r in range(FIRST, ws.max_row + 1):
                    _restyle(ws[f"{col}{r}"], d)
                n += 1
        for coord, v in s.get("values", {}).items():
            if str(ws[coord].value) == v["from"]:
                ws[coord].value = v["to"]
                n += 1
    return n


if __name__ == "__main__":
    import sys
    print(capture(sys.argv[1], sys.argv[2]), "changes recorded")
    log("owner_layout", "ok", "captured from files given on the command line")
