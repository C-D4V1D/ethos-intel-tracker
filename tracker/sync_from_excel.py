"""Pull hand edits from the workbook back into data/*.csv before a refresh.

Reads each editable sheet by its header labels (so re-ordering columns in Excel
is fine), skips formula columns, and keeps any CSV-only fields (e.g. news_query).
Run this whenever you have edited the workbook, or let run_all.py do it.
"""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import load_workbook

from .common import OUTPUT, WORKBOOK_NAME, log, read_table, write_table
from .layout import HEADER_ROW, SYNCED


def cell_text(v) -> str:
    if v is None:
        return ""
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10]
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _published() -> Path:
    from . import common
    return common.DATA / ".published.json"


def _same(a: str, b: str) -> bool:
    a, b = (a or "").strip(), (b or "").strip()
    if a == b:
        return True
    try:
        return abs(float(a) - float(b)) < 1e-9
    except ValueError:
        return False


def read_workbook(wb_path) -> dict:
    """{table: {"key": key, "rows": {key_value: {field: text}}}} read from the synced tabs."""
    wb = load_workbook(wb_path, data_only=False)
    out = {}
    for sheet, (table, layout) in SYNCED.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        labels = {c.value: i for i, c in enumerate(ws[HEADER_ROW], start=1) if c.value}
        cols = [(k, labels[lbl]) for k, lbl, _w, kind in layout if not k.startswith("_") and lbl in labels]
        if not cols:
            continue
        key = "id" if "id" in dict(cols) else cols[0][0]
        url_cols = [(k, labels[lbl]) for k, lbl, _w, kind in layout if kind == "url" and lbl in labels]
        rows = {}
        for r in range(HEADER_ROW + 1, ws.max_row + 1):
            rec = {k: cell_text(ws.cell(r, ci).value) for k, ci in cols}
            for k, ci in url_cols:  # link columns show a short domain; keep the full address
                link = ws.cell(r, ci).hyperlink
                if link and link.target:
                    rec[k] = link.target
            if any(rec.values()) and rec.get(key):
                rows[rec[key]] = rec
        out[table] = {"key": key, "rows": rows}
    return out


def remember_published(wb_path) -> None:
    """Called right after the workbook is built: read it back exactly as a person would see it,
    so a later sync can tell which cells a person changed (only those are applied)."""
    import json
    _published().write_text(json.dumps(read_workbook(wb_path), ensure_ascii=False))


def run(path: str | None = None) -> int:
    """Three-way merge: a cell is applied only if it differs from what was published.
    Rows are never deleted by a sync, and newer data written by the robots is never overwritten
    by an older copy of the workbook."""
    import json
    wb_path = Path(path) if path else OUTPUT / WORKBOOK_NAME
    if not wb_path.exists():
        log("sync_from_excel", "skipped", f"{wb_path.name} not found")
        return 0
    if not _published().exists():
        log("sync_from_excel", "skipped", "no record of the last published workbook yet")
        return 0
    pub = json.loads(_published().read_text())
    got = read_workbook(wb_path)
    edited, added = 0, 0
    for table, data in got.items():
        if table not in pub:
            continue
        key, published = pub[table]["key"], pub[table]["rows"]
        current = read_table(table)
        by_key = {r.get(key, ""): r for r in current}
        changed = False
        for k, rec in data["rows"].items():
            if k in published:
                if k not in by_key:      # the robots removed it since; leave it removed
                    continue
                for f, v in rec.items():
                    if f != key and not _same(v, published[k].get(f, "")) and not _same(v, by_key[k].get(f, "")):
                        by_key[k][f] = v
                        edited += 1
                        changed = True
            elif k not in by_key:         # a row the person added
                current.append(rec)
                by_key[k] = rec
                added += 1
                changed = True
        if changed:
            write_table(table, current)
    log("sync_from_excel", "ok", f"{edited} edited cells and {added} new rows taken from {wb_path.name}")
    return edited + added


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--workbook", help="path to an edited workbook (default: output/)")
    run(p.parse_args().workbook)
