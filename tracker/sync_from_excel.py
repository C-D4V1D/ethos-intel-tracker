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


def run(path: str | None = None) -> int:
    wb_path = Path(path) if path else OUTPUT / WORKBOOK_NAME
    if not wb_path.exists():
        log("sync_from_excel", "skipped", f"{wb_path.name} not found")
        return 0
    wb = load_workbook(wb_path, data_only=False)
    synced = 0
    for sheet, (table, layout) in SYNCED.items():
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        labels = {c.value: i for i, c in enumerate(ws[HEADER_ROW], start=1) if c.value}
        cols = [(k, labels[lbl]) for k, lbl, _w, kind in layout
                if not k.startswith("_") and lbl in labels]
        key = "id" if "id" in dict(cols) else cols[0][0]
        old = {r[key]: r for r in read_table(table)}
        rows = []
        for r in range(HEADER_ROW + 1, ws.max_row + 1):
            rec = {k: cell_text(ws.cell(r, ci).value) for k, ci in cols}
            if url_col := dict(cols).get("url"):
                link = ws.cell(r, url_col).hyperlink
                if link and link.target:
                    rec["url"] = link.target
            if not any(rec.values()):
                continue
            base = dict(old.get(rec.get(key, ""), {}))
            base.update(rec)
            rows.append(base)
        if rows:
            write_table(table, rows)
            synced += len(rows)
    log("sync_from_excel", "ok", f"{synced} rows read from {wb_path.name}")
    return synced


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--workbook", help="path to an edited workbook (default: output/)")
    run(p.parse_args().workbook)
