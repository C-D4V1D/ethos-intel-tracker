"""Pull Spain's monthly generation mix from Red Electrica (REE) REData API.

Free, no API key. Tracks how far Spain is from the 90% renewable level at which
Real Decreto 1032/2026's hourly-matching rule for data centres eases.
Docs: https://www.ree.es/en/datos/apidata
"""
from __future__ import annotations

import argparse
import datetime as dt
from collections import defaultdict

import requests

from .common import fixture, load_config, log, now, read_table, write_table

URL = "https://apidatos.ree.es/en/datos/generacion/estructura-generacion"


def month_windows(months_back: int) -> list[tuple[str, str]]:
    """REE caps monthly queries at about a year, so ask in yearly slices."""
    end = dt.date.today().replace(day=1) - dt.timedelta(days=1)  # last full month
    start = (end.replace(day=1) - dt.timedelta(days=31 * (months_back - 1))).replace(day=1)
    windows, cur = [], start
    while cur <= end:
        nxt = min(dt.date(cur.year, 12, 31), end)
        windows.append((f"{cur.isoformat()}T00:00", f"{nxt.isoformat()}T23:59"))
        cur = nxt + dt.timedelta(days=1)
    return windows


def parse(payload: dict) -> dict[str, dict]:
    ren, tot = defaultdict(float), defaultdict(float)
    for item in payload.get("included", []):
        attrs = item.get("attributes", {})
        kind = (attrs.get("type") or "").lower()
        title = (attrs.get("title") or "").lower()
        for v in attrs.get("values", []):
            month = (v.get("datetime") or "")[:7]
            val = float(v.get("value") or 0)
            if "total" in title or "total" in kind:
                continue
            tot[month] += val
            if kind.startswith("renovable") or kind.startswith("renewable"):
                ren[month] += val
    out = {}
    for m in tot:
        if tot[m] > 0:
            out[m] = {"renewable_gwh": ren[m] / 1000, "total_gwh": tot[m] / 1000,
                      "share": 100 * ren[m] / tot[m]}
    return out


def run(dry_run: bool = False) -> int:
    cfg = load_config()["grid_data"]
    months: dict[str, dict] = {}
    if dry_run:
        months.update(parse(fixture("ree_generation.json")))
    else:
        for start, end in month_windows(cfg["spain_months_back"]):
            r = requests.get(URL, params={"start_date": start, "end_date": end,
                                          "time_trunc": "month"}, timeout=30)
            r.raise_for_status()
            months.update(parse(r.json()))
    existing = {row["month"]: row for row in read_table("grid_es")}
    stamp = now()
    for m, v in months.items():
        existing[m] = {"month": m, "renewable_share_pct": f"{v['share']:.1f}",
                       "renewable_gwh": f"{v['renewable_gwh']:.0f}",
                       "total_gwh": f"{v['total_gwh']:.0f}", "fetched": stamp}
    rows = sorted(existing.values(), key=lambda r: r["month"])[-cfg["spain_months_back"]:]
    write_table("grid_es", rows)
    latest = rows[-1] if rows else {}
    log("refresh_grid_data", "dry-run" if dry_run else "ok",
        f"{len(months)} months; latest {latest.get('month')} = {latest.get('renewable_share_pct')}% renewable")
    return len(months)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    run(p.parse_args().dry_run)
