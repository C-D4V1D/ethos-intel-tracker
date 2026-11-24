"""Monthly renewable share of electricity generation for every market.

Free, no key:
  * EU markets: Energy-Charts API (Fraunhofer ISE) - daily renewable share, averaged by month.
  * United Kingdom: NESO / National Grid Carbon Intensity API - half-hourly generation mix.
Why it matters: Ireland (CRU), Germany (EnEfG, 100% from 2027) and Spain (RD 1032/2026)
tie new data centres to renewable supply, so a cleaner grid makes compliance easier.
Day-ahead wholesale prices are added only where the feed is CC BY 4.0 (SMARD/Bundesnetzagentur):
Germany, France, Italy (North), Denmark (DK1). Spain, Sweden and Ireland prices come from EPEX
under a "private and internal use only" licence, so they are not republished here.
"""
from __future__ import annotations

import argparse
import datetime as dt
import statistics

import requests

from .common import log, now, read_table, write_table

EC = {"Spain": "es", "Germany": "de", "France": "fr", "Italy": "it", "Sweden": "se",
      "Ireland": "ie", "Denmark": "dk"}
RENEWABLE_UK = {"wind", "solar", "hydro", "biomass"}


def energy_charts(code: str) -> dict[str, float]:
    r = requests.get(f"https://api.energy-charts.info/ren_share_daily_avg?country={code}", timeout=60)
    r.raise_for_status()
    js = r.json()
    by_month: dict[str, list[float]] = {}
    for day, v in zip(js.get("days", []), js.get("data", [])):
        if v is None:
            continue
        d, m, y = day.split(".")
        by_month.setdefault(f"{y}-{m}", []).append(float(v))
    return {k: round(statistics.mean(v), 1) for k, v in by_month.items() if len(v) >= 20}


OPEN_PRICE_ZONES = {"Germany": "DE-LU", "France": "FR", "Italy": "IT-North", "Denmark": "DK1"}


def prices(zone: str) -> dict[str, float]:
    end = dt.date.today()
    start = (end.replace(day=1) - dt.timedelta(days=365)).replace(day=1)
    r = requests.get(f"https://api.energy-charts.info/price?bzn={zone}&start={start}&end={end}", timeout=120)
    r.raise_for_status()
    js = r.json()
    if "CC BY" not in js.get("license_info", ""):
        return {}  # never republish restricted feeds
    by_month: dict[str, list[float]] = {}
    for t, v in zip(js.get("unix_seconds", []), js.get("price", [])):
        if v is not None:
            by_month.setdefault(dt.datetime.fromtimestamp(t, dt.timezone.utc).strftime("%Y-%m"), []).append(float(v))
    return {k: round(statistics.mean(v), 1) for k, v in by_month.items() if len(v) > 200}


def uk(months: int = 12) -> dict[str, float]:
    out = {}
    first = dt.date.today().replace(day=1)
    for i in range(1, months + 1):
        end = first if i == 1 else start  # noqa: F821 - start set in previous loop
        start = (end - dt.timedelta(days=1)).replace(day=1)
        url = f"https://api.carbonintensity.org.uk/generation/{start:%Y-%m-%d}T00:00Z/{end:%Y-%m-%d}T00:00Z"
        r = requests.get(url, timeout=60, headers={"Accept": "application/json"})
        if r.status_code != 200:
            continue
        vals = [sum(f["perc"] for f in h["generationmix"] if f["fuel"] in RENEWABLE_UK)
                for h in r.json().get("data", [])]
        if len(vals) > 500:
            out[f"{start:%Y-%m}"] = round(statistics.mean(vals), 1)
    return out


def run(dry_run: bool = False) -> int:
    if dry_run:
        log("grid_markets", "dry-run", "skipped (network)")
        return 0
    rows, notes = [], []
    stamp = now()
    for country, code in EC.items():
        try:
            pr = {}
            if country in OPEN_PRICE_ZONES:
                try:
                    pr = prices(OPEN_PRICE_ZONES[country])
                except Exception as e:  # noqa: BLE001
                    notes.append(f"{country} price: {type(e).__name__}")
            for m, v in energy_charts(code).items():
                rows.append({"country": country, "month": m, "renewable_share_pct": str(v),
                             "price_eur_mwh": str(pr.get(m, "")),
                             "source": "Energy-Charts (Fraunhofer ISE)" + ("; prices SMARD, CC BY 4.0" if pr else ""),
                             "fetched": stamp})
        except Exception as e:  # noqa: BLE001
            notes.append(f"{country}: {type(e).__name__}")
    try:
        for m, v in uk().items():
            rows.append({"country": "United Kingdom", "month": m, "renewable_share_pct": str(v),
                         "source": "NESO Carbon Intensity API (GB)", "fetched": stamp})
    except Exception as e:  # noqa: BLE001
        notes.append(f"UK: {type(e).__name__}")
    if rows:
        write_table("grid_markets", sorted(rows, key=lambda r: (r["month"], r["country"])))
    months = sorted({r["month"] for r in rows})
    log("grid_markets", "ok", f"{len({r['country'] for r in rows})} markets, {len(months)} months"
        + (f"; failed: {', '.join(notes)}" if notes else ""))
    return len(rows)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dry-run", action="store_true")
    run(ap.parse_args().dry_run)
