"""Find public tenders that mention data centres in the target markets.

Free, official, no key:
  * TED Search API v3 (EU Tenders Electronic Daily): POST api.ted.europa.eu/v3/notices/search
  * UK Find a Tender Service OCDS API: GET find-tender.service.gov.uk/api/1.0/ocdsReleasePackages

Public-sector work is a small share of data centre demand, but these notices show
government/university/hospital data centre projects and sometimes private campuses'
grid or enabling works. New notices are added to data/tenders.csv.
"""
from __future__ import annotations

import argparse
import datetime as dt
import re

import requests

from .common import load_config, log, read_table, today, write_table

TED_URL = "https://api.ted.europa.eu/v3/notices/search"
FTS_URL = "https://www.find-tender.service.gov.uk/api/1.0/ocdsReleasePackages"
ISO3 = {"Spain": "ESP", "Germany": "DEU", "France": "FRA", "Italy": "ITA", "Sweden": "SWE",
        "Ireland": "IRL", "Denmark": "DNK", "Netherlands": "NLD", "Poland": "POL", "Portugal": "PRT"}


def _first(v):
    """TED returns many fields as lists or {lang: text} dicts; take something readable."""
    if isinstance(v, dict):
        v = v.get("eng") or v.get("ENG") or next(iter(v.values()), "")
    if isinstance(v, list):
        v = v[0] if v else ""
    if isinstance(v, (dict, list)):
        return _first(v)
    return str(v or "")


def ted(cfg, since: dt.date) -> list[dict]:
    terms = " OR ".join(f'"{k}"' for k in cfg["keywords"])
    countries = " ".join(ISO3[c] for c in cfg["countries"] if c in ISO3)
    query = (f"FT~({terms}) AND buyer-country IN ({countries}) "
             f"AND publication-date>={since:%Y%m%d}")
    body = {"query": query,
            "fields": ["publication-number", "notice-title", "buyer-name", "buyer-country",
                       "publication-date", "deadline-receipt-tender-date-lot", "notice-type"],
            "limit": cfg.get("max_results", 50), "page": 1, "scope": "ALL",
            "paginationMode": "PAGE_NUMBER"}
    r = requests.post(TED_URL, json=body, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"TED HTTP {r.status_code}: {r.text[:300]}")
    out = []
    for n in r.json().get("notices", []):
        pn = _first(n.get("publication-number"))
        out.append({"source": "TED", "notice_id": pn, "title": _first(n.get("notice-title")),
                    "buyer": _first(n.get("buyer-name")), "country": _first(n.get("buyer-country")),
                    "published": _first(n.get("publication-date"))[:10],
                    "deadline": _first(n.get("deadline-receipt-tender-date-lot"))[:10],
                    "notice_type": _first(n.get("notice-type")),
                    "url": f"https://ted.europa.eu/en/notice/-/detail/{pn}"})
    return out


def fts(cfg, since: dt.date) -> list[dict]:
    pat = re.compile("|".join(re.escape(k) for k in cfg["keywords"]), re.I)
    params = {"updatedFrom": f"{since:%Y-%m-%d}T00:00:00", "limit": 100}
    out, url, pages = [], FTS_URL, 0
    while url and pages < cfg.get("fts_max_pages", 15):
        r = requests.get(url, params=params if pages == 0 else None, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"FTS HTTP {r.status_code}: {r.text[:300]}")
        data = r.json()
        for rel in data.get("releases", []):
            t = rel.get("tender") or {}
            text = f"{t.get('title','')} {t.get('description','')}"
            m = pat.search(text)
            if not m:
                continue
            buyer = (rel.get("buyer") or {}).get("name", "")
            period = t.get("tenderPeriod") or {}
            out.append({"source": "FTS", "notice_id": rel.get("ocid", rel.get("id", "")),
                        "title": t.get("title", ""), "buyer": buyer, "country": "UK",
                        "published": (rel.get("date") or "")[:10],
                        "deadline": (period.get("endDate") or "")[:10],
                        "notice_type": ", ".join(rel.get("tag", [])), "matched": m.group(0).lower(),
                        "url": f"https://www.find-tender.service.gov.uk/Notice/{rel.get('id','').split('-')[-1]}"
                               if rel.get("id") else ""})
        url = (data.get("links") or {}).get("next")
        pages += 1
    return out


def run(dry_run: bool = False) -> int:
    cfg = load_config().get("tenders", {})
    if not cfg.get("enabled", True):
        return 0
    rows = read_table("tenders")
    seen = {(r["source"], r["notice_id"]) for r in rows}
    since = dt.date.today() - dt.timedelta(days=cfg.get("lookback_days", 30))
    added, notes = 0, []
    for name, fn in (("TED", ted), ("FTS", fts)):
        if dry_run:
            notes.append(f"{name} skipped (dry-run)")
            continue
        try:
            found = fn(cfg, since)
        except (RuntimeError, requests.RequestException, ValueError) as e:
            notes.append(f"{name} failed: {str(e)[:250]}")
            continue
        new = 0
        for t in found:
            if (t["source"], t["notice_id"]) in seen or not t["notice_id"]:
                continue
            t.setdefault("matched", next((k for k in cfg["keywords"]
                                          if k.lower() in t["title"].lower()), "full text"))
            t["date_found"] = today()
            rows.append(t)
            seen.add((t["source"], t["notice_id"]))
            new += 1
        added += new
        notes.append(f"{name}: {len(found)} matches, {new} new")
    rows.sort(key=lambda r: (r["published"], r["date_found"]), reverse=True)
    write_table("tenders", rows[:300])
    log("tender_watch", "dry-run" if dry_run else "ok", "; ".join(notes))
    return added


if __name__ == "__main__":
    a = argparse.ArgumentParser(description=__doc__)
    a.add_argument("--dry-run", action="store_true")
    run(a.parse_args().dry_run)
