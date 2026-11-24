"""Thin SerpApi client with a monthly credit budget.

Engines used:
  google_ai_mode - Google AI Mode answer + references (status scans)
  google_news    - Google News results with direct article links (news watch)
  google_jobs    - job postings by developers (hiring signals)
  google_events  - industry events in target markets
Docs: https://serpapi.com/google-ai-mode-api , https://serpapi.com/google-news-api
"""
from __future__ import annotations

import re

import requests

from .common import credits_used_this_month, fixture, load_config, serpapi_keys

ENDPOINT = "https://serpapi.com/search.json"


class BudgetExceeded(RuntimeError):
    pass


class Serp:
    def __init__(self, dry_run: bool = False):
        self.cfg = load_config()["serpapi"]
        self.dry_run = dry_run
        self.keys = [] if dry_run else serpapi_keys()
        self.k = 0
        self.key = self.keys[0] if self.keys else None
        self.spent = 0
        if not dry_run and not self.key:
            raise RuntimeError("SERPAPI_KEY not set (add it to .env or GitHub secret)")

    def remaining(self) -> int:
        # budget is per key: e.g. 2 keys x 230 = 460 searches a month
        return self.cfg["monthly_budget"] * max(len(self.keys), 1) - credits_used_this_month() - self.spent

    def _rotate(self) -> bool:
        if self.k + 1 < len(self.keys):
            self.k += 1
            self.key = self.keys[self.k]
            return True
        return False

    def search(self, engine: str, q: str, **params) -> dict:
        if self.dry_run:
            if engine not in ("google_ai_mode", "google_news"):
                return {}
            name = "serpapi_ai_mode.json" if engine == "google_ai_mode" else "serpapi_news.json"
            return fixture(name)
        if self.remaining() <= 0:
            raise BudgetExceeded("Monthly SerpApi budget reached (see config.yaml)")
        while True:
            p = {"engine": engine, "q": q, "api_key": self.key,
                 "hl": self.cfg["hl"], "gl": self.cfg["gl"], **params}
            r = requests.get(ENDPOINT, params=p, timeout=90)
            data = r.json() if r.headers.get("content-type", "").startswith("application/json") else {}
            err = data.get("error", "") if r.status_code != 200 or data.get("error") else ""
            if err and re.search(r"run out|limit|exceeded|invalid api key", err, re.I) and self._rotate():
                continue  # this key is used up: try the next one
            if r.status_code != 200 and not data:
                r.raise_for_status()
            if err and "hasn't returned any results" not in err:
                raise RuntimeError(f"SerpApi error: {err}")
            break
        # Cached searches are free on SerpApi; only count fresh ones.
        meta = data.get("search_metadata", {})
        if not meta.get("cached", False):
            self.spent += 1
        return data


def pick_items(kind: str, limit: int, rescan_days: int, last_scanned: dict[str, str]) -> list[dict]:
    """Choose which projects/accounts to query this run.

    Order: not scanned recently first, then highest priority (stage closest to
    consultant appointment, existing client, bigger MW).
    """
    from .common import days_since, read_table
    stage_rank = {"Permitting": 0, "Announced": 1, "Permitted": 2,
                  "Under construction": 3, "Operational": 5}
    rows = read_table(kind)
    due = []
    for r in rows:
        ds = days_since(last_scanned.get(r["id"], ""))
        if ds is not None and ds < rescan_days:
            continue
        try:
            mw = float(r.get("capacity_mw") or 0)
        except ValueError:
            mw = 0
        rel = 0 if "existing" in (r.get("ethos_relationship") or "").lower() else 1
        due.append((ds is not None, stage_rank.get(r.get("stage", ""), 4), rel, -mw, r))
    due.sort(key=lambda t: t[:4])
    return [t[4] for t in due[:limit]]
