"""Thin SerpApi client with a monthly credit budget.

Engines used:
  google_ai_mode - Google AI Mode answer + references (status scans)
  google_news    - Google News results (news watch)
Docs: https://serpapi.com/google-ai-mode-api , https://serpapi.com/google-news-api
"""
from __future__ import annotations

import requests

from .common import credits_used_this_month, fixture, load_config, serpapi_key

ENDPOINT = "https://serpapi.com/search.json"


class BudgetExceeded(RuntimeError):
    pass


class Serp:
    def __init__(self, dry_run: bool = False):
        self.cfg = load_config()["serpapi"]
        self.dry_run = dry_run
        self.key = None if dry_run else serpapi_key()
        self.spent = 0
        if not dry_run and not self.key:
            raise RuntimeError("SERPAPI_KEY not set (add it to .env or GitHub secret)")

    def remaining(self) -> int:
        return self.cfg["monthly_budget"] - credits_used_this_month() - self.spent

    def search(self, engine: str, q: str, **params) -> dict:
        if self.dry_run:
            name = "serpapi_ai_mode.json" if engine == "google_ai_mode" else "serpapi_news.json"
            return fixture(name)
        if self.remaining() <= 0:
            raise BudgetExceeded("Monthly SerpApi budget reached (see config.yaml)")
        p = {"engine": engine, "q": q, "api_key": self.key,
             "hl": self.cfg["hl"], "gl": self.cfg["gl"], **params}
        r = requests.get(ENDPOINT, params=p, timeout=90)
        r.raise_for_status()
        data = r.json()
        if data.get("error"):
            raise RuntimeError(f"SerpApi error: {data['error']}")
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
