"""Collect new headlines for tracked projects and accounts via SerpApi Google News.

Each query costs one SerpApi search. New articles are de-duplicated by URL and
added to data/news.csv; the workbook's News tab shows the latest first.
"""
from __future__ import annotations

import argparse

from .common import load_config, log, read_table, today, write_table
from .serp import BudgetExceeded, Serp, pick_items


def query_for(row: dict, kind: str) -> str:
    if row.get("news_query"):
        return row["news_query"]
    if kind == "projects":
        return f"\"{row['developer']}\" data centre {row['city_region'] or row['country']}"
    return f"\"{row['account']}\" data centre Europe"


def run(dry_run: bool = False) -> int:
    cfg = load_config()["serpapi"]
    serp = Serp(dry_run=dry_run)
    news = read_table("news")
    seen = {n["url"] for n in news}
    last = {}
    for n in news:
        last[n["item_id"]] = max(last.get(n["item_id"], ""), n["date_found"])
    items = [("projects", r) for r in pick_items("projects", cfg["news_per_run"], 7, last)]
    added, queries = 0, 0
    for kind, row in items:
        q = query_for(row, kind)
        try:
            data = serp.search("google_news", q)
        except BudgetExceeded as e:
            log("news_watch", "stopped", str(e))
            break
        queries += 1
        for art in (data.get("news_results") or [])[:5]:
            # Story clusters nest articles under "stories"
            arts = art.get("stories") or [art]
            for a in arts[:2]:
                url = a.get("link") or ""
                if not url or url in seen:
                    continue
                seen.add(url)
                src = a.get("source") or {}
                news.append({
                    "date_found": today(), "item_id": row["id"],
                    "item_name": row.get("project") or row.get("account"),
                    "title": a.get("title", ""),
                    "publisher": src.get("name", "") if isinstance(src, dict) else str(src),
                    "published": a.get("date", ""), "url": url,
                    "snippet": a.get("snippet", ""), "query": q,
                })
                added += 1
    news.sort(key=lambda n: (n["date_found"], n["published"]), reverse=True)
    write_table("news", news[:400])
    log("news_watch", "dry-run" if dry_run else "ok",
        f"{queries} queries, {added} new articles", 0 if dry_run else serp.spent)
    return added


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    run(p.parse_args().dry_run)
