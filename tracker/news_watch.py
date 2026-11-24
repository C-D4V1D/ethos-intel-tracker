"""Collect new headlines for tracked projects.

Default source: Google News RSS (free, no key, no credit limit), every project, every run.
Optional: SerpApi Google News (config news.source: serpapi), 1 credit per query.

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


def rss(q: str) -> list[dict]:
    import urllib.parse
    import xml.etree.ElementTree as ET
    import requests
    url = ("https://news.google.com/rss/search?q=" + urllib.parse.quote(q + " when:90d")
           + "&hl=en-GB&gl=GB&ceid=GB:en")
    r = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0 EthosIntelTracker/1.0"})
    r.raise_for_status()
    out = []
    for it in ET.fromstring(r.content).iter("item"):
        src = it.find("source")
        title = it.findtext("title", "")
        pub = src.text if src is not None else ""
        if pub and title.endswith(" - " + pub):
            title = title[: -len(pub) - 3]
        out.append({"title": title, "link": it.findtext("link", ""), "source": {"name": pub},
                    "date": it.findtext("pubDate", "")[5:16], "snippet": ""})
    return out


def bing_rss(q: str) -> list[dict]:
    import urllib.parse
    import xml.etree.ElementTree as ET
    import requests
    url = "https://www.bing.com/news/search?format=rss&q=" + urllib.parse.quote(q)
    r = requests.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0 EthosIntelTracker/1.0"})
    r.raise_for_status()
    out = []
    for it in ET.fromstring(r.content).iter("item"):
        link = it.findtext("link", "")
        if "url=" in link:  # Bing wraps the publisher link
            link = urllib.parse.unquote(urllib.parse.parse_qs(urllib.parse.urlparse(link).query).get("url", [link])[0])
        src = it.findtext("{https://www.bing.com/news/search?q=rss}Source") or urllib.parse.urlparse(link).netloc
        out.append({"title": it.findtext("title", ""), "link": link, "source": {"name": src},
                    "date": it.findtext("pubDate", "")[5:16], "snippet": it.findtext("description", "")[:200]})
    return out


def run_rss(dry_run: bool) -> int:
    import time
    news = read_table("news")
    seen = {n["url"] for n in news}
    seen_titles = {n["title"].lower() for n in news}
    added, queries, fails = 0, 0, 0
    used, errs = {}, []
    for row in read_table("projects"):
        q = query_for(row, "projects")
        arts = []
        if not dry_run:
            for fn in (rss, bing_rss):
                try:
                    arts = fn(q)
                    used[fn.__name__] = used.get(fn.__name__, 0) + 1
                    break
                except Exception as e:  # noqa: BLE001 - try the next free feed
                    errs.append(f"{fn.__name__}: {type(e).__name__} {str(e)[:120]}")
            else:
                fails += 1
                continue
        queries += 1
        for a in arts[:3]:
            if not a["link"] or a["link"] in seen or a["title"].lower() in seen_titles:
                continue
            seen.add(a["link"]); seen_titles.add(a["title"].lower())
            news.append({"date_found": today(), "item_id": row["id"], "item_name": row["project"],
                         "title": a["title"], "publisher": a["source"]["name"], "published": a["date"],
                         "url": a["link"], "snippet": "", "query": q})
            added += 1
        time.sleep(1)
    # market-wide feeds: raw material for automatic discovery of NEW projects
    from .auto_update import CODES
    for country in load_config().get("discovery", {}).get("markets", []):
        q = f'"data centre" OR "data center" {country} MW (planning OR approved OR announces OR construction)'
        try:
            arts = [] if dry_run else rss(q)
        except Exception:  # noqa: BLE001
            arts = []
        for a in arts[:8]:
            if not a["link"] or a["link"] in seen or a["title"].lower() in seen_titles:
                continue
            seen.add(a["link"]); seen_titles.add(a["title"].lower())
            news.append({"date_found": today(), "item_id": f"MKT-{CODES.get(country, country)}",
                         "item_name": f"{country} market", "title": a["title"],
                         "publisher": a["source"]["name"], "published": a["date"], "url": a["link"],
                         "snippet": "", "query": q})
            added += 1
        time.sleep(1)
    news.sort(key=lambda n: n["date_found"], reverse=True)
    write_table("news", news[:400])
    log("news_watch", "dry-run" if dry_run else "ok",
        f"RSS: {queries} queries ({used}), {added} new articles, {fails} failed"
        + (f"; first error: {errs[0]}" if errs else ""))
    return added


def run(dry_run: bool = False) -> int:
    if load_config().get("news", {}).get("source", "rss") == "rss":
        return run_rss(dry_run)
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
