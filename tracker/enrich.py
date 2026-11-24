"""AI enrichment: turns raw feeds into short, readable intelligence. Gemini (free tier) for
the writing, SerpApi for extra searches, Jina / direct fetch to read pages.

  news_summaries  - 2-sentence summary + "what it means for Ethos" for each new headline
  tender_triage   - English title, relevance to Ethos (1-5) and why, for each tender
  hiring_signals  - Google Jobs: data-centre jobs at tracked developers, with what each signals
  events          - Google Events: industry events in the focus markets, with why to go
  gap_answers     - Google AI Mode tries to answer each open question -> Review tab
  policy_watch    - Google AI Mode checks each market's rules -> proposed updates in Review tab
  account_notes   - one-paragraph "latest" per account from its projects, news and jobs
  briefing        - 5-bullet weekly briefing for the Dashboard
Every step is optional and skips quietly when a key is missing.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os

from . import review as RV
from .common import days_since, load_config, log, read_table, today, write_table
from .llm import gemini, gemini_json, gemini_key, read_page, resolve
from .serp import BudgetExceeded, Serp

ETHOS = ("Ethos Engineering is an Irish M&E (mechanical & electrical), energy and sustainability engineering "
         "consultancy that designs data centres (existing clients include Digital Realty, EdgeConneX and atNorth). "
         "It wins work as design consultant, mostly between project announcement and planning approval.")


def _chunks(xs, n):
    for i in range(0, len(xs), n):
        yield xs[i:i + n]


def news_summaries(max_items=48) -> str:
    news = read_table("news")
    todo = [n for n in news if not n.get("summary")][:max_items]
    done = 0
    for batch in _chunks(todo, 6):
        items = []
        for i, n in enumerate(batch):
            text = "" if "news.google.com" in n["url"] else read_page(n["url"])[:2500]
            items.append({"i": i, "headline": n["title"], "publisher": n["publisher"],
                          "about": n["item_name"], "text": text or n.get("snippet", "")})
        prompt = (f"{ETHOS}\nFor each article return JSON {{\"items\": [{{\"i\": n, \"summary\": \"<2 plain sentences, "
                  "facts only, max 45 words>\", \"implication\": \"<1 sentence: what it means for Ethos's "
                  "business development, max 25 words>\"}}]}}. If the text is missing, summarise from the "
                  "headline and say so briefly. Articles:\n" + json.dumps(items, ensure_ascii=False))
        try:
            res = gemini_json(prompt)[0].get("items", [])
        except Exception:  # noqa: BLE001
            continue
        for x in res:
            try:
                n = batch[int(x["i"])]
            except (KeyError, ValueError, IndexError):
                continue
            n["summary"], n["implication"] = x.get("summary", ""), x.get("implication", "")
            done += 1
    write_table("news", news)
    return f"news summarised {done}"


def tender_triage(max_items=60) -> str:
    tenders = read_table("tenders")
    todo = [t for t in tenders if not t.get("relevance")][:max_items]
    done = 0
    for batch in _chunks(todo, 12):
        items = [{"i": i, "title": t["title"], "buyer": t["buyer"], "country": t["country"],
                  "type": t["notice_type"]} for i, t in enumerate(batch)]
        prompt = (f"{ETHOS}\nFor each public tender return JSON {{\"items\": [{{\"i\": n, \"title_en\": \"<English "
                  "title, short>\", \"relevance\": <1-5, 5 = Ethos could bid for design/engineering work on a data "
                  "centre or its power/cooling; 1 = unrelated (software, cleaning, IT hardware)>, \"why\": \"<max 15 "
                  "words>\"}}]}}. Tenders:\n" + json.dumps(items, ensure_ascii=False))
        try:
            res = gemini_json(prompt)[0].get("items", [])
        except Exception:  # noqa: BLE001
            continue
        for x in res:
            try:
                t = batch[int(x["i"])]
            except (KeyError, ValueError, IndexError):
                continue
            t["title_en"], t["relevance"], t["why"] = x.get("title_en", ""), str(x.get("relevance", "")), x.get("why", "")
            done += 1
    tenders.sort(key=lambda t: t["published"], reverse=True)
    tenders.sort(key=lambda t: -(int(t["relevance"]) if t.get("relevance", "").isdigit() else 0))
    write_table("tenders", tenders)
    return f"tenders triaged {done}"


def hiring_signals(serp) -> str:
    accounts, projects = read_table("accounts"), read_table("projects")
    jobs = read_table("jobs")
    seen = {(j["company"], j["title"], j["location"]) for j in jobs}
    pairs = []
    for a in accounts:
        name = a["account"].split(" (")[0]
        countries = sorted({p["country"] for p in projects if name.lower()[:6] in p["developer"].lower()})
        for c in countries[:2]:
            pairs.append((name, c))
    new = []
    for name, country in pairs:
        try:
            data = serp.search("google_jobs", f"{name} data center", location=country)
        except BudgetExceeded:
            break
        except Exception:  # noqa: BLE001
            continue
        for j in data.get("jobs_results", [])[:6]:
            comp = j.get("company_name", "")
            if name.lower()[:6] not in comp.lower():
                continue
            key = (comp, j.get("title", ""), j.get("location", ""))
            if key in seen:
                continue
            seen.add(key)
            link = (j.get("apply_options") or [{}])[0].get("link", "") or j.get("share_link", "")
            new.append({"date_found": today(), "company": comp, "title": j.get("title", ""),
                        "location": j.get("location", ""), "country": country,
                        "posted": (j.get("detected_extensions") or {}).get("posted_at", ""),
                        "via": j.get("via", ""), "url": link, "signal": "",
                        "_desc": (j.get("description") or "")[:600]})
    if new and gemini_key():
        for batch in _chunks(new, 12):
            items = [{"i": i, "company": x["company"], "title": x["title"], "location": x["location"],
                      "description": x["_desc"]} for i, x in enumerate(batch)]
            try:
                res = gemini_json(f"{ETHOS}\nFor each job posting return JSON {{\"items\": [{{\"i\": n, \"signal\": "
                                  "\"<max 20 words: what this hire suggests about the company's data-centre plans "
                                  "in that location (e.g. new build moving to construction) and whether Ethos should "
                                  "act>\"}}]}}. Jobs:\n" + json.dumps(items, ensure_ascii=False))[0].get("items", [])
            except Exception:  # noqa: BLE001
                res = []
            for x in res:
                try:
                    batch[int(x["i"])]["signal"] = x.get("signal", "")
                except (KeyError, ValueError, IndexError):
                    pass
    for x in new:
        x.pop("_desc", None)
    jobs = new + jobs
    jobs = [j for j in jobs if (days_since(j["date_found"]) or 0) <= 120][:300]
    write_table("jobs", jobs)
    return f"jobs {len(new)} new"


def events(serp) -> str:
    cfg = load_config().get("discovery", {})
    rows = read_table("events")
    seen = {(e["name"], e["date"]) for e in rows}
    new = []
    for country in cfg.get("markets", []):
        try:
            data = serp.search("google_events", f"data centre conference {country}")
        except BudgetExceeded:
            break
        except Exception:  # noqa: BLE001
            continue
        for e in data.get("events_results", [])[:6]:
            name, when = e.get("title", ""), (e.get("date") or {}).get("when", "")
            if (name, when) in seen:
                continue
            seen.add((name, when))
            new.append({"date_found": today(), "name": name, "date": when,
                        "venue": ", ".join(e.get("address") or []), "country": country,
                        "url": e.get("link", ""), "why": ""})
    if new and gemini_key():
        items = [{"i": i, "name": x["name"], "when": x["date"], "where": x["venue"]} for i, x in enumerate(new)]
        try:
            res = gemini_json(f"{ETHOS}\nFor each event return JSON {{\"items\": [{{\"i\": n, \"why\": \"<max 18 words: "
                              "who Ethos could meet or why it matters; say 'Low relevance' if it is not about data "
                              "centres, energy or construction>\"}}]}}. Events:\n" + json.dumps(items))[0].get("items", [])
        except Exception:  # noqa: BLE001
            res = []
        for x in res:
            try:
                new[int(x["i"])]["why"] = x.get("why", "")
            except (KeyError, ValueError, IndexError):
                pass
    new = [e for e in new if not e["why"].lower().startswith("low relevance")]
    write_table("events", (new + rows)[:150])
    return f"events {len(new)} new"


def _ai_mode(serp, q):
    from .auto_update import _ai_mode as am
    return am(serp, q)


def gap_answers(serp, max_items=8) -> str:
    gaps, reviews = read_table("gaps"), read_table("reviews")
    asked = 0
    for g in [g for g in gaps if g["status"] in ("Open", "In progress")][:max_items]:
        if any(r["item_id"] == g["id"] and r["decision"] == "Pending" for r in reviews):
            continue
        try:
            text, refs = _ai_mode(serp, f"{g['gap']}. Context: European data centre market, 2026. Give the answer with sources and dates.")
        except BudgetExceeded:
            break
        except Exception:  # noqa: BLE001
            continue
        asked += 1
        try:
            res = gemini_json("Does the text below answer this question? Question: " + g["gap"] +
                              "\nReturn JSON {\"answered\": true/false, \"answer\": \"<max 40 words>\", "
                              "\"quote\": \"<exact supporting words>\", \"url\": \"<reference URL>\"}.\nReferences:\n"
                              + "\n".join(u for _t, u in refs[:8]) + "\nText:\n" + text[:5000])[0]
        except Exception:  # noqa: BLE001
            continue
        if res.get("answered") and res.get("answer"):
            RV.add(reviews, "Gaps", g["id"], g["gap"][:60], "answer", "", res["answer"], res.get("quote", ""),
                   resolve(res.get("url") or ""), "AI found a possible answer; approve to close the question")
    write_table("reviews", reviews)
    return f"gaps asked {asked}"


def policy_watch(serp) -> str:
    policy, reviews = read_table("policy"), read_table("reviews")
    n = 0
    for row in policy:
        try:
            text, refs = _ai_mode(serp, f"Latest 2026 changes to rules and grid connection for data centres in "
                                        f"{row['country']}: new laws, regulator decisions, connection pauses or "
                                        "fast-track schemes, with dates and the next milestone. Cite sources.")
        except BudgetExceeded:
            break
        except Exception:  # noqa: BLE001
            continue
        current = {k: row[k] for k in ("policy_change", "policy_date", "next_milestone", "milestone_date", "rag")}
        try:
            res = gemini_json("Compare the current record with the new search answer. Propose updates ONLY for "
                              "facts that are newer or more precise. Return JSON {\"updates\": [{\"field\": one of "
                              "policy_change|policy_date|next_milestone|milestone_date|rag, \"value\": \"...\", "
                              "\"quote\": \"<exact words>\", \"url\": \"<reference>\"}]} (dates YYYY-MM-DD; rag = Red, "
                              "Amber or Green for how hard it is to build a new data centre now). Empty list if "
                              f"nothing new.\nCurrent record: {json.dumps(current)}\nReferences:\n"
                              + "\n".join(u for _t, u in refs[:8]) + "\nAnswer:\n" + text[:5000])[0].get("updates", [])
        except Exception:  # noqa: BLE001
            continue
        for u in res[:3]:
            f = u.get("field")
            if f in current and u.get("value") and str(u["value"]) != current[f]:
                RV.add(reviews, "Grid & Policy", row["country"], row["country"], f, current[f], u["value"],
                       u.get("quote", ""), resolve(u.get("url") or ""), "AI search suggests a newer rule or date")
                n += 1
    write_table("reviews", reviews)
    return f"policy proposals {n}"


def fill_prices(serp) -> str:
    """Markets with no open price feed: ask Google AI Mode for last month's average day-ahead price,
    and write it only if the number is on the cited page next to EUR/MWh (otherwise -> Review)."""
    from .auto_update import _ai_mode as am
    from .llm import read_page
    rows = read_table("grid_markets")
    months = sorted({r["month"] for r in rows})
    if len(months) < 2:
        return "prices: no months"
    target = months[-2]  # last complete month
    reviews, done = read_table("reviews"), 0
    zone = {"Spain": "Spain (OMIE)", "United Kingdom": "Great Britain (convert GBP to EUR only if the source does)",
            "Ireland": "Ireland (SEM / SEMO)", "Sweden": "Sweden SE3 (Stockholm)"}
    for country, z in zone.items():
        row = next((r for r in rows if r["country"] == country and r["month"] == target), None)
        if not row or row.get("price_eur_mwh"):
            continue
        try:
            text, refs = am(serp, f"Average day-ahead wholesale electricity price in {z} for {target} in EUR/MWh. Cite the source.")
            res = gemini_json("From the answer, return JSON {\"price\": <number in EUR/MWh or null>, \"quote\": \"<exact "
                              f"words>\", \"url\": \"<reference>\"}} for the {target} monthly average day-ahead price in "
                              f"{z}. Null if not stated for that exact month.\nReferences:\n"
                              + "\n".join(u for _t, u in refs[:8]) + "\nAnswer:\n" + text[:4000])[0]
        except BudgetExceeded:
            break
        except Exception:  # noqa: BLE001
            continue
        if not res.get("price"):
            continue
        val = f"{float(res['price']):.1f}"
        url = resolve(res.get("url") or "")
        page = read_page(url).lower() if url else ""
        num = val.rstrip("0").rstrip(".")
        if page and (num in page or val in page) and ("mwh" in page):
            row["price_eur_mwh"] = val
            row["source"] = (row["source"] + f"; price via AI search: {url}").strip("; ")
            done += 1
        else:
            RV.add(reviews, "Grid by market", f"{country}|{target}", f"{country} price {target}", "price_eur_mwh",
                   "", val, res.get("quote", ""), url, "AI found a price but the cited page does not show it")
    write_table("grid_markets", rows)
    write_table("reviews", reviews)
    return f"prices filled {done}"


def source_dates(max_items=40) -> str:
    """Fill blank 'published' dates on Sources from the page itself (meta tags), free."""
    import re
    import requests
    from .llm import UA
    src = read_table("sources")
    done = 0
    for s in [s for s in src if not s.get("published")][:max_items]:
        try:
            html = requests.get(s["url"], timeout=20, headers=UA).text[:200000]
        except Exception:  # noqa: BLE001
            continue
        m = re.search(r'(?:article:published_time|datePublished|pubdate|publish-date|date)["\']?\s*(?:content|:)\s*=?\s*["\']'
                      r'(\d{4}-\d{2}-\d{2})', html, re.I)
        if m:
            s["published"] = m.group(1)
            done += 1
    write_table("sources", src)
    return f"source dates {done}"


def account_notes() -> str:
    accounts, projects = read_table("accounts"), read_table("projects")
    news, jobs = read_table("news"), read_table("jobs")
    done = 0
    for a in accounts:
        name = a["account"].split(" (")[0]
        ps = [p for p in projects if name.lower()[:6] in p["developer"].lower()]
        ids = {p["id"] for p in ps}
        facts = {"projects": [f"{p['project']} - {p['stage']}, {p['capacity_mw'] or '?'} MW" for p in ps],
                 "news": [f"{n['title']} ({n['published']})" for n in news if n["item_id"] in ids][:5],
                 "jobs": [f"{j['title']} in {j['location']}" for j in jobs if name.lower()[:6] in j["company"].lower()][:5]}
        try:
            txt, _m = gemini(f"{ETHOS}\nWrite 2 short sentences (max 45 words) on what is new at {name} and the best "
                             f"opening for Ethos now. Use only these facts: {json.dumps(facts)}")
        except Exception:  # noqa: BLE001
            continue
        a["why_now"] = " ".join(txt.split())
        a["last_checked"] = today()
        done += 1
    write_table("accounts", accounts)
    return f"accounts noted {done}"


def briefing() -> str:
    week = lambda d: (days_since(d) or 999) <= 7
    ch = [c for c in read_table("changes") if week(c["date"])][:40]
    nw = [n for n in read_table("news") if week(n["date_found"]) and n.get("implication")][:20]
    td = [t for t in read_table("tenders") if t.get("relevance") in ("4", "5")][:8]
    jb = [j for j in read_table("jobs") if week(j["date_found"])][:10]
    pend = sum(1 for r in read_table("reviews") if r["decision"] == "Pending")
    facts = {"changes": [f"{c['item_name']}: {c['field']} {c['old_value']} -> {c['new_value']}" for c in ch],
             "news": [f"{n['title']} -> {n['implication']}" for n in nw],
             "tenders": [t.get("title_en") or t["title"] for t in td],
             "jobs": [f"{j['company']}: {j['title']} ({j['location']})" for j in jb],
             "pending_reviews": pend}
    try:
        txt, _m = gemini(f"{ETHOS}\nWrite this week's briefing for Ethos's business development team: exactly 5 bullet "
                         "points starting with '• ', max 25 words each, most important first, plain English, "
                         "only from these facts (say if a week was quiet):\n" + json.dumps(facts, ensure_ascii=False))
    except Exception as e:  # noqa: BLE001
        return f"briefing failed: {str(e)[:80]}"
    rows = read_table("briefing")
    rows = [{"date": today(), "text": txt.strip()}] + [r for r in rows if r["date"] != today()]
    write_table("briefing", rows[:30])
    return "briefing written"


def run(dry_run: bool = False, weekly: bool = True) -> None:
    if dry_run or not gemini_key():
        log("enrich", "skipped", "dry-run or no GEMINI_API_KEY")
        return
    notes = []
    for fn in (news_summaries, tender_triage):
        try:
            notes.append(fn())
        except Exception as e:  # noqa: BLE001
            notes.append(f"{fn.__name__} failed: {str(e)[:80]}")
    if weekly and os.environ.get("SKIP_SERP") != "true":
        try:
            serp = Serp()
        except RuntimeError:
            serp = None
        if serp:
            for fn in (hiring_signals, events, gap_answers, policy_watch, fill_prices):
                try:
                    notes.append(fn(serp))
                except Exception as e:  # noqa: BLE001
                    notes.append(f"{fn.__name__} failed: {str(e)[:80]}")
            log("enrich", "credits", "SerpApi searches used by enrichment", serp.spent)
    for fn in (source_dates, account_notes, briefing):
        try:
            notes.append(fn())
        except Exception as e:  # noqa: BLE001
            notes.append(f"{fn.__name__} failed: {str(e)[:80]}")
    log("enrich", "ok", "; ".join(notes))


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--light", action="store_true", help="no SerpApi searches")
    a = ap.parse_args()
    run(a.dry_run, not a.light)
