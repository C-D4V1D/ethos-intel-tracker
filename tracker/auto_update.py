"""AI auto-update: keep every project current and find new ones, with a source check.

For each project (weekly, within the SerpApi budget):
  1. Ask Google AI Mode for the latest facts (via SerpApi).
  2. Gemini turns the answer into structured fields + the exact supporting quote + the URL.
  3. Jina Reader opens that URL; the value must actually appear on the page.
  4. Verified values are written to the Pipeline automatically, the page is added to
     Sources, and the change is logged ("AI (verified)"). Anything that fails the check,
     or would move a project backwards, goes to the AI Scan tab for a person.

For each market: AI Mode + the free market news feed are asked for NEW projects.
Verified new projects are added to the Pipeline with confidence Low and owner "Auto".

Without GEMINI_API_KEY only the AI Mode answer is stored and flagged (no auto-apply).
"""
from __future__ import annotations

import argparse
import re
import time

START = time.time()


def out_of_time(cfg) -> bool:
    return (time.time() - START) / 60 > cfg.get("auto_update", {}).get("max_minutes", 40)

from .common import load_config, log, read_table, today, write_table
from .llm import gemini_json, gemini_key, page_supports, page_title, read_page
from .serp import BudgetExceeded, Serp, pick_items

FIELDS = ["stage", "capacity_mw", "capacity_basis", "investment_eur_m", "permit",
          "construction_start", "target_online", "design_team"]
STAGES = ["Announced", "Permitting", "Permitted", "Under construction", "Operational", "Cancelled / paused"]
ORDER = {s: i for i, s in enumerate(STAGES)}
CODES = {"Spain": "ES", "United Kingdom": "UK", "Ireland": "IE", "Germany": "DE", "France": "FR",
         "Italy": "IT", "Sweden": "SE", "Denmark": "DK", "Netherlands": "NL", "Norway": "NO",
         "Finland": "FI", "Poland": "PL", "Portugal": "PT", "Greece": "GR"}

EXTRACT = """You extract facts about one data centre project from a search answer. Return JSON only:
{{"facts": {{ "<field>": {{"value": <value>, "quote": "<exact words from the answer that state it>",
   "url": "<the reference URL that supports it, from the list>", "as_of": "<date of the fact if given>"}} }} }}
Fields allowed: stage (one of {stages}), capacity_mw (number, MW), capacity_basis ("IT" or "Grid"),
investment_eur_m (number, € millions; skip if not in euros), permit (date of main approval, YYYY-MM-DD or YYYY),
construction_start (YYYY-MM-DD or YYYY), target_online (YYYY), design_team (named architects/engineers/contractors).
Only include a field if the answer states it explicitly for THIS project. Never guess. Omit unknown fields.

Project: {project} by {developer} in {place}.
References:
{refs}

Answer:
{text}"""

DISCOVER = """From the text below, list data centre projects in {country} that are NOT in this known list:
{known}
Return JSON only: {{"projects": [{{"project": "<name>", "developer": "<company>", "city_region": "<place>",
"capacity_mw": <number or null>, "investment_eur_m": <number or null>, "stage": "<one of {stages}>",
"quote": "<exact words stating the project>", "url": "<reference URL>"}}]}}
Only real, named projects with a developer and a place, announced or progressing in the last 12 months.
Ignore operational sites unless they are expanding. Max 5.

Text:
{text}"""


def _flatten(blocks) -> str:
    out = []
    for b in blocks or []:
        if b.get("snippet"):
            out.append(b["snippet"])
        for li in b.get("list") or []:
            if li.get("snippet"):
                out.append("- " + li["snippet"])
        if b.get("text_blocks"):
            out.append(_flatten(b["text_blocks"]))
    return "\n".join(x for x in out if x)


def _ai_mode(serp, q):
    data = serp.search("google_ai_mode", q)
    text = data.get("reconstructed_markdown") or _flatten(data.get("text_blocks"))
    refs = [(r.get("title") or r.get("source") or "", r.get("link", "")) for r in data.get("references") or []]
    return text, refs


def _norm_value(field, v):
    if v is None or v == "":
        return ""
    if field in ("capacity_mw", "investment_eur_m"):
        try:
            x = float(str(v).replace(",", ""))
            return f"{x:g}"
        except ValueError:
            return ""
    if field == "stage":
        return v if v in STAGES else ""
    if field == "capacity_basis":
        return v if v in ("IT", "Grid") else ""
    return str(v).strip()


def _source_row(sources, url, page, item):
    for s in sources:
        if s["url"] == url:
            if item not in s["used_in"]:
                s["used_in"] = (s["used_in"] + "; " + item).strip("; ")
            return s["id"]
    sid = f"SRC{len(sources) + 1:03d}"
    host = re.sub(r"^www\.", "", re.sub(r"^https?://([^/]+).*", r"\1", url))
    sources.append({"id": sid, "title": page_title(page, url), "publisher": host, "type": "S",
                    "published": "", "url": url, "last_checked": today(), "http_status": "200",
                    "link_status": "OK (reader)", "used_in": item})
    return sid


def _decide(field, old, new):
    """'apply' | 'review' | 'same'."""
    if not new or new == old:
        return "same"
    if not old:
        return "apply"
    if field == "stage":
        return "apply" if ORDER.get(new, -1) > ORDER.get(old, -1) else "review"
    if field in ("capacity_mw", "investment_eur_m"):
        try:
            o, n = float(old), float(new)
        except ValueError:
            return "review"
        return "same" if abs(n - o) <= 0.05 * o else "apply"
    if field == "design_team" and old.lower().startswith("not"):
        return "apply"
    return "apply"


def update_projects(serp, cfg, projects, sources, scans, changes, use_ai):
    scfg = cfg["serpapi"]
    last = {}
    for s in scans:
        last[s["item_id"]] = max(last.get(s["item_id"], ""), s["date"])
    items = pick_items("projects", scfg["ai_mode_per_run"], scfg["rescan_after_days"], last)
    byid = {p["id"]: p for p in projects}
    applied = reviewed = scanned = 0
    for row in items:
        if out_of_time(cfg):
            log("auto_update", "stopped", "time budget reached; remaining projects next run")
            break
        p = byid[row["id"]]
        print(f"[auto_update] {p['id']} {p['project'][:40]}", flush=True)
        place = ", ".join(x for x in (p["city_region"], p["country"]) if x)
        q = p.get("ai_query") or (
            f"Latest status of the {p['project']} data centre by {p['developer']} in {place}: "
            "capacity in MW (IT or grid), investment, planning permit date, construction start, "
            "expected opening, and named architects, engineers or contractors. Cite sources with dates.")
        try:
            text, refs = _ai_mode(serp, q)
        except BudgetExceeded as e:
            log("auto_update", "stopped", str(e))
            break
        scanned += 1
        notes, todo = [], []
        if use_ai and text:
            reflist = "\n".join(f"- {t}: {u}" for t, u in refs[:10]) or "none"
            try:
                facts = gemini_json(EXTRACT.format(stages=", ".join(STAGES), project=p["project"],
                                                   developer=p["developer"], place=place, refs=reflist,
                                                   text=text[:6000]))[0].get("facts", {})
            except Exception as e:  # noqa: BLE001
                facts, notes = {}, [f"Gemini extraction failed: {str(e)[:120]}"]
            for field, f in (facts or {}).items():
                if field not in FIELDS or not isinstance(f, dict):
                    continue
                new = _norm_value(field, f.get("value"))
                old = (p.get(field) or "").strip()
                decision = _decide(field, old, new)
                if decision == "same":
                    continue
                urls = [f.get("url")] + [u for _t, u in refs if u != f.get("url")]
                proof = next((u for u in urls[:2] if u and page_supports(read_page(u), new if field not in ("stage", "design_team") else "", f.get("quote", "")) ), "")
                if decision == "apply" and proof:
                    sid = _source_row(sources, proof, read_page(proof), p["id"])
                    changes.append({"date": today(), "item_type": "projects", "item_id": p["id"],
                                    "item_name": p["project"], "field": field, "old_value": old,
                                    "new_value": new, "impact": "Auto-updated: value found on the cited page",
                                    "source": sid, "logged_by": "AI (verified)"})
                    p[field] = new
                    if sid not in p["source_ids"]:
                        p["source_ids"] = (p["source_ids"] + "; " + sid).strip("; ")
                    p["last_checked"] = today()
                    applied += 1
                else:
                    todo.append(f"{field}: '{old or 'blank'}' -> '{new}'"
                                + (" (not found on cited page)" if not proof else " (would move backwards)"))
                    reviewed += 1
        elif text:
            todo.append("AI answer stored; add GEMINI_API_KEY to extract and auto-apply facts")
        scans.append({"date": today(), "item_id": p["id"], "item_name": p["project"], "query": q,
                      "summary": text[:1800],
                      "references": " | ".join(f"{t}: {u}" for t, u in refs[:6]),
                      "flags": "; ".join(todo + notes) or "Up to date (nothing new, or all changes verified and applied)",
                      "review_status": "Needs review" if todo and use_ai else "No action"})
    return scanned, applied, reviewed


def discover(serp, cfg, projects, sources, scans, changes, use_ai):
    if not use_ai:
        return 0, 0
    dcfg = cfg.get("discovery", {})
    markets = dcfg.get("markets", [])
    news = read_table("news")
    added = queued = 0
    for country in markets:
        if out_of_time(cfg):
            break
        print(f"[auto_update] discovery {country}", flush=True)
        known = [f"{p['project']} ({p['developer']}, {p['city_region']})" for p in projects if p["country"] == country]
        text, refs = "", []
        if serp is not None:
            try:
                text, refs = _ai_mode(serp, f"New data centre projects announced, approved or started in {country} "
                                            "in the last 60 days: developer, location, MW, investment, stage. Cite sources.")
            except BudgetExceeded:
                serp = None
            except Exception:  # noqa: BLE001
                pass
        mk = [n for n in news if n.get("item_id") == f"MKT-{CODES.get(country, country)}"][:8]
        text += "\n" + "\n".join(f"- {n['title']} ({n['publisher']}) {n['url']}" for n in mk)
        refs += [(n["title"], n["url"]) for n in mk]
        if not text.strip():
            continue
        try:
            found = gemini_json(DISCOVER.format(country=country, known="\n".join(known) or "none",
                                                stages=", ".join(STAGES), text=text[:7000]))[0].get("projects", [])
        except Exception:  # noqa: BLE001
            continue
        n_country = sum(1 for p in projects if p["country"] == country)
        for c in found[: dcfg.get("max_new_per_market", 3)]:
            name, dev = (c.get("project") or "").strip(), (c.get("developer") or "").strip()
            if not name or not dev:
                continue
            dup = any(dev.lower()[:6] in p["developer"].lower() and
                      ((c.get("city_region") or "").lower()[:5] in p["city_region"].lower() or
                       name.lower()[:8] in p["project"].lower()) for p in projects)
            if dup:
                continue
            urls = [c.get("url")] + [u for _t, u in refs]
            proof = next((u for u in urls[:2] if u and page_supports(read_page(u), "", f"{dev} {c.get('city_region','')}")), "")
            if not proof:
                scans.append({"date": today(), "item_id": "NEW", "item_name": f"{name} ({country})",
                              "query": "discovery", "summary": c.get("quote", ""), "references": c.get("url", ""),
                              "flags": f"Possible new project: {dev}, {c.get('city_region','')} - not confirmed on a source page",
                              "review_status": "Needs review"})
                queued += 1
                continue
            n_country += 1
            pid = f"P-{CODES.get(country, 'XX')}-{n_country:02d}"
            while any(p["id"] == pid for p in projects):
                n_country += 1
                pid = f"P-{CODES.get(country, 'XX')}-{n_country:02d}"
            sid = _source_row(sources, proof, read_page(proof), pid)
            stage = c.get("stage") if c.get("stage") in STAGES else "Announced"
            mw = _norm_value("capacity_mw", c.get("capacity_mw"))
            inv = _norm_value("investment_eur_m", c.get("investment_eur_m"))
            projects.append({"id": pid, "project": name, "developer": dev, "country": country,
                             "city_region": c.get("city_region", ""),
                             "capacity_mw": mw if mw and page_supports(read_page(proof), mw) else "",
                             "capacity_basis": "Unclear",
                             "investment_eur_m": inv if inv and page_supports(read_page(proof), inv) else "",
                             "project_type": "New build", "stage": stage, "announced": today()[:7],
                             "permit": "", "construction_start": "", "target_online": "",
                             "ethos_fit": "M&E design", "ethos_relationship": "Target",
                             "design_team": "Not public", "next_action": "", "owner": "Auto",
                             "confidence": "Low", "source_ids": sid,
                             "notes": f"Auto-added {today()} from AI search; confirmed on {proof[:60]}. Check fit and owner.",
                             "last_checked": today(), "news_query": f"{dev} data centre {c.get('city_region','')}",
                             "ai_query": "", "suggested_action": "", "suggested_by": ""})
            changes.append({"date": today(), "item_type": "projects", "item_id": pid, "item_name": name,
                            "field": "(new item)", "old_value": "", "new_value": "added",
                            "impact": "New project found automatically and confirmed on a source page",
                            "source": sid, "logged_by": "AI (verified)"})
            added += 1
    return added, queued


def run(dry_run: bool = False) -> dict:
    cfg = load_config()
    use_ai = bool(gemini_key()) and not dry_run
    try:
        serp = Serp(dry_run=dry_run)
    except RuntimeError:
        serp = None
    projects, sources = read_table("projects"), read_table("sources")
    scans, changes = read_table("ai_scan"), []
    s = a = r = 0
    if serp is not None:
        s, a, r = update_projects(serp, cfg, projects, sources, scans, changes, use_ai)
    added, queued = discover(serp, cfg, projects, sources, scans, changes, use_ai)
    if not dry_run:
        write_table("projects", projects)
        write_table("sources", sources)
        if changes:
            write_table("changes", changes + read_table("changes"))
    scans.sort(key=lambda x: x["date"], reverse=True)
    write_table("ai_scan", scans[:400])
    detail = (f"{s} projects checked with AI Mode, {a} facts auto-applied (verified on source page), "
              f"{r} sent for review; {added} new projects added, {queued} possible new projects for review")
    if serp is None:
        detail += " [no SERPAPI_KEY: AI Mode skipped]"
    if not use_ai:
        detail += " [no GEMINI_API_KEY: no extraction or auto-apply]"
    log("auto_update", "dry-run" if dry_run else "ok", detail, 0 if (dry_run or serp is None) else serp.spent)
    return {"scanned": s, "applied": a, "review": r, "added": added}


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--dry-run", action="store_true")
    run(ap.parse_args().dry_run)
