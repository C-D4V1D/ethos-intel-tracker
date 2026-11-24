"""Suggest a next action for every Pipeline project.

1. A rule engine writes a suggestion from the row's stage, Ethos relationship,
   design team, freshness, recent news, AI Scan flags and the market's policy
   rating. Always runs, free, no key.
2. If GEMINI_API_KEY is set, Google's Gemini (free tier) rewrites each
   suggestion using the same facts plus the latest headlines and AI answer.
   It is told not to invent facts; output is labelled "Gemini".

Suggestions go to 'suggested_action' / 'suggested_by'. The human 'next_action'
column is never touched.
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import time

import requests

from .common import days_since, load_config, load_env, log, read_table, today, write_table

EARLY = {"Announced", "Permitting", "Permitted"}


def _recent(rows, item_id, key, days):
    out = [r for r in rows if r.get("item_id") == item_id and (days_since(r.get(key, "")) or 999) <= days]
    return sorted(out, key=lambda r: r.get(key, ""), reverse=True)


def rule_action(p: dict, news, scans, policy, cfg) -> str:
    dev = p["developer"] or "the developer"
    place = p["city_region"] or p["country"]
    team = (p["design_team"] or "").strip()
    team_known = bool(team) and not team.lower().startswith("not ")
    rel = p["ethos_relationship"]
    stage = p["stage"]
    parts = []

    age = days_since(p.get("last_checked", ""))
    if age is not None and age >= cfg["freshness"]["amber_days"]:
        parts.append(f"Re-check status first (last checked {age} days ago).")

    flagged = [s for s in _recent(scans, p["id"], "date", 30) if s.get("review_status") == "Needs review"]
    if flagged:
        parts.append(f"Verify the AI Scan flag ({flagged[0].get('flags','')[:90]}) against a primary source and update the row.")
    fresh = _recent(news, p["id"], "date_found", 14)
    if fresh:
        parts.append(f"Read the new coverage ('{fresh[0].get('title','')[:70]}') and update stage or dates if it changed.")

    if stage == "Announced":
        parts.append(f"Ask {dev} who will lead design for {place} and when consultants will be appointed."
                     if rel == "Existing client" else
                     f"Find {dev}'s design / preconstruction lead for {place} and get an introduction before the design team is set.")
    elif stage == "Permitting":
        parts.append(f"Track the planning decision; prepare a short pack of similar-size Ethos projects for {dev} so Ethos is ready when consent lands.")
    elif stage == "Permitted":
        parts.append(f"Design team held by {team}: offer specialist packages (CFD, energy, commissioning) or target later phases."
                     if team_known else
                     f"Approved and no design team named yet: approach {dev} now, this is when design firms are usually hired.")
    elif stage == "Under construction":
        parts.append(f"Late stage: target fit-out, commissioning or the next phase; log contractor contacts" +
                     (f" ({team})." if team_known else "."))
    elif stage == "Operational":
        parts.append("Watch for expansion or efficiency / heat-reuse retrofit work.")
    elif stage.startswith("Cancelled"):
        parts.append("Archive the opportunity; watch for site resale or relaunch.")

    pol = policy.get(p["country"])
    if pol and pol.get("rag") == "Red" and stage in EARLY:
        parts.append(f"Use the new rules ({pol['policy_change']}) as a talking point on design.")
    return " ".join(parts)


PROMPT = """You help the business development team at Ethos Engineering, an Irish M&E and sustainability
engineering consultancy that designs data centres. Write ONE next action (max 35 words, imperative, specific)
for the project below. Use only the facts given; do not invent names, dates or numbers. If facts are thin,
suggest how to find the missing fact. No preamble, no quotes.

Project facts:
{facts}

Latest headlines:
{news}

Latest Google AI Mode answer (unverified):
{ai}

Market rule context: {policy}

Rule-engine suggestion to improve on: {rule}"""


def gemini(prompt: str, models: list[str], key: str) -> tuple[str, str]:
    last_err = ""
    for m in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{m}:generateContent"
        r = requests.post(url, headers={"x-goog-api-key": key, "Content-Type": "application/json"},
                          json={"contents": [{"parts": [{"text": prompt}]}],
                                "generationConfig": {"temperature": 0.3, "maxOutputTokens": 200}},
                          timeout=60)
        if r.status_code == 200:
            try:
                txt = r.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                return " ".join(txt.split()), m
            except (KeyError, IndexError):
                last_err = f"{m}: empty answer"
                continue
        last_err = f"{m}: HTTP {r.status_code} {r.text[:160]}"
        if r.status_code == 429:
            time.sleep(20)
    raise RuntimeError(last_err)


def run(dry_run: bool = False) -> int:
    cfg = load_config()
    gcfg = cfg.get("gemini", {})
    projects = read_table("projects")
    news, scans = read_table("news"), read_table("ai_scan")
    policy = {r["country"]: r for r in read_table("policy")}
    load_env()
    key = os.environ.get("GEMINI_API_KEY", "").strip()
    use_ai = bool(key) and not dry_run and gcfg.get("enabled", True)
    done_ai, err = 0, ""
    for p in projects:
        rule = rule_action(p, news, scans, policy, cfg)
        p["suggested_action"], p["suggested_by"] = rule, f"Rules {today()}"
        if use_ai and done_ai < gcfg.get("max_per_run", 25) and not err:
            facts = "\n".join(f"- {k}: {v}" for k, v in p.items()
                              if v and k not in ("suggested_action", "suggested_by", "news_query", "ai_query"))
            nw = "\n".join(f"- {n['title']} ({n['publisher']}, {n['published']})"
                           for n in _recent(news, p["id"], "date_found", 60)[:3]) or "none"
            ai = next((s["summary"][:900] for s in sorted(scans, key=lambda s: s["date"], reverse=True)
                       if s["item_id"] == p["id"]), "none")
            pol = policy.get(p["country"], {})
            prompt = PROMPT.format(facts=facts, news=nw, ai=ai, rule=rule,
                                   policy=f"{pol.get('rag','')} - {pol.get('policy_change','')}")
            try:
                txt, model = gemini(prompt, gcfg.get("models", ["gemini-flash-latest"]), key)
                p["suggested_action"], p["suggested_by"] = txt, f"Gemini {today()}"
                done_ai += 1
                time.sleep(gcfg.get("pause_seconds", 5))
            except (RuntimeError, requests.RequestException) as e:
                err = str(e)
    write_table("projects", projects)
    detail = f"{len(projects)} suggestions; Gemini rewrote {done_ai}"
    if not key:
        detail += " (no GEMINI_API_KEY: rules only)"
    if err:
        detail += f"; Gemini stopped: {err[:200]}"
    log("next_actions", "dry-run" if dry_run else "ok", detail)
    return len(projects)


if __name__ == "__main__":
    a = argparse.ArgumentParser(description=__doc__)
    a.add_argument("--dry-run", action="store_true")
    run(a.parse_args().dry_run)
