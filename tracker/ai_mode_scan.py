"""Ask Google AI Mode (via SerpApi) for the latest status of each tracked project.

The answer and its references go to data/ai_scan.csv with automatic flags where
the answer seems to disagree with the tracker (stage, MW, dates). Nothing in the
pipeline is overwritten: a person reviews each flag and, if it is right, updates
the project and its source. This keeps the tracker's "verified sources only" rule.
"""
from __future__ import annotations

import argparse
import re

from .common import load_config, log, read_table, today, write_table
from .serp import BudgetExceeded, Serp, pick_items

MW_RE = re.compile(r"(\d{1,4}(?:[.,]\d+)?)\s?(?:MW|megawatts?)\b", re.I)
YEAR_RE = re.compile(r"\b(202[4-9]|203[0-5])\b")


def build_query(row: dict) -> str:
    if row.get("ai_query"):
        return row["ai_query"]
    place = ", ".join(x for x in (row.get("city_region"), row.get("country")) if x)
    return (f"What is the latest status of the {row['project']} data centre project by "
            f"{row['developer']} in {place}? Give capacity in MW, planning or permit status, "
            f"construction start and expected opening, with dates and sources.")


def flatten(blocks: list) -> str:
    parts = []
    for b in blocks or []:
        if b.get("snippet"):
            parts.append(b["snippet"])
        for li in b.get("list") or []:
            if li.get("snippet"):
                parts.append("- " + li["snippet"])
        if b.get("text_blocks"):
            parts.append(flatten(b["text_blocks"]))
    return "\n".join(p for p in parts if p)


def make_flags(text: str, row: dict, stage_kw: dict) -> list[str]:
    flags, low = [], text.lower()
    hits = [st for st, words in stage_kw.items() if any(w in low for w in words)]
    cur = row.get("stage", "")
    for st in hits:
        if st != cur and not (cur == "Operational" and st == "Under construction"):
            flags.append(f"Stage? answer mentions '{st}' (tracker: {cur or 'blank'})")
    try:
        mw = float(row.get("capacity_mw") or 0)
    except ValueError:
        mw = 0
    found = {float(m.replace(",", ".")) for m in MW_RE.findall(text)}
    if found and mw and not any(abs(f - mw) <= 0.1 * mw for f in found):
        flags.append(f"MW? answer mentions {', '.join(f'{f:g}' for f in sorted(found))} MW (tracker: {mw:g})")
    tgt = (row.get("target_online") or "")[:4]
    years = sorted(set(YEAR_RE.findall(text)))
    if tgt.isdigit() and years and tgt not in years:
        flags.append(f"Dates? answer mentions {', '.join(years)} (tracker target: {tgt})")
    return flags


def run(dry_run: bool = False) -> int:
    cfg = load_config()
    scfg, stage_kw = cfg["serpapi"], cfg["ai_scan"]["stage_keywords"]
    serp = Serp(dry_run=dry_run)
    scans = read_table("ai_scan")
    last = {}
    for s in scans:
        last[s["item_id"]] = max(last.get(s["item_id"], ""), s["date"])
    items = pick_items("projects", scfg["ai_mode_per_run"], scfg["rescan_after_days"], last)
    done, flagged = 0, 0
    for row in items:
        q = build_query(row)
        try:
            data = serp.search("google_ai_mode", q)
        except BudgetExceeded as e:
            log("ai_mode_scan", "stopped", str(e))
            break
        text = data.get("reconstructed_markdown") or flatten(data.get("text_blocks"))
        refs = data.get("references") or []
        ref_txt = " | ".join(f"{r.get('source') or r.get('title', '')}: {r.get('link', '')}"
                             for r in refs[:6])
        flags = make_flags(text, row, stage_kw)
        flagged += bool(flags)
        scans.append({
            "date": today(), "item_id": row["id"], "item_name": row["project"],
            "query": q, "summary": text[:1800], "references": ref_txt,
            "flags": "; ".join(flags) or "No conflict spotted",
            "review_status": "Needs review" if flags else "No action",
        })
        done += 1
    scans.sort(key=lambda s: s["date"], reverse=True)
    write_table("ai_scan", scans[:300])
    log("ai_mode_scan", "dry-run" if dry_run else "ok",
        f"{done} projects scanned, {flagged} flagged for review", 0 if dry_run else serp.spent)
    return done


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    run(p.parse_args().dry_run)
