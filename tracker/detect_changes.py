"""Compare today's data with the last run and log every change.

Watches key fields on projects, policy and sources. Anything that moved is
written to data/changes.csv (logged_by = "auto"), which feeds the workbook's
Changes tab and the dashboard's "latest changes" box.
"""
from __future__ import annotations

import argparse
import json

from .common import DATA, append_rows, log, read_table, today

WATCH = {
    "projects": ("id", "project", ["stage", "capacity_mw", "investment_eur_m",
                                    "target_online", "ethos_relationship", "design_team"]),
    "policy": ("country", "country", ["policy_change", "next_milestone",
                                      "milestone_date", "rag"]),
    "sources": ("id", "title", ["link_status"]),
}
IMPACT = {
    "stage": "Check consultant appointment window",
    "capacity_mw": "Re-size opportunity",
    "investment_eur_m": "Re-size opportunity",
    "target_online": "Timeline moved",
    "ethos_relationship": "Account status changed",
    "design_team": "Competitor or partner named",
    "policy_change": "New rule: review watchlist and affected projects",
    "next_milestone": "New date to watch",
    "milestone_date": "Date moved",
    "rag": "Risk rating changed",
    "link_status": "Source needs attention",
}


def snapshot() -> dict:
    snap = {}
    for table, (key, _name, fields) in WATCH.items():
        snap[table] = {r[key]: {f: r[f] for f in fields} | {"_name": r[_name]}
                       for r in read_table(table) if r[key]}
    return snap


def run(dry_run: bool = False) -> int:
    path = DATA / ".snapshot.json"
    cur = snapshot()
    if not path.exists():
        path.write_text(json.dumps(cur, indent=1))
        log("detect_changes", "baseline", "first run: snapshot saved")
        return 0
    old = json.loads(path.read_text())
    changes = []
    for table, (_key, _name, fields) in WATCH.items():
        o, n = old.get(table, {}), cur.get(table, {})
        for k in n.keys() - o.keys():
            changes.append({"item_type": table, "item_id": k, "item_name": n[k]["_name"],
                            "field": "(new item)", "old_value": "", "new_value": "added",
                            "impact": "New item to review"})
        for k in o.keys() - n.keys():
            changes.append({"item_type": table, "item_id": k, "item_name": o[k]["_name"],
                            "field": "(removed)", "old_value": "present", "new_value": "",
                            "impact": "Item removed"})
        for k in n.keys() & o.keys():
            for f in fields:
                if (o[k].get(f) or "") != (n[k].get(f) or ""):
                    if table == "sources" and n[k].get(f) == "OK":
                        continue  # a link recovering is not news
                    changes.append({"item_type": table, "item_id": k, "item_name": n[k]["_name"],
                                    "field": f, "old_value": o[k].get(f, ""),
                                    "new_value": n[k].get(f, ""), "impact": IMPACT.get(f, "")})
    for c in changes:
        c.update({"date": today(), "source": "", "logged_by": "auto"})
    if changes and not dry_run:
        append_rows("changes", changes)
    if not dry_run:
        path.write_text(json.dumps(cur, indent=1))
    log("detect_changes", "dry-run" if dry_run else "ok", f"{len(changes)} changes logged")
    return len(changes)


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    run(p.parse_args().dry_run)
