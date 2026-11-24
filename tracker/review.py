"""Review queue: everything the automation could not confirm on its own waits here for a
one-click decision, then is applied automatically.

Ways to decide (both work, use either):
  * Excel (OneDrive copy or any downloaded copy): 'Review' tab, column 'Your decision'
    -> Approve / Reject. Pulled back automatically on the next daily run.
  * GitHub: the pinned issue "Review queue" lists every pending item with tick boxes
    (works on the GitHub phone app). Ticking a box applies it within a minute or two.

Approve = the proposed value is written to its tab, its source is added to Sources and
the change is logged ("Approved"). Reject = never proposed again.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess

from .common import log, read_table, today, write_table

TABLE = {"Pipeline": "projects", "Grid & Policy": "policy", "Gaps": "gaps",
         "Accounts": "accounts", "Markets": "markets"}
KEYCOL = {"projects": "id", "policy": "country", "gaps": "id", "accounts": "id", "markets": "country"}
ISSUE_TITLE = "Review queue (tick to approve or reject)"


def add(reviews, area, item_id, item_name, field, current, proposed, evidence="", source_url="", reason=""):
    """Queue a decision unless the same proposal is already pending or was rejected before."""
    proposed = str(proposed)
    for r in reviews:
        if r["item_id"] == item_id and r["field"] == field and r["proposed_value"] == proposed \
                and r["decision"] in ("Pending", "Reject", ""):
            return None
    # a newer proposal for the same field replaces an older pending one
    for r in reviews:
        if r["item_id"] == item_id and r["field"] == field and r["decision"] in ("Pending", ""):
            r["decision"], r["decided_on"], r["applied"] = "Superseded", today(), "no"
    n = max([int(r["id"][2:]) for r in reviews if re.fullmatch(r"R-\d+", r["id"])] or [0]) + 1
    row = {"id": f"R-{n:04d}", "created": today(), "area": area, "item_id": item_id,
           "item_name": item_name, "field": field, "current_value": str(current or ""),
           "proposed_value": proposed, "evidence": (evidence or "")[:600], "source_url": source_url or "",
           "reason": reason, "decision": "Pending", "decided_on": "", "applied": ""}
    reviews.append(row)
    return row


def _source(sources, url, used, title=""):
    if not url:
        return ""
    for s in sources:
        if s["url"] == url:
            if used not in s["used_in"]:
                s["used_in"] = (s["used_in"] + "; " + used).strip("; ")
            return s["id"]
    sid = f"SRC{max(int(s['id'][3:]) for s in sources if s['id'][3:].isdigit()) + 1:03d}"
    host = re.sub(r"^www\.", "", re.sub(r"^https?://([^/]+).*", r"\1", url))
    sources.append({"id": sid, "title": title or host, "publisher": host, "type": "S", "published": "",
                    "url": url, "last_checked": today(), "http_status": "", "link_status": "Not checked",
                    "used_in": used})
    return sid


def apply_decisions() -> dict:
    reviews = read_table("reviews")
    todo = [r for r in reviews if r["decision"] in ("Approve", "Reject") and r["applied"] != "yes"]
    if not todo:
        return {"applied": 0, "rejected": 0}
    tables = {t: read_table(t) for t in set(TABLE.values())}
    sources, changes = read_table("sources"), []
    applied = rejected = 0
    for r in todo:
        r["decided_on"] = r["decided_on"] or today()
        if r["decision"] == "Reject":
            r["applied"] = "yes"
            rejected += 1
            continue
        if r["field"] == "(new project)":
            p = json.loads(r["proposed_value"])
            sid = _source(sources, r["source_url"], p["id"])
            p["source_ids"] = sid
            tables["projects"].append(p)
            changes.append({"date": today(), "item_type": "projects", "item_id": p["id"], "item_name": p["project"],
                            "field": "(new item)", "old_value": "", "new_value": "added",
                            "impact": "New project approved from the review queue", "source": sid,
                            "logged_by": "Approved"})
        else:
            t = TABLE.get(r["area"])
            if not t:
                r["applied"] = "no (unknown tab)"
                continue
            row = next((x for x in tables[t] if x[KEYCOL[t]] == r["item_id"]), None)
            if row is None:
                r["applied"] = "no (item not found)"
                continue
            old = row.get(r["field"], "")
            row[r["field"]] = r["proposed_value"]
            if t == "gaps" and r["field"] == "answer":
                row["status"] = "Closed"
            sid = _source(sources, r["source_url"], r["item_id"])
            if sid and "source_ids" in row and sid not in row["source_ids"]:
                row["source_ids"] = (row["source_ids"] + "; " + sid).strip("; ")
            if "last_checked" in row:
                row["last_checked"] = today()
            changes.append({"date": today(), "item_type": t, "item_id": r["item_id"], "item_name": r["item_name"],
                            "field": r["field"], "old_value": old, "new_value": r["proposed_value"],
                            "impact": f"Approved from the review queue ({r['id']})", "source": sid,
                            "logged_by": "Approved"})
        r["applied"] = "yes"
        applied += 1
    for t, rows in tables.items():
        write_table(t, rows)
    write_table("sources", sources)
    if changes:
        write_table("changes", changes + read_table("changes"))
    write_table("reviews", reviews)
    log("review", "ok", f"{applied} approvals applied, {rejected} rejected")
    return {"applied": applied, "rejected": rejected}


# ---------------------------------------------------------------- GitHub issue checklist
def _gh(*args, input_=None):
    return subprocess.run(["gh", *args], capture_output=True, text=True, input=input_)


def _issue_number():
    r = _gh("issue", "list", "--state", "open", "--search", f'in:title "{ISSUE_TITLE}"', "--json", "number,title")
    try:
        for i in json.loads(r.stdout or "[]"):
            if i["title"] == ISSUE_TITLE:
                return i["number"]
    except json.JSONDecodeError:
        pass
    return None


def _line(r):
    short = lambda v: (v[:70] + "…") if len(v) > 70 else v
    if r["field"] == "(new project)":
        p = json.loads(r["proposed_value"])
        what = f"**New project:** {p['project']} — {p['developer']}, {p['city_region']} ({p['country']})"
    else:
        what = (f"**{r['item_name']}** · {r['field']}: `{short(r['current_value']) or 'blank'}` → "
                f"`{short(r['proposed_value'])}`")
    src = f" · [source]({r['source_url']})" if r["source_url"] else ""
    return (f"- [ ] ✅ Approve `{r['id']}` {what}{src}\n"
            f"  - [ ] ❌ Reject `{r['id']}` — _{r['reason'] or 'needs a person'}_")


def read_issue_ticks() -> int:
    """Copy ticked boxes from the issue into reviews.csv."""
    if not os.environ.get("GH_TOKEN"):
        return 0
    n = _issue_number()
    if not n:
        return 0
    body = json.loads(_gh("issue", "view", str(n), "--json", "body").stdout or "{}").get("body", "")
    pairs = re.findall(r"- \[[xX]\] (✅ Approve|❌ Reject) `(R-\d+)`", body)
    kinds = {rid: kind for kind, rid in pairs}  # if both are ticked, Reject (listed second) wins
    if not kinds:
        return 0
    reviews = read_table("reviews")
    changed = 0
    for r in reviews:
        if r["id"] in kinds and r["decision"] in ("Pending", ""):
            r["decision"] = "Approve" if "Approve" in kinds[r["id"]] else "Reject"
            r["decided_on"] = today()
            changed += 1
    write_table("reviews", reviews)
    return changed


def publish_issue() -> None:
    """Rewrite the issue so it lists exactly the items still pending."""
    if not os.environ.get("GH_TOKEN"):
        return
    pending = [r for r in read_table("reviews") if r["decision"] in ("Pending", "")]
    head = (f"**{len(pending)} item(s) waiting.** Tick ✅ to approve or ❌ to reject; changes are applied "
            "automatically within a few minutes and the workbook is rebuilt. You can also decide in the "
            "workbook's *Review* tab.\n\n")
    groups = {}
    for r in pending:
        groups.setdefault(r["area"] if r["field"] != "(new project)" else "New projects", []).append(r)
    body = head + "\n\n".join(f"### {g}\n" + "\n".join(_line(r) for r in rs[:60]) for g, rs in groups.items())
    if not pending:
        body = head + "_Nothing to review._"
    n = _issue_number()
    if n:
        _gh("issue", "edit", str(n), "--body-file", "-", input_=body[:65000])
    else:
        _gh("issue", "create", "--title", ISSUE_TITLE, "--body-file", "-", input_=body[:65000])
        n = _issue_number()
        if n:
            _gh("issue", "pin", str(n))


def run(from_issue: bool = True) -> None:
    ticked = read_issue_ticks() if from_issue else 0
    res = apply_decisions()
    publish_issue()
    if ticked:
        log("review", "ok", f"{ticked} decisions read from the GitHub review issue")
    return res


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--no-issue", action="store_true")
    run(not ap.parse_args().no_issue)
