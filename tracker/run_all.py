"""One command to refresh everything, in a safe order.

  python -m tracker.run_all            # live run (needs SERPAPI_KEY)
  python -m tracker.run_all --dry-run  # offline test with sample data
  python -m tracker.run_all --no-serp  # live, but skip SerpApi (no key / save credits)

Steps
 1. sync_from_excel   - keep hand edits made in the workbook
 2. check_sources     - is every source link still alive? stamp the date
 3. refresh_grid_data - Spain renewable share from Red Electrica
 4. news_watch        - new headlines per project (free Google News RSS)
 5. tender_watch      - public tenders mentioning data centres (EU TED, UK Find a Tender)
 6. grid_markets      - renewable share and power price for every market (free APIs)
 6b. auto_update      - AI Mode + Gemini update every project and find new ones; values are
                       written only when found on the cited page (needs keys)
 7. next_actions      - suggested next action per project (rules; Gemini if key set)
 8. detect_changes    - log what moved since last run
 9. build_tracker     - rebuild the formatted Excel workbook
A failed step is logged and the run continues, so one broken site never blocks
the weekly update.
"""
from __future__ import annotations

import argparse
import os
import traceback

from . import (auto_update, build_tracker, enrich, grid_markets, onedrive, review, check_sources, detect_changes,
               news_watch, next_actions, refresh_grid_data, sync_from_excel, tender_watch)
from .common import log, serpapi_key


def step(name, fn, *a):
    try:
        return fn(*a)
    except Exception as e:  # noqa: BLE001 - keep the pipeline running
        log(name, "failed", f"{type(e).__name__}: {e}")
        traceback.print_exc()
        return None


def sandbox() -> None:
    """Dry runs work on a throw-away copy of data/ and write to output/dry-run/,
    so sample (fixture) results never reach the real master data or workbook."""
    import shutil
    import sys
    import tempfile
    from pathlib import Path
    from . import common
    tmp = Path(tempfile.mkdtemp(prefix="tracker-dry-"))
    shutil.copytree(common.DATA, tmp / "data", dirs_exist_ok=True)
    out = common.OUTPUT / "dry-run"
    out.mkdir(parents=True, exist_ok=True)
    for mod in list(sys.modules.values()):
        if getattr(mod, "__name__", "").startswith("tracker"):
            if hasattr(mod, "DATA"):
                mod.DATA = tmp / "data"
            if hasattr(mod, "OUTPUT"):
                mod.OUTPUT = out
    print(f"[run_all] dry-run sandbox: data copied to {tmp/'data'}, workbook -> {out}")


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--dry-run", action="store_true", help="offline test with fixtures")
    p.add_argument("--no-serp", action="store_true", help="skip SerpApi steps")
    p.add_argument("--skip-links", action="store_true", help="skip the link checker")
    p.add_argument("--light", action="store_true",
                   help="daily run: pull OneDrive edits, apply approvals, free feeds, rebuild (no SerpApi)")
    p.add_argument("--approvals-only", action="store_true",
                   help="just apply review decisions and rebuild (used when you tick boxes on GitHub)")
    a = p.parse_args()
    dry = a.dry_run
    if dry:
        sandbox()
    if a.light or a.approvals_only:
        os.environ["SKIP_SERP"] = "true"
    log("run_all", "start", "dry-run" if dry else "approvals" if a.approvals_only else "light" if a.light else "full")
    if not dry:
        step("onedrive_pull", onedrive.pull)         # your edits + Approve/Reject from the OneDrive copy
    step("sync_from_excel", sync_from_excel.run)
    step("review", review.run, not dry)               # apply approvals (Excel tab or GitHub tick boxes)
    if not a.approvals_only:
        if not (a.skip_links or a.light):
            step("check_sources", check_sources.run, dry)
        if not a.light:
            step("refresh_grid_data", refresh_grid_data.run, dry)
            step("grid_markets", grid_markets.run, dry)
        step("news_watch", news_watch.run, dry)
        step("tender_watch", tender_watch.run, dry)
        if not (a.no_serp or a.light):
            step("auto_update", auto_update.run, dry)  # AI Mode + Gemini + source check
        step("enrich", enrich.run, dry, not (a.light or a.no_serp))
        step("next_actions", next_actions.run, dry)
        if not dry:
            step("review_issue", review.publish_issue)   # refresh the GitHub tick-box list
    step("detect_changes", detect_changes.run, dry)
    step("build_tracker", build_tracker.run)
    log("run_all", "done", "workbook rebuilt in output/")


if __name__ == "__main__":
    main()
