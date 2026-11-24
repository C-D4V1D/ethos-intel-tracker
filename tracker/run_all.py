"""One command to refresh everything, in a safe order.

  python -m tracker.run_all            # live run (needs SERPAPI_KEY)
  python -m tracker.run_all --dry-run  # offline test with sample data
  python -m tracker.run_all --no-serp  # live, but skip SerpApi (no key / save credits)

Steps
 1. sync_from_excel   - keep hand edits made in the workbook
 2. check_sources     - is every source link still alive? stamp the date
 3. refresh_grid_data - Spain renewable share from Red Electrica
 4. news_watch        - new headlines per project (SerpApi Google News)
 5. ai_mode_scan      - Google AI Mode status check per project (SerpApi)
 6. detect_changes    - log what moved since last run
 7. build_tracker     - rebuild the formatted Excel workbook
A failed step is logged and the run continues, so one broken site never blocks
the weekly update.
"""
from __future__ import annotations

import argparse
import traceback

from . import (ai_mode_scan, build_tracker, check_sources, detect_changes,
               news_watch, refresh_grid_data, sync_from_excel)
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
    a = p.parse_args()
    dry = a.dry_run
    if dry:
        sandbox()
    log("run_all", "start", "dry-run" if dry else "live")
    step("sync_from_excel", sync_from_excel.run)
    if not a.skip_links:
        step("check_sources", check_sources.run, dry)
    step("refresh_grid_data", refresh_grid_data.run, dry)
    if dry or (not a.no_serp and serpapi_key()):
        step("news_watch", news_watch.run, dry)
        step("ai_mode_scan", ai_mode_scan.run, dry)
    else:
        log("serpapi", "skipped", "no SERPAPI_KEY set or --no-serp used")
    step("detect_changes", detect_changes.run, dry)
    step("build_tracker", build_tracker.run)
    log("run_all", "done", "workbook rebuilt in output/")


if __name__ == "__main__":
    main()
