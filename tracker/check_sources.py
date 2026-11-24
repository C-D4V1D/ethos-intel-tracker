"""Open every source URL, record whether it still works, and stamp the date.

link_status values:
  OK        - page answered 2xx/3xx
  BLOCKED   - site refuses automated checks (401/403/429); open it by hand
  BROKEN    - 404/410: the page is gone; find a replacement
  ERROR     - no answer (timeout / unreachable) or a server error; retried next run
  OK (reader) - the site blocks robots, but the free Jina Reader (r.jina.ai) could
              still fetch the page, so the source is alive
"""
from __future__ import annotations

import argparse
import time

import requests

from .common import load_config, log, read_table, today, write_table

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0 Safari/537.36 EthosIntelTracker/1.0")


def classify(code: int | None) -> str:
    if code is None:
        return "ERROR"  # timeouts are usually temporary or a block on cloud IPs
    if 200 <= code < 400:
        return "OK"
    if code in (401, 403, 429, 999):
        return "BLOCKED"
    if code in (404, 410):
        return "BROKEN"
    return "ERROR"


def check_url(url: str, timeout: int) -> int | None:
    """Check once; if the site timed out, errored or said 404, wait and retry with a
    full GET so one slow or flaky response never marks a good source as broken."""
    import time
    code = _check(url, timeout)
    if code is None or code in (404, 410) or code >= 500:
        time.sleep(5)
        code = _check(url, timeout * 2, force_get=True)
    return code


def _check(url: str, timeout: int, force_get: bool = False) -> int | None:
    headers = {"User-Agent": UA, "Accept-Language": "en,es;q=0.8"}
    try:
        r = None if force_get else requests.head(url, headers=headers, timeout=timeout,
                                                 allow_redirects=True)
        if r is None or r.status_code in (405, 403, 400) or r.status_code >= 500:
            r = requests.get(url, headers=headers, timeout=timeout,
                             allow_redirects=True, stream=True)
            r.close()
        return r.status_code
    except requests.RequestException:
        return None


def jina_alive(url: str, timeout: int) -> bool:
    """Ask the free Jina Reader to fetch the page. True if it returns real page text."""
    import os
    headers = {"Accept": "text/plain", "X-Return-Format": "text"}
    if os.environ.get("JINA_API_KEY"):  # optional: higher rate limit
        headers["Authorization"] = f"Bearer {os.environ['JINA_API_KEY']}"
    try:
        r = requests.get("https://r.jina.ai/" + url, headers=headers, timeout=timeout * 3)
    except requests.RequestException:
        return False
    text = r.text or ""
    bad = ("403 forbidden", "access denied", "404 not found", "page not found", "just a moment",
           "attention required", "target url returned error")
    return r.status_code == 200 and len(text) > 800 and not any(b in text[:1500].lower() for b in bad)


def run(dry_run: bool = False) -> dict:
    cfg = load_config()["source_check"]
    rows = read_table("sources")
    counts = {"OK": 0, "OK (reader)": 0, "BLOCKED": 0, "BROKEN": 0, "ERROR": 0}
    for i, row in enumerate(rows):
        if i >= cfg["max_per_run"]:
            break
        if not row["url"]:
            continue
        if dry_run:
            code = 200
        else:
            code = check_url(row["url"], cfg["timeout_seconds"])
        status = classify(code)
        if status in ("BLOCKED", "ERROR") and not dry_run and cfg.get("jina_fallback", True):
            if jina_alive(row["url"], cfg["timeout_seconds"]):
                status = "OK (reader)"
                time.sleep(3)  # free Jina tier is rate-limited
        counts[status] += 1
        row["http_status"] = str(code or "")
        row["link_status"] = status
        if status.startswith("OK"):
            # A working link counts as re-checked; blocked/broken ones keep
            # their old date so the freshness flag still shows.
            row["last_checked"] = today()
    write_table("sources", rows)
    log("check_sources", "dry-run" if dry_run else "ok",
        ", ".join(f"{k} {v}" for k, v in counts.items()))
    return counts


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    run(p.parse_args().dry_run)
