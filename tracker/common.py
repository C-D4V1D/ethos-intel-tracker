"""Shared helpers: paths, config, CSV tables, run log.

The CSV files in data/ are the single source of truth. The Excel workbook is
rebuilt from them on every run (build_tracker.py), and edits made in Excel are
pulled back first (sync_from_excel.py), so nothing typed by hand is lost.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import os
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUTPUT = ROOT / "output"
FIXTURES = ROOT / "tests" / "fixtures"
WORKBOOK_NAME = "Ethos_European_DC_Intelligence_Tracker.xlsx"

# Column order for every table. Keep in sync with build_tracker.py.
SCHEMAS: dict[str, list[str]] = {
    "projects": [
        "id", "project", "developer", "country", "city_region", "capacity_mw",
        "capacity_basis", "investment_eur_m", "project_type", "stage",
        "announced", "permit", "construction_start", "target_online",
        "ethos_fit", "ethos_relationship", "design_team", "next_action",
        "owner", "confidence", "source_ids", "notes", "last_checked",
        "news_query", "ai_query", "suggested_action", "suggested_by",
    ],
    "accounts": [
        "id", "account", "hq", "european_pipeline", "why_now",
        "ethos_relationship", "client_partner", "linked_projects",
        "next_step", "source_ids", "last_checked",
    ],
    "markets": [
        "country", "status", "status_reason", "pipeline_score", "capex_score", "grid_score",
        "policy_score", "fit_score", "evidence", "source_ids", "last_checked",
    ],
    "policy": [
        "country", "grid_lead_time", "policy_change", "policy_date",
        "next_milestone", "milestone_date", "rag", "rag_reason",
        "ethos_opening", "source_ids", "last_checked",
    ],
    "sources": [
        "id", "title", "publisher", "type", "published", "url",
        "last_checked", "http_status", "link_status", "used_in",
    ],
    "changes": [
        "date", "item_type", "item_id", "item_name", "field", "old_value",
        "new_value", "impact", "source", "logged_by",
    ],
    "gaps": [
        "id", "gap", "why_it_matters", "how_to_close", "owner", "status",
        "opened",
    ],
    "grid_markets": ["country", "month", "renewable_share_pct", "price_eur_mwh", "source", "fetched"],
    "tenders": [
        "date_found", "source", "notice_id", "title", "buyer", "country",
        "published", "deadline", "notice_type", "url", "matched",
    ],
    "news": [
        "date_found", "item_id", "item_name", "title", "publisher",
        "published", "url", "snippet", "query",
    ],
    "ai_scan": [
        "date", "item_id", "item_name", "query", "summary", "references",
        "flags", "review_status",
    ],
    "grid_es": [
        "month", "renewable_share_pct", "renewable_gwh", "total_gwh",
        "fetched",
    ],
    "run_log": ["run_at", "step", "status", "detail", "credits_used"],
}


def today() -> str:
    return dt.date.today().isoformat()


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%d %H:%M UTC")


def load_config() -> dict:
    with open(ROOT / "config.yaml", encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def load_env() -> None:
    """Load .env if python-dotenv is installed (local runs). CI uses secrets."""
    try:
        from dotenv import load_dotenv
        load_dotenv(ROOT / ".env")
    except ImportError:
        pass


def serpapi_key() -> str | None:
    load_env()
    key = os.environ.get("SERPAPI_KEY", "").strip()
    return key or None


def read_table(name: str) -> list[dict]:
    path = DATA / f"{name}.csv"
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as fh:
        rows = list(csv.DictReader(fh))
    cols = SCHEMAS[name]
    return [{c: (r.get(c) or "").strip() for c in cols} for r in rows]


def write_table(name: str, rows: list[dict]) -> None:
    DATA.mkdir(exist_ok=True)
    cols = SCHEMAS[name]
    with open(DATA / f"{name}.csv", "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})


def append_rows(name: str, new_rows: list[dict]) -> None:
    rows = read_table(name)
    rows.extend(new_rows)
    write_table(name, rows)


def log(step: str, status: str, detail: str = "", credits: int = 0) -> None:
    append_rows("run_log", [{
        "run_at": now(), "step": step, "status": status,
        "detail": detail[:500], "credits_used": str(credits),
    }])
    print(f"[{step}] {status}: {detail}")


def credits_used_this_month() -> int:
    month = dt.date.today().strftime("%Y-%m")
    total = 0
    for r in read_table("run_log"):
        if r["run_at"].startswith(month):
            try:
                total += int(r["credits_used"] or 0)
            except ValueError:
                pass
    return total


def fixture(name: str) -> dict:
    with open(FIXTURES / name, encoding="utf-8") as fh:
        return json.load(fh)


def days_since(iso: str) -> int | None:
    try:
        return (dt.date.today() - dt.date.fromisoformat(iso[:10])).days
    except (ValueError, TypeError):
        return None
