"""Sheet layouts shared by build_tracker.py (write) and sync_from_excel.py (read).

Each column: (key, header label, width, kind)
  kind: text | num | date | url | list:<name> | formula (computed in Excel, not synced)
Keys starting with "_" are formula columns and are never read back.
"""

HEADER_ROW = 4  # rows 1-3 hold the sheet title, purpose line and key

PIPELINE = [
    ("id", "ID", 9, "text"),
    ("project", "Project / campus", 30, "text"),
    ("developer", "Developer / operator", 22, "text"),
    ("country", "Country", 13, "list:countries"),
    ("city_region", "City / region", 20, "text"),
    ("capacity_mw", "Capacity (MW)", 11, "num"),
    ("capacity_basis", "MW basis", 10, "list:basis"),
    ("investment_eur_m", "Investment (€m)", 12, "num"),
    ("project_type", "Type", 14, "list:types"),
    ("stage", "Stage", 17, "list:stages"),
    ("_priority", "Priority score", 10, "formula"),
    ("_tier", "Priority", 10, "formula"),
    ("announced", "Announced", 11, "text"),
    ("permit", "Permit / approval", 13, "text"),
    ("construction_start", "Construction start", 13, "text"),
    ("target_online", "Target online", 12, "text"),
    ("ethos_fit", "Ethos service fit", 30, "text"),
    ("ethos_relationship", "Ethos relationship", 15, "list:relationship"),
    ("design_team", "Known design team / contractors", 28, "text"),
    ("next_action", "Next action (agreed)", 32, "text"),
    ("suggested_action", "Suggested next action (auto)", 40, "auto"),
    ("suggested_by", "Suggested by", 10, "auto"),
    ("owner", "Owner", 12, "text"),
    ("confidence", "Confidence", 11, "list:confidence"),
    ("source_ids", "Source IDs", 12, "text"),
    ("notes", "Notes & assumptions", 40, "text"),
    ("last_checked", "Last checked", 12, "date"),
    ("_age", "Days since check", 10, "formula"),
]

ACCOUNTS = [
    ("id", "ID", 9, "text"),
    ("account", "Account", 22, "text"),
    ("hq", "HQ", 16, "text"),
    ("ethos_relationship", "Ethos relationship", 15, "list:relationship"),
    ("_projects", "Projects tracked", 10, "formula"),
    ("_mw", "MW tracked", 10, "formula"),
    ("european_pipeline", "European pipeline (public)", 44, "text"),
    ("why_now", "Why now", 44, "text"),
    ("client_partner", "Ethos Client Partner", 18, "text"),
    ("linked_projects", "Linked project IDs", 14, "text"),
    ("next_step", "Next step", 36, "text"),
    ("source_ids", "Source IDs", 12, "text"),
    ("last_checked", "Last checked", 12, "date"),
]

MARKETS = [
    ("country", "Country", 14, "text"),
    ("pipeline_score", "Pipeline (1-5)", 10, "num"),
    ("capex_score", "Investment (1-5)", 10, "num"),
    ("grid_score", "Grid access (1-5)", 10, "num"),
    ("policy_score", "Policy (1-5)", 10, "num"),
    ("fit_score", "Ethos fit (1-5)", 10, "num"),
    ("_total", "Total /25", 9, "formula"),
    ("_rank", "Rank", 7, "formula"),
    ("_projects", "Projects tracked", 10, "formula"),
    ("_mw", "MW tracked", 10, "formula"),
    ("evidence", "Evidence behind the scores", 70, "text"),
    ("source_ids", "Source IDs", 12, "text"),
    ("last_checked", "Last checked", 12, "date"),
]

POLICY = [
    ("country", "Country", 13, "text"),
    ("rag", "Risk", 8, "list:rag"),
    ("grid_lead_time", "Grid connection (large loads)", 38, "text"),
    ("policy_change", "Latest rule change (2025-26)", 48, "text"),
    ("policy_date", "In force / dated", 13, "date"),
    ("next_milestone", "Next milestone to watch", 34, "text"),
    ("milestone_date", "Milestone date", 12, "date"),
    ("rag_reason", "Why this rating", 36, "text"),
    ("ethos_opening", "Commercial opening for Ethos", 40, "text"),
    ("source_ids", "Source IDs", 12, "text"),
    ("last_checked", "Last checked", 12, "date"),
]

SOURCES = [
    ("id", "ID", 9, "text"),
    ("title", "Title", 52, "text"),
    ("publisher", "Publisher", 22, "text"),
    ("type", "P / S", 6, "list:ps"),
    ("published", "Published", 12, "text"),
    ("url", "Link", 40, "url"),
    ("last_checked", "Last checked", 12, "date"),
    ("_age", "Days since check", 10, "formula"),
    ("link_status", "Link status", 11, "text"),
    ("http_status", "HTTP", 7, "text"),
    ("used_in", "Used in", 22, "text"),
]

CHANGES = [
    ("date", "Date", 11, "date"),
    ("item_type", "Area", 10, "text"),
    ("item_id", "ID", 8, "text"),
    ("item_name", "Item", 30, "text"),
    ("field", "What changed", 16, "text"),
    ("old_value", "Before", 22, "text"),
    ("new_value", "After", 22, "text"),
    ("impact", "Why it matters", 38, "text"),
    ("source", "Source", 12, "text"),
    ("logged_by", "Logged by", 10, "text"),
]

GAPS = [
    ("id", "ID", 9, "text"),
    ("gap", "What we don't know yet", 50, "text"),
    ("why_it_matters", "Why it matters", 40, "text"),
    ("how_to_close", "How to close it", 40, "text"),
    ("owner", "Owner", 12, "text"),
    ("status", "Status", 10, "list:gapstatus"),
    ("opened", "Opened", 11, "date"),
]

NEWS = [
    ("date_found", "Found", 11, "date"),
    ("item_id", "ID", 9, "text"),
    ("item_name", "Project / account", 26, "text"),
    ("title", "Headline", 60, "text"),
    ("publisher", "Publisher", 18, "text"),
    ("published", "Published", 14, "text"),
    ("url", "Link", 34, "url"),
]

TENDERS = [
    ("date_found", "Found", 11, "date"),
    ("source", "Source", 7, "text"),
    ("country", "Country", 10, "text"),
    ("title", "Tender title", 60, "text"),
    ("buyer", "Buyer", 30, "text"),
    ("published", "Published", 12, "date"),
    ("deadline", "Deadline", 12, "date"),
    ("notice_type", "Notice type", 18, "text"),
    ("matched", "Matched", 16, "text"),
    ("url", "Link", 30, "url"),
]

AI_SCAN = [
    ("date", "Scanned", 11, "date"),
    ("item_id", "ID", 9, "text"),
    ("item_name", "Project", 26, "text"),
    ("review_status", "Review", 13, "list:review"),
    ("flags", "Possible changes spotted", 44, "text"),
    ("summary", "Google AI Mode answer (unverified)", 80, "text"),
    ("references", "References given", 50, "text"),
]

GRID_ES = [
    ("month", "Month", 10, "text"),
    ("renewable_share_pct", "Renewable share (%)", 14, "num"),
    ("_gap", "Gap to 90% (pts)", 12, "formula"),
    ("renewable_gwh", "Renewable (GWh)", 13, "num"),
    ("total_gwh", "Total generation (GWh)", 15, "num"),
    ("fetched", "Fetched", 20, "text"),
]

LISTS = {
    "countries": ["Germany", "United Kingdom", "Spain", "France", "Italy", "Sweden",
                  "Ireland", "Norway", "Finland", "Poland", "Portugal", "Netherlands",
                  "Denmark", "Greece"],
    "basis": ["IT", "Grid", "Unclear"],
    "types": ["New build", "Expansion", "Retrofit", "Energy-BESS"],
    "stages": ["Announced", "Permitting", "Permitted", "Under construction",
               "Operational", "Cancelled / paused"],
    "relationship": ["Existing client", "Target", "Competitor-held"],
    "confidence": ["High", "Medium", "Low"],
    "rag": ["Red", "Amber", "Green"],
    "ps": ["P", "S"],
    "gapstatus": ["Open", "In progress", "Closed"],
    "review": ["Needs review", "Accepted", "Dismissed", "No action"],
}

# Priority scoring (transparent, editable on the Lists tab)
STAGE_POINTS = {"Announced": 4, "Permitting": 5, "Permitted": 4,
                "Under construction": 2, "Operational": 0, "Cancelled / paused": 0}
REL_POINTS = {"Existing client": 3, "Target": 2, "Competitor-held": 1}
CONF_POINTS = {"High": 1, "Medium": 0.5, "Low": 0}

# sheet name -> (table key, layout)
SYNCED = {
    "Pipeline": ("projects", PIPELINE),
    "Accounts": ("accounts", ACCOUNTS),
    "Markets": ("markets", MARKETS),
    "Grid & Policy": ("policy", POLICY),
    "Sources": ("sources", SOURCES),
    "Changes": ("changes", CHANGES),
    "Gaps": ("gaps", GAPS),
    "AI Scan": ("ai_scan", AI_SCAN),
}
