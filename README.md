# European Data Centre Pipeline & Grid Intelligence Tracker

A structured tracker of data-centre projects, accounts, a grid and policy watchlist, and the sources behind them. It covers the European markets where Ethos Engineering can win work. It refreshes itself every week, so the numbers and links stay current without manual effort.

Prepared by David Chukwuemeka.

## Get the latest workbook

[Download the latest Excel workbook](https://github.com/C-D4V1D/ethos-intel-tracker/releases/latest/download/Ethos_European_DC_Intelligence_Tracker.xlsx)

It is rebuilt every Monday.

## What updates automatically

| Step | What it does | Source |
|---|---|---|
| Link checker | Checks that every source link still works and stamps the date checked | The source URLs listed in the tracker |
| Spain grid data | Updates Spain's monthly renewable share of electricity (24 months of history) | Red Eléctrica REData API (free, no key) |
| Project news | Finds new headlines for each project | SerpApi, Google News |
| AI Mode status scan | Asks Google AI Mode whether a project's status has changed and flags possible changes for a person to review. It never overwrites data. | SerpApi, Google AI Mode |
| Change log | Records what moved since the last run | The tracker's own data |
| Workbook rebuild | Rebuilds the formatted Excel workbook and publishes it | The tracker's own data |

AI Mode answers are leads to verify, not facts. Check them against a primary source before relying on them.

A step that fails is logged and the run carries on, so one broken website never blocks the weekly update.

## How it works

```text
data/*.csv  (source of truth)
     |
     v
tracker scripts (Python)
     |
     v
output/ Excel workbook
     |
     v
GitHub Release "latest"
```

**Making hand edits.** Edit the workbook as normal. Then, on GitHub, open the `inbox/` folder and choose **Add file > Upload files** to upload your edited workbook. The weekly workflow syncs your edits back into the CSV files, then removes the uploaded file. Formula columns are ignored.

## Set up the SerpApi key (one time)

1. In this repository, go to **Settings > Secrets and variables > Actions > New repository secret**.
2. Name it `SERPAPI_KEY` and paste your key as the value.

Never paste the key into chat, code or the workbook.

The SerpApi free plan allows 250 searches a month. The config caps usage at 230, and cached searches don't count. Without a key, the run skips the SerpApi steps and does everything else.

## Run it now

Go to the **Actions** tab, choose **Weekly refresh**, then **Run workflow**. There is an option to skip SerpApi and save search credits.

The schedule is Mondays at 06:17 UTC. The workflow also re-runs when an edited workbook is uploaded to `inbox/`.

## Run locally (optional)

```bash
pip install -r requirements.txt
cp .env.example .env        # then add your SERPAPI_KEY
python -m tracker.run_all   # live run
```

Options:

- `--dry-run`: offline test with sample data. It works on a throw-away copy of the data and writes to `output/dry-run/`, so it never touches real data.
- `--no-serp`: live run, but skip the SerpApi steps.
- `--skip-links`: skip the link checker.

## Repository layout

- `data/`: the CSV tables that are the source of truth (projects, accounts, markets, policy, sources, changes, gaps, news, AI scan results, Spain grid data, run log)
- `tracker/`: the Python scripts, one per step, plus `run_all.py` which runs them in order
- `config.yaml`: settings (search budget, freshness thresholds, timeouts, AI scan keywords). No code changes needed.
- `.github/workflows/weekly-refresh.yml`: the weekly automation
- `output/`: the generated Excel workbook (`output/dry-run/` holds dry-run output)
- `inbox/`: where edited workbooks are uploaded (created when first used)
- `tests/`: test fixtures
- `requirements.txt`, `.env.example`: dependencies and local key template

## Data rules

- Every row has source IDs and a last-checked date.
- Sources are marked P (primary) or S (secondary).
- Freshness: amber if not checked for more than 90 days, red after 180 days (both set in `config.yaml`).
- Confidence is High, Medium or Low.
- Megawatt (MW) figures state their basis: IT, Grid or Unclear.
- Scoring weights live on the Lists sheet of the workbook.


## Free services used

| Service | What it does here | Key needed? |
|---|---|---|
| Google News RSS (Bing News RSS as backup) | Weekly headlines for every project | No |
| EU TED Search API v3 | Public tenders mentioning data centres, limited to construction / engineering / M&E categories (CPV) | No |
| UK Find a Tender (OCDS API) | UK public tenders mentioning data centres | No |
| Jina Reader (r.jina.ai) | Second attempt for sources whose sites block the link checker; a page it can read is marked `OK (reader)` | No (optional `JINA_API_KEY` for higher limits) |
| Red Eléctrica REData API | Spain monthly renewable share | No |
| SerpApi Google AI Mode | Weekly AI status check per project (review tab, never overwrites data) | `SERPAPI_KEY` secret |
| Google Gemini API, free tier | Rewrites each project's suggested next action from the row, news and AI answer | `GEMINI_API_KEY` secret |

### Suggested next actions
Every project gets a **Suggested next action (auto)** each week. A rule engine writes it from the stage,
Ethos relationship, design-team status, freshness, recent headlines, AI Scan flags and the market's policy
rating. With `GEMINI_API_KEY` set, Gemini rewrites it (labelled "Gemini", unverified). The human
**Next action (agreed)** column is never changed by automation.

### Adding the two keys (one time)
GitHub repo > Settings > Secrets and variables > Actions > New repository secret:
`SERPAPI_KEY` (from serpapi.com > Dashboard > API key) and `GEMINI_API_KEY`
(from aistudio.google.com > Get API key). Never paste keys into chat, code or the workbook.


## AI auto-update (how the AI keeps the tracker current)
Every week, for every project, Google AI Mode (via SerpApi) is asked for the latest facts. Gemini turns the
answer into fields with the exact supporting quote and the cited link. The link is opened and the value
must appear on that page, next to its unit or currency, and the page must be about the same project.
Only then is it written, with the page added to Sources and the change logged as **AI (verified)**.

Safety rules: blanks can be filled and stages can move forward; existing figures, dates and design teams are
never overwritten automatically (they go to the AI Scan tab as *Needs review*); construction start is only
filled once a project is under construction; very small investment figures on large projects are rejected.

Each week the same engine also searches every focus market for **new projects** (AI Mode + free market news).
A new project is added (confidence *Low*, owner *Auto*) only when a source page names the developer and place;
programmes, portfolios and multi-site plans go to review instead.


## Approving with one click
Anything the automation cannot confirm on its own goes to a review queue. Decide in either place:
* **Workbook, Review tab** (OneDrive copy): set *Your decision* to Approve or Reject. Applied on the next daily run.
* **GitHub, pinned issue "Review queue"**: tick ✅ or ❌ (works in the GitHub phone app). Applied within minutes.

Approve writes the value to its tab, adds the source and logs the change; Reject means it is never proposed again.

## Live copy on OneDrive (view-only link for anyone)
1. Register a free Microsoft app: entra.microsoft.com > App registrations > New registration; name it
   "Ethos tracker"; account types: "Accounts in any organizational directory and personal Microsoft accounts";
   Register. Then Authentication > "Allow public client flows" = Yes > Save. Copy the *Application (client) ID*.
2. GitHub secret `MS_CLIENT_ID` = that ID.
3. GitHub secret `GH_PAT` = a fine-grained personal access token (github.com/settings/personal-access-tokens/new,
   only this repository, permission "Secrets: Read and write") so the workflow can keep the Microsoft sign-in fresh.
4. Actions tab > "Connect OneDrive (one time)" > Run workflow; open the log, go to microsoft.com/devicelogin,
   enter the code and sign in. The workbook is uploaded and the view-only link printed (also shown on the Read me tab).

## Schedule
* Daily 07:17 UTC: OneDrive edits and approvals, free news and tenders with AI summaries, rebuild, OneDrive update.
* Monday 06:17 UTC: everything, including SerpApi (AI Mode checks, new projects, Google News/Jobs/Events) and the
  weekly AI briefing. A second SerpApi key can be added as secret `SERPAPI_KEY_2`; it is used when the first runs out.
