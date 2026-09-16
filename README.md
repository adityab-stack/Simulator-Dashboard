# Snitch Control Room

Internal BI dashboard for Snitch. Static frontend (no backend server) — data is
pre-computed nightly and committed as JSON files; the browser just fetches and
filters them. Hosted for free on GitHub Pages.

## How it works

```
Snowflake + Google Sheets  --(build_data.py)-->  data/*.json  --(git push)-->  GitHub Pages
```

1. **`build_data.py`** connects to Snowflake (primary data source: sales,
   inventory, ageing, returns, clicks, etc.) and a couple of Google Sheets
   tabs (Inventory, Pipeline — via the Sheets API), runs a large set of
   queries/transforms, and writes the results into `data/*.json`.
2. **GitHub Actions** (`.github/workflows/refresh-data.yml`) runs this script
   on a daily cron (04:30 UTC / 10:00 AM IST) and on manual dispatch. It
   commits and pushes any changed `data/*.json` files back to the repo as
   `github-actions[bot]`.
3. **GitHub Pages** redeploys automatically whenever the branch updates —
   no server, no downtime, nothing to restart.
4. **`index.html`** is a single-page vanilla JS app. On load it fetches the
   JSON files it needs and renders ~15 tabs (Dashboard, Sales Tracking, Exec
   Summary, Returns, Analytics, Pareto, Quartile, Live SKU Groups, Clicks,
   Ageing, Store Cut Size, Sales vs Inwards, Repeat Evaluation, Store
   Returns) via a single `render()` dispatcher that toggles tab visibility.

## Data files (`data/`)

Written by `build_data.py` on every run. Approximate current sizes:

| File | ~Size | Notes |
|---|---|---|
| `sales_by_sku.json` | ~82 MB | **Largest file — watch this.** Rolling window (`ROLLING_WINDOW_DAYS`), currently 400 days. Already past half of GitHub's 100MB hard file-size limit. Do not shrink the window to fix size — see Known Issues below for the actual plan. |
| `sales.json` | ~26 MB | |
| `sku_meta.json` | ~19 MB | |
| `clicks.json` | ~21 MB | |
| `returns_actual.json` | ~23 MB | |
| `repeat_eval.json` | ~22 MB | |
| `livesku.json` | ~12 MB | |
| `meta.json` | ~3 MB | |
| `storereturns.json` | ~6 MB | |
| `pipeline.json`, `salesvsinwards.json`, `storecutsize.json`, `targets.json` | <2 MB each | |
| `inventory.json` | should be non-trivial | **If this writes as 0.0 MB, something's wrong** — see Known Issues. |

Total is currently ~215 MB across 14 files.

## Snowflake authentication (key-pair, not password/MFA)

`build_data.py` connects to Snowflake using **RSA key-pair authentication**,
not a password. This is deliberate: Snowflake now enforces MFA (TOTP) on
password logins for this account, and a TOTP code can't be supplied by an
unattended script or a scheduled GitHub Action — key-pair auth is Snowflake's
own recommended method for exactly this case, and doesn't trigger MFA at all.

**Local setup** (`snowflake_credentials.json`, gitignored, never committed):

```json
{
  "SNOWFLAKE_ACCOUNT": "...",
  "SNOWFLAKE_USER": "...",
  "SNOWFLAKE_PRIVATE_KEY_PATH": "C:\\path\\to\\snowflake_rsa_key.p8",
  "SNOWFLAKE_WAREHOUSE": "...",
  "SNOWFLAKE_DATABASE": "...",
  "SNOWFLAKE_SCHEMA": "..."
}
```

**GitHub Actions setup**: repo secret `SNOWFLAKE_PRIVATE_KEY` holds the
**base64-encoded** contents of the `.p8` private key file (base64, not raw
PEM text — raw multi-line PEM pasted into GitHub's secret box has repeatedly
had its newlines mangled in transit, which breaks parsing; base64 has no
newlines to lose). Generate it with:

```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("snowflake_rsa_key.p8")) | Set-Clipboard
```

then paste directly into the secret's value box. `_load_snowflake_private_key()`
in `build_data.py` auto-detects base64 vs. raw PEM, so no code changes are
needed if you ever regenerate the key.

**If you ever need to rotate the key**: generate a new pair
(`gen_snowflake_key.py` in this repo does this without needing OpenSSL
installed), attach the new public key via `ALTER USER <user> SET
RSA_PUBLIC_KEY='...'` in a Snowflake worksheet, update both the local
`snowflake_credentials.json` path and the `SNOWFLAKE_PRIVATE_KEY` GitHub
secret. The key has no built-in expiry — this is only needed if it's
compromised or you want to rotate it as routine hygiene, not on any fixed
schedule.

**Never commit**: `snowflake_rsa_key.p8` (private key), `snowflake_credentials.json`,
`service_account.json` — all must stay in `.gitignore`. The `.pub` file is
harmless to commit if you want, but there's no need to.

## Running it locally

1. **Install Python 3.9+** and open this folder in VS Code.
2. **Install dependencies:**
   ```bash
   pip install openpyxl google-api-python-client google-auth snowflake-connector-python
   ```
3. **Set up `snowflake_credentials.json`** as shown above, and make sure your
   Google service account JSON is in place for the Sheets API calls.
4. **Run the build:**
   ```bash
   python build_data.py
   ```
   This connects to Snowflake, pulls the Google Sheets tabs, and rewrites
   every file in `data/`.
5. **View the dashboard** — don't just double-click `index.html`; browsers
   block local `fetch()` calls that way. Instead:
   ```bash
   python -m http.server 8000
   ```
   then open **http://localhost:8000**. (Or use the VS Code Live Server
   extension.)

Whenever you re-run `build_data.py`, just refresh the browser tab.

## Two contributors, one repo

A colleague independently pushes to this same repo/branch, owning the Store
Cut Size, Sales vs Inwards, and Store Returns tabs. Merge conflicts are
currently resolved manually: fetch `main`, diff function lists between both
versions of `index.html`/`build_data.py`, reconcile, re-commit. An in-progress
restructuring plan (splitting `index.html` into per-tab JS modules) is partly
aimed at reducing how often this happens — see Known Issues / open work below.

## Known issues / open work

- **`sales_by_sku.json` file size**: plan is to split it into half-year files
  (`sales_by_sku_2025H2.json`, `sales_by_sku_2026H1.json`, etc.) rather than
  shrinking `ROLLING_WINDOW_DAYS`, since prior-year data needs to stay
  available for ROS/live-start calculations. Not yet implemented.
- **`index.html` monolith**: currently ~4,700 lines covering all 15 tabs in
  one file. Planned restructure: a small shell `index.html` (header/filter
  bar/nav only) + `filters.js` (shared global filter state) + one `.js`
  module per tab, loaded via dynamic `import()` on first click. Hash-based
  routing (`#pareto`, `#ageing`) for deep links, since GitHub Pages has no
  server-side router. Not yet implemented.
- **Inventory occasionally comes back empty** (`inventory.json` writes as
  0.0 MB, `Inv Data 2` fetches only its header row). The script itself
  detects and warns about this — it's the known symptom of an unauthenticated
  `IMPORTRANGE` failing to resolve in the underlying Google Sheet. Fix: open
  `Automation_Data.xlsx`'s source sheet manually in a browser once to
  re-authorize the `IMPORTRANGE`, then re-run.
- **Ageing tab Channel filter**: `daily_fifo_ageing` (the source table) has no
  channel column, so Channel filtering only works for the Sales Mix%/GM% by
  Month view and the SKU-Level Ageing List's Current Inv figure (via a join
  to `store_inv_by_sku`) — not for MoM bucket tables, Inv Mix, or the STR%
  inventory side. This is a data limitation, not a bug to fix in code.
- **Live backend was evaluated and rejected**: an always-on API serving
  Snowflake data live was considered, to avoid committing large JSON files.
  Rejected — always-on server cost, multi-second query times against
  7.5M-row tables, and a larger credential exposure surface, for no real
  reduction in what the browser ultimately has to download. Static
  daily-refresh JSON stays the approach unless a genuine real-time
  requirement appears.

## Key business logic (do not casually change)

- **DOI has two distinct definitions**: target-based vs. L30D-actual-rate —
  intentionally different, don't conflate them.
- **ROS denominator** uses each SKU's live-start date, not raw date-range length.
- **"Latest Inward"** only counts restock events of ≥50 units as real.
- **Current Inv** = Warehouse (Online) + Store (Offline), from
  `inv_for_auto_detail` (no trailing "s" in the table name — has caused a
  real bug before from mistyping it).
- **Minimum/ideal GM% threshold**: 55%.
- **Reorder/Repeat/Fresh SKU classification**: Reorder = 2+ qualifying
  ≥50-unit inwards in the same season, ≥60 days apart. Repeat = qualifying
  inwards spanning 2+ different seasons. Fresh = everything else. Reused via
  `buildRepeatClassificationMap()`/`identifyFreshSkus()` in both the Repeat
  Evaluation tab and as a filter in Pareto/Quartile, so the definition can
  never diverge between them.
- **Global SKU_Group exclusion filter** strips null/`CD*`/`NT*`-prefixed SKU
  groups from all SKU-grain datasets, applied in both Python and JS as
  belt-and-suspenders.
- **Target Attainment** is restricted to the Aug–Dec window only.
- **Fresh SKU Reorder tab was removed** per explicit request — don't re-add
  without asking.