# Geopolitics Article Scraper

Scrapes think-tank websites for the newest articles under a chosen geopolitical topic
(Iran, Ukraine, …), parses each into structured rows, and appends them to an Excel
repository workbook. Runs daily via Windows Task Scheduler.

This is the web-scraping counterpart to the newsletter **Repository Scraper**
(the Outlook email tool). It shares that tool's design — a two-stage *fetch → append*
pipeline, duplicate suppression, OneDrive-lock handling, the same 26-column workbook
schema, and a canonical/read-only-copy split — but its source is the web instead of Outlook.

The topic is **configurable**: change `active_topic` in `topics.json` (or pass `--topic`)
to switch briefs without touching code.

---

## Sources (topic = Iran)

| Think tank | Iran source | How it's read |
|---|---|---|
| Atlantic Council | `/region/iran` feed | RSS (`feedparser`) |
| CSIS | `/regions/middle-east/iran` | HTML scrape (`article` cards) |
| CFR | `/middle-east-and-north-africa/iran` | HTML scrape (`h3 > a[/articles/]` + `<time>`) |
| Chatham House | `/regions/middle-east-and-north-africa/iran` | HTML scrape (`article.teaser`) |

---

## Files

| File | Role |
|---|---|
| `run_daily_article_scraper_update.ps1` | Orchestrator: fetch → append → refresh review copy. **Point Task Scheduler at this.** |
| `fetch_thinktank_articles.py` | Reads `topics.json`, scrapes each source for the active topic, writes a per-run JSON. |
| `append_article_repository_workbook.py` | Reads that JSON and appends new rows to the workbook (duplicate suppression). |
| `sources.py` | The scraper/parser (per-site extractors, RSS, date parsing). **Imported by the fetcher.** |
| `topics.json` | Topic profiles: each = a name + four source URLs + optional keywords. |
| `requirements.txt` | Python dependencies. |

**Data flow**
```
topics.json → fetch_thinktank_articles.py → <per-run JSON>
            → append_article_repository_workbook.py
            → canonical workbook → (copied to) read-only review copy
```

---

## Prerequisites

- **Windows** with **Python 3.11+**.
  ```
  pip install -r requirements.txt
  ```
- **PowerShell 5.1+** (built into Windows).
- **Outbound internet** to the four sites (works through a corporate proxy if
  `HTTP_PROXY`/`HTTPS_PROXY` are set — `requests` honors them).

No API keys or secrets are needed.

---

## Configure

Open `run_daily_article_scraper_update.ps1` and set the `param(...)` defaults:

- `$WorkbookPath` — the **canonical** workbook the automation writes. Put it somewhere
  **only the automation touches** (see "Sole-writer" below).
- `$ReviewCopyPath` — a read-only copy for people to review (refreshed after each run).
- `$OutputDirectory` — folder for per-run JSON exports and `scheduled_task.log`.
- `$Topic` — leave empty to use `active_topic` from `topics.json`, or set e.g. `"Ukraine"`.
- `$Python` — leave empty to auto-discover, or set the full path to `python.exe`.

### Adding a new topic (e.g. Ukraine)

Edit `topics.json`. A topic profile is just a name, four source URLs, and optional keywords:

```jsonc
"Ukraine": {
  "require_keywords": false,
  "keywords": ["ukraine", "kyiv", "zelensky"],
  "sources": [
    { "name": "Atlantic Council", "parser": "rss",     "url": "<UkraineAlert feed URL>" },
    { "name": "CSIS",             "parser": "csis",    "url": "<CSIS Ukraine page>" },
    { "name": "CFR",              "parser": "cfr",     "url": "<CFR Ukraine page>" },
    { "name": "Chatham House",    "parser": "chatham", "url": "<Chatham Ukraine page>" }
  ]
}
```
Then set `"active_topic": "Ukraine"` (or run with `--topic Ukraine`). A starter Ukraine
profile is included but its URLs are **unverified guesses** — confirm each one loads the
right topic page before relying on it. `parser` is one of `rss`, `cfr`, `csis`, `chatham`,
or `generic` (a best-effort fallback for a brand-new site).

Per-topic options:

- `lookback_days` (default in config: 30) — only keep articles published within the last
  N days, so a run surfaces recent items instead of years of backlog. Override at runtime
  with `--lookback-days N`, or disable the window entirely with `--all` (useful for a first
  full backfill). Articles with no parseable date are always kept (never silently dropped).
- `require_keywords: false` (default) trusts each site's own topic curation. Set it `true`
  to additionally drop any article whose title/summary/URL contains none of the `keywords`
  — useful if a site's topic page mixes in off-topic items.

Per-source options:

- `fetch_summary: true` — for a source whose listing page has no abstract (e.g. CFR), visit
  each kept article and pull its meta/`og:description` (or first paragraph) into the summary.
  Adds one light request per new article; leave it off for sources that already list summaries.

---

## Run once (test)

```powershell
& ".\run_daily_article_scraper_update.ps1"
```
Or the stages directly:
```powershell
python fetch_thinktank_articles.py --output-json test.json
python append_article_repository_workbook.py --input-json test.json --workbook test_repo.xlsx
```
Re-running is safe: duplicates are skipped (identity = article URL).

---

## Schedule (Task Scheduler)

```powershell
$action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-NonInteractive -ExecutionPolicy Bypass -File "C:\path\run_daily_article_scraper_update.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At "9:15AM"
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 20)
Register-ScheduledTask -TaskName "DailyIranArticleScrape" -Action $action -Trigger $trigger -Settings $settings
```
`-StartWhenAvailable` catches up a run missed because the machine was off. No Outlook
needed (unlike the newsletter tool), so it does not depend on a logged-in mail session.

---

## Known limitations & gotchas

- **Parsers are HTML-structure-dependent.** CSIS/CFR/Chatham are scraped from page markup;
  if a site redesigns, its extractor in `sources.py` may need updating (each site is one small
  function). Atlantic Council uses RSS and is the most stable. The fetcher logs a per-source
  count each run — a source suddenly returning 0 is the signal that its layout changed.
- **Dates.** CFR/Chatham expose machine-readable `<time datetime>`; CSIS renders the date as
  text and is parsed heuristically. An article with no parseable date still lands in the
  workbook with a blank Publication Date.
- **Sole-writer workbook.** Keep the canonical workbook where nobody opens it; let people use
  the read-only review copy. Every run's raw articles are saved as JSON, so the workbook is
  always rebuildable by re-running the appender over them (it de-duplicates on URL).
- **OneDrive sync latency / locks.** The appender retries on locks (`retry_on_lock`).
- **Politeness.** One request per source with a 1s pause between sources and a real
  User-Agent. This is light, read-only traffic on public pages.

---

## Troubleshooting

Open `scheduled_task.log` in your `$OutputDirectory`:
- A source shows `WARNING: source '<name>' failed` → that site was unreachable or changed;
  the run still completes for the others.
- A source returns 0 articles when it shouldn't → its page layout likely changed; update that
  extractor in `sources.py`.
- `PermissionError` on save → the workbook was open/locked (close Excel; the retry usually clears it).
- No Python found → pass `-Python "C:\path\to\python.exe"` to the orchestrator.
