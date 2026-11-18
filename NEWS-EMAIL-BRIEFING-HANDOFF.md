# Daily News Email Briefing — Handoff Notes (2026-11-18)

Plain-English recap for picking this up in a new conversation or handing it to another team.
Code lives in [`news-email-briefing/`](news-email-briefing/) in this repo — see its
[`README.md`](news-email-briefing/README.md) for setup. This file covers the operational context,
history, and current status that isn't obvious from the code alone.

## ⚠️ Current known-broken state (as of 2026-11-18)

**The tool is NOT currently sending briefings.** OpenAI API calls are failing:

```
AuthenticationError: Error code: 401 - {'error': {'message': 'You do not have access to the
project tied to the API key.', 'type': 'invalid_request_error', 'code': 'invalid_project', ...}}
```

Background: this followed weeks of a *different* error (`429` / organization spend-limit
exceeded on the Deloitte-managed OpenAI org, first seen ~2026-11-03) that was escalated to
Deloitte org admins via the user's director with no resolution. As of today the error changed to
the `invalid_project` message above — which reads as the key's access being revoked/removed
outright, not just rate-limited, so this is a **more severe lockdown than before, not a fix**. The
user is separately pursuing a personal OpenAI API key as a workaround; if/when that lands, just
set `OPENAI_API_KEY` (and drop `OPENAI_BASE_URL`'s `us.api.openai.com` override, which is a
Deloitte-key-specific workaround — see below) to the new key's values.

**Before relying on anything below "works," confirm the API key issue is actually resolved** by
running `python daily_news_summary.py` manually and checking `daily_email.log` for a clean run.

## What this tool does

Reads today's (plus genuine yesterday-leftover) rows from the shared Expert Commentary Repository
workbook, has an LLM (OpenAI `gpt-5.4` by default) write a senior-executive-level daily briefing
grouped into "Executive Summary" + "Key Themes," runs it through two independent quality gates,
and opens it as a fully-addressed Outlook draft — **a human must click Send every single day**;
nothing is ever auto-sent to the distribution list, regardless of how the quality checks go.

This is a *consumer* of the [`expert-commentary-scraper`](expert-commentary-scraper/) tool's
output workbook, not a sibling scraper itself — it has no scraping logic of its own. That means
it's fragile to that tool's shared-folder path changing (see below), even though it's not the tool
being handed to the research department.

## Where everything lives

**Code:** `news-email-briefing/` in this repo (`daily_news_summary.py`, `run_briefing.ps1`,
`requirements.txt`, `README.md`).

**Live copy actually run by the Scheduled Task** (not synced to this repo automatically — copy
changes back manually, or point the task at the repo copy):
```
C:\Users\mivogel\OneDrive - Deloitte (O365D)\Desktop\Tom's News Email Automation Practice\
    email-briefing\
        daily_news_summary.py
        run_briefing.ps1
        daily_email.log              (run log, not in repo)
        used_articles_log.json       (leftover-tracking ledger, not in repo)
```

**Source workbook it reads from** — **as of today, points at the New Process & Scope folder**
(previously pointed at Old Process & Scope; the scraper tool's source set changed on 2026-10-21,
see [`expert-commentary-scraper`](expert-commentary-scraper/) memory/handoff for why):
```
C:\Users\mivogel\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1\
    New Process & Scope\
        Expert Commentary Repository (review).xlsx
```
Sheet `"Repository Rows"`, header row 1. **Important schema gotcha**: the New Process workbook
names its date column `"Publication date"` (lowercase "date"); the Old Process workbook used
`"Publication Date"` (capital), or on some older snapshots a split
Day/Month/Year set of columns. `load_todays_news()` in `daily_news_summary.py` matches the date
column case-insensitively and normalizes it, specifically so this doesn't silently break again if
the scraper's schema drifts further — but a genuinely new column name (not just a case change)
would still need a code fix.

**Windows Scheduled Task:** `"Daily News Email Briefing"`. Runs `run_briefing.ps1` (sets
`OPENAI_BASE_URL` for the Deloitte-network DNS workaround, then calls the Desktop copy of
`daily_news_summary.py`). Registered Mon–Fri only (`DaysOfWeek` bitmask `62`), `AllowStartIfOnBatteries`
on. **Trigger currently shows 11:00 AM local (Eastern) as the real recurring fire time** (checked
directly via `Get-ScheduledTask`/`NextRunTime`, not just the task's name) — flag this to the user
if a 12:00 PM send was expected; the registered trigger and the intended time may have drifted
apart at some point.

## Recipients (as of 2026-11-18)

- **`YOUR_EMAIL`** (drafts open here for review/send): `mivogel@deloitte.com`
- **CC distribution** (only reached once the draft is manually sent):
  `stewatkins@deloitte.com`, `gacorbridge@deloitte.com`, `jstiller@deloitte.de`,
  `victyu@deloitte.com`, `tsellersclyde@deloitte.es`, `jeschiu@deloitte.com`

## Architecture notes worth knowing

- **Manual-send-only by design**: `open_email_for_review()` uses Outlook's `Display()`, never
  `Send()`, for the real distribution-list email — this was an explicit, firm user requirement
  ("every day requires your click"), not a bug to "fix" by automating it further.
- **Two-layer quality gate before the full list is ever addressed**: a fast deterministic check
  (`check_structural_quality` — required sections present, not too short, no leaked sentinel
  text) followed by a second, independent LLM review pass (`review_summary_quality`). Either
  failure routes a private notice (raw draft included) to `YOUR_EMAIL` only — added after a real
  incident where a content-free day produced a near-empty, broken-looking email to the whole list.
- **1-day leftover carry-forward**: `used_articles_log.json` records exactly which
  `(Source, Article Title)` pairs were drafted each day; the next day's run folds in any
  yesterday-dated rows NOT in that set (arrived too late for yesterday's draft), with an explicit
  overlap-check log line every run. Deliberately capped at 1 day — not an indefinitely
  accumulating backlog.
- **Quote attribution uses the workbook's `Source` column** (the organization, e.g. "Eurasia
  Group", "Middle East Council on Global Affairs"), never `Newsletter Name` (too generic/
  ambiguous — multiple orgs publish something literally called "Daily Brief") and never an
  individual advisor's name. This was tightened twice after real leaks: once from an earlier
  Newsletter-Name-based attribution, and once from a literal `"Source: XX"` string leaking into
  generated text due to ambiguous prompt wording.
- **Resilient read on a 1-15 min retry loop** (`load_todays_news_resilient`, 16 attempts/60s
  apart): the shared workbook is often edited on another device, and OneDrive can take several
  minutes to sync the update down to this machine while still showing the original edit
  timestamp, so "no rows found yet" isn't necessarily "no news today."
- **Run-level retry** (`run_daily_job_with_retry`, 3 attempts/120s apart) around the whole job for
  transient OpenAI/network/Outlook failures.

## Known real bugs found and fixed

1. **Date-schema drift, multiple times** — the source workbook's date columns changed shape more
   than once (combined `"Publication Date"` ↔ split Day/Month/Year ↔ New Process's lowercase
   `"Publication date"`). A naive string-concatenation date rebuild from split columns silently
   produced all-NaT dates when Excel/pandas stored day/year as floats (`"24.0"` doesn't match a
   `"%b %d %Y"` format) — fixed by building the date numerically
   (`pd.to_datetime(dict(year=, month=, day=))`) instead of via string formatting.
2. **Quote-attribution leaks** — fixed twice: switched from citing `Newsletter Name` to `Source`
   (Newsletter Name was too ambiguous, e.g. multiple orgs' newsletters both called "Daily Brief"),
   and separately fixed a literal `"Source: XX"` string leaking into LLM output due to prompt
   wording that didn't make clear the label itself shouldn't be echoed.
3. **Missing "KEY THEMES" heading** on some low-item-count days — the prompt wasn't explicit that
   the heading must always appear even with very few themes; tightened to require it verbatim,
   exactly once, even for 1-2 item days.
4. **A content-free day sent a broken-looking email to the full distribution list** — root cause
   was no pre-send content/quality check at all. Led directly to the two-layer quality gate
   described above, plus the `MIN_ITEMS_FOR_BRIEFING` item-count floor.
5. **Scheduled task silently not firing** — root-caused to `DisallowStartIfOnBatteries=True` while
   the laptop was unplugged; fixed via `AllowStartIfOnBatteries`/`DontStopIfGoingOnBatteries`.
6. **Whole automation folder was deleted once** (cause unknown) — fully recreated from
   conversation history; `used_articles_log.json` had to be manually reseeded.
7. **Workbook path moved multiple times** as the scraper team reorganized the shared folder,
   most recently the Old→New Process & Scope split on 2026-10-21 (fixed today, 2026-11-18, along
   with the case-sensitive date-column bug that fix would otherwise have triggered — see above).

## How to proceed

Start a new conversation with this repo as the working directory — point a new session at this
file for full context; there's also a persistent Claude memory file for the user's own sessions
that duplicates this, but this file is the durable, shareable version.

## Small things still open

- **The OpenAI API key is currently broken** (see top of this doc) — nothing else matters until
  this is resolved one way or another.
- **Scheduled task trigger time should be double-checked against what the user actually wants** —
  observed real trigger is 11:00 AM ET Mon-Fri; there's reason to believe 12:00 PM ET was the last
  intended time. Verify with the user before treating either as correct.
- **No automatic fallback search for `EXCEL_PATH`** — if the scraper team moves the shared folder
  again, this will break again in exactly the same way. Worth adding a fallback path search (e.g.
  checking both Old/New Process subfolders, or a small config file shared between the two tools)
  if these path moves keep recurring — not yet requested or built.
- **Personal OpenAI API key as a workaround**: the user was in the process of creating one outside
  the Deloitte-managed org as of 2026-11-18; a real data-governance tradeoff (company content
  routed through a personal, non-Deloitte-sanctioned account) was flagged to the user but the
  final call is theirs to make — if you see a personal key in use, that's why.
