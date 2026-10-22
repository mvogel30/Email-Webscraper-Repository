---
name: Expert Commentary Scraper
description: Run or check the expert-commentary scraper - a tool that pulls expert-commentary articles from 8 think-tank/news hub pages and vendor emails (Signum, Eurasia Group, Middle East Council on Global Affairs, Vriens & Partners) into a shared Excel workbook. Use this whenever the user asks to run/update/refresh the scraper or "the commentary tracker/workbook", check whether today's update ran, check whether the workbook copies are synced, or asks what's in the repository. This is a real tool with real side effects (writes to a shared team Excel file, hits live external websites, drives a real running Outlook) - only invoke it when the user is actually asking to run or check it, not just discussing the topic areas it covers.
shell: powershell
---

# Expert Commentary Scraper

This skill wraps an existing, already-working tool - it does not reimplement anything. Every
action below just runs one of the scripts already in this folder (`fetch_articles.py`,
`append_article_repository_workbook.py`, `run_daily_full_update.ps1`, the `email-scraper/`
pipeline, `check_sync.py`). If you ever find yourself writing NEW scraping/parsing logic to answer
a request this skill covers, stop - read the existing script first, since it almost certainly
already does what's needed.

This tool ALSO runs completely independently as a Windows Scheduled Task on its original author's
laptop (once a day). This skill is a second, conversational way to invoke the SAME scripts and
read the SAME shared workbook - running it here never conflicts with or duplicates the scheduled
task; both just append-and-dedupe into the same file.

**Where the tool's files are**: this skill's own folder is `${CLAUDE_SKILL_DIR}`. The tool's
scripts live three levels up, at `${CLAUDE_SKILL_DIR}/../../..` (i.e., `.claude/skills/run-update/`
→ `.claude/skills/` → `.claude/` → the `expert-commentary-scraper` folder itself). Resolve that
path once at the start and use it as the working directory for everything below.

## Before running anything: check prerequisites

1. **Python 3** must be installed and on PATH (or discoverable the way the existing scripts look
   for it - see `Resolve-Python` in `run_articles_update.ps1` if `python`/`py` aren't found
   directly).
2. **Python packages**: `pip install -r requirements.txt` (one-time, or after a `git pull`/zip
   update changes it).
3. **Playwright's headless browser** (only for the WEF source, which blocks plain HTTP clients):
   one-time `python -m playwright install chromium` (~300MB download). If this hasn't been done,
   the web scrape still runs fine for every other source - WEF alone will error out non-fatally.
4. **Outlook desktop app** must be open and logged in for the vendor-email half of the pipeline
   (Signum, Eurasia Group, etc.) - it works via COM automation of an already-logged-in Outlook, not
   an API key. If Outlook isn't running, that half fails independently; the web-scraping half is
   unaffected (see `run_daily_full_update.ps1`'s own try/catch - one half failing doesn't sink the
   other).

If any of these are missing, tell the user what's missing and offer to install/set it up rather
than silently failing partway through a run.

## Running a full update

Run this from the tool's folder (see path above):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "run_daily_full_update.ps1"
```

This takes several minutes (the web-scraping half alone typically takes ~5-7 minutes fetching full
article text from every source). Let it run to completion rather than assuming it hung.

Read the output carefully and summarize it for the user rather than just saying "done":
- Per-source counts for the web half (`Atlantic Council`, `CSIS`, `CFR`, `Chatham House`, `Lowy
  Institute`, `China Macro Group`, `Deloitte Insights`, `WEF`) - a source showing as an `error`
  object (not a plain number) genuinely failed; note which one and the error type.
- `rows_added` / `rows_skipped_duplicate` for both the web and vendor-email halves - a run showing
  0 rows added is not necessarily broken, it just means everything found was already in the
  workbook (normal on a re-run, or if nothing genuinely new was published/sent since the last run).
- Whether the review copy / team copy / personal mirror refreshed successfully at the end, or hit
  a file-lock warning (see `references/troubleshooting.md` - this is a known, common, easily-fixed
  situation, not a real bug).

You can also run just one half if that's specifically what's being asked:
- Web sources only: `run_articles_update.ps1`
- Vendor emails only: `email-scraper\run_vendor_email_update.ps1`

## Checking status ("did it run?", "is it up to date?", "is everything synced?")

**Did today's scheduled run happen, and did it succeed?**

```powershell
Get-ScheduledTaskInfo -TaskName "Expert Commentary Scraper - Daily 0600" | Format-List LastRunTime, LastTaskResult, NextRunTime, NumberOfMissedRuns
```

`LastTaskResult` of `0` means success. This only tells you the task fired and exited cleanly - it
does NOT tell you whether anything NEW was actually added (a run can succeed with 0 new rows, see
above). For that, check the workbook's own Run Log sheet (a sheet inside the workbook itself,
separate from `Repository Rows`) - its last few rows show the timestamp, rows found, rows added,
and rows skipped as duplicates for every run to date, scheduled or manual.

**Is the workbook actually the same across all its copies?**

```powershell
python check_sync.py
```

This compares the canonical workbook against its review copy, team copy, and (if present) a
personal mirror, row by row - not just file size or timestamp, since those can differ even when
the data is identical (e.g. simply opening a workbook in Excel and re-saving it changes the
underlying file's bytes with zero visible data change - a real false alarm to watch out for, don't
report files as "out of sync" from a size/hash difference alone without checking actual content).

## Where the shared workbook lives

Everything below is under one shared folder (a OneDrive-synced Teams/SharePoint library):
`<OneDrive root>\Global PP+G - Webscraping & Email Scraping tool 1\`

(2026-08-18: the SharePoint site was restructured - a real move, not just a rename, of what used
to be under `Global PP+G - 06 Sensing (Inputs)\10 - Geopolitics Sensing & Scraping Tools\
Webscraping & Email Scraping tool\`. The local OneDrive shortcut is now added directly at the
`Webscraping & Email Scraping tool` folder level, so there's no extra nesting under the shortcut
root anymore. Every script's `-SharedLibraryRoot`/`--shared-library-root` default was updated to
match, and `-SharedSubfolder` now defaults to empty.)

(2026-08-23: moved again, to a location under the same site's `Shared Documents\03 Cross-Team
Projects\08 - Geopolitics Sensing & Scraping Tools\Webscraping & Email Scraping tool`. On this
machine the freshly-added shortcut collided by name with the still-intact old-location folder, so
OneDrive suffixed it " 1" - hence the `... tool 1` above. That old-location folder was
deliberately left untouched as a revert point: to roll back, drop the " 1" suffix in every
script's `-SharedLibraryRoot` default and in the Scheduled Task's stored action argument.)

(2026-10-21, per user request: everything below now lives one level deeper, under
`Old Process & Scope\` - a sibling `New Process & Scope\` folder also exists, reserved for
whatever eventually replaces this pipeline, currently empty. Every script's
`-SharedSubfolder`/`--shared-subfolder` default was updated from empty to `Old Process & Scope`
to match, including the secondary Teams-mirror copy path, which also needed a long-path-safe copy
fix - see the comment above the `Copy-ItemLongPathSafe` function in `run_articles_update.ps1` -
since that mirror's own base path is long enough that this extra nesting pushed the canonical
file's full path past Windows' 260-char limit.)

- `DO NOT OPEN - Reserved for Automated Updates - Expert Commentary Repository.xlsx` - the
  canonical file. Scripts write here directly. **Never open this one in Excel** - if it's open,
  the next scheduled/manual run can't save to it.
- `Expert Commentary Repository (review).xlsx` - safe to open, refreshed from canonical after
  every run.
- `Expert Commentary Repository (team copy).xlsx` - same content as review, kept as a second copy
  because the review file's Teams/SharePoint sync has been unreliable when it's open in Excel;
  this one is never opened by anyone so it stays syncable.
- A personal mirror may also exist per-user under their own `Documents\Expert Commentary
  Repository\` - convenience only, never the source of truth.
- `Setup\expert-commentary-scraper.zip` - the distributable team-install package. **Rebuilt
  automatically every day** by `team-install\Build-DistributableZip.ps1`, run as the last step of
  `run_daily_full_update.ps1` (non-fatal - a packaging hiccup never blocks the real scraping/email
  work). This means any code change made in ANY session - including one you're doing right now -
  is picked up in the shared zip within a day with no manual rebuild step needed; you do NOT need
  to remember to rebuild/redistribute it after editing this project's code.

## If something looks wrong

Read `references/troubleshooting.md` first - it covers real, previously-diagnosed issues (a
persistently-blocked source, a source that looks broken but isn't, a common file-lock pattern, and
what "the workbook schema" actually means) before assuming something is newly broken.
