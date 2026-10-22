# Expert Commentary Scraper — Handoff Notes (2026-08-05)

Plain-English recap for picking this up in a new conversation. Claude's persistent memory
(`expert-commentary-scraper.md`) should surface automatically in a new thread too — this file is
a durable, in-repo backup of the same context.

## What this tool does

Sibling to the Iran-specific `geopolitics-article-scraper/` tool, but **unfiltered** — no topic/
keyword filter. Pulls every article from 8 expert-commentary hub pages, plus every relevant email
from 4 vendors, into one shared Excel workbook.

## Where everything lives

**Code:** `expert-commentary-scraper/` in this repo (repo root, sibling to
`geopolitics-article-scraper/`), with `email-scraper/` for the vendor-email pipeline.

**Shared workbook (open this one):**
```
C:\Users\mivogel\OneDrive - Deloitte (O365D)\Global PP+G - 06 Sensing (Inputs)\
    10 - Geopolitics Sensing & Scraping Tools\
        Webscraping & Email Scraping tool\
            Expert Commentary Repository (review).xlsx
```
(Don't open the canonical file — as of 2026-08-06 it's named
`DO NOT OPEN - Reserved for Automated Updates - Expert Commentary Repository.xlsx` in the same
folder (renamed from `(auto-managed).xlsx` at the user's request, warning leading the filename so
it's the first thing seen rather than trailing in parens). Scripts write there directly; the
review copy is the safe one to open.)

**If Teams shows a stale version of the review file** (2026-08-05: this happened — Explorer, which
reads the local OneDrive cache directly, showed the review file as current, but the user and
teammates viewing it through Teams did not), there's a second copy in the SAME folder:
`Expert Commentary Repository (team copy).xlsx`. It's refreshed on every run exactly like the
review file, but nobody ever opens it in Excel, so it's never blocked from syncing up to Teams the
way the review file can be when someone has it open locally. Both scripts (`run_articles_update.ps1`,
`email-scraper/run_vendor_email_update.ps1`) write it via a new `$TeamCopyFileName` param, threaded
through `run_daily_full_update.ps1` too.

**Run manually:** double-click `Run Expert Commentary Scraper.cmd` in the project folder.

**Automated daily run:** Windows Scheduled Task **"Expert Commentary Scraper - Daily 0600"** on
this machine. Note: this machine's clock is **Pacific Time**, and the user wanted "9am EST" — 6am
Pacific = 9am Eastern, confirmed by checking the sibling Iran tool's own task the same way. If
touching this task again, remember `RunAtTime` is always LOCAL wall-clock time, no timezone
conversion built in.

**Team install package:** `expert-commentary-scraper/team-install/` (generic, prompts for a
time — no US/India-pinned variants like the Iran tool). Distributable zip lives at
`...\Webscraping & Email Scraping tool\Setup\expert-commentary-scraper.zip` — **kept in sync all
session**: rebuilt and redistributed after every single code change today (verified current as of
this handoff — zip timestamp is after every source file's last edit).

## Web sources — all 8 working

| Source | Notes |
|---|---|
| Atlantic Council (Dispatches) | New extractor, real HTML, not RSS |
| CSIS (Analysis) | Reused existing extractor unmodified |
| CFR (Expert Takes) | Reused unmodified |
| Chatham House (Expert Comment) | Reused unmodified |
| Lowy Institute (The Interpreter) | RSS feed |
| China Macro Group (Insights) | Reused unmodified |
| Deloitte Insights | **Solved via their public sitemap**, not the hub page (which is dateless/JS-rendered) — see memory for the JSON-LD date discovery |
| WEF (Publications) | **Requires Playwright** (headless Chromium) — WEF blocks plain HTTP clients but not real browsers. New dependency: `playwright>=1.62` in `requirements.txt`, PLUS a one-time `python -m playwright install chromium` (~300MB) NOT covered by pip alone. WEF's "Full Text" is the listing description, not the linked PDF (out of scope). |

## Email vendors — 4 configured, domain-based matching

`email-scraper/vendors.json` lists `{name, domain, parser_module}`. Matches ANY sender at the
configured domain — not specific people, not subject tags.

- **Signum** (`signumglobal.com`) — own parser. Newsletter Name: "Week Ahead", "Signum" (default
  country/topic notes). **Event invitations are EXCLUDED from scraping entirely** — user request,
  these aren't research content. Named exceptions: "SIGNUM Zoom", "SIGNUM Investor Trip", "SIGNUM
  Policy Day" (matched generically, e.g. "SIGNUM Paris Policy Day" — location varies). Plus a
  general catch-all **"SIGNUM Event"** (2026-08-07, added after "SIGNUM New York: Hybrid
  Roundtable...") for the shared shape `SIGNUM <word(s)>:` (a location/descriptor between "SIGNUM"
  and a colon) — real reports always use "SIGNUM:" with the colon immediately after, so this
  doesn't false-trigger on them; verified against every real captured Signum subject to date.
- **Eurasia Group** (`eurasiagroup.net`) — own parser. Newsletter Name: "EG Update" and "EG Daily
  Brief" are BOTH rich multi-section digests (same "the top line"/"eg in depth"/"LABEL—Title"
  heading grammar) and both get split into one row per topic/section — "EG Daily Brief" was
  originally misjudged as single-topic from only seeing a trailing snippet of 2 real samples;
  reading the full bodies on 2026-08-05 showed 6-8+ sections per email. Also: "Daily Research
  Summary", generic "Eurasia Group" (internal admin threads) are kept whole. **"EG Conference
  Call" (invitations) are EXCLUDED entirely** — same reasoning as Signum's events.
- **Middle East Council on Global Affairs** (`mecouncil.org`) — domain confirmed via a REAL
  received email 2026-08-05 ("Daily Brief | August 5"). Uses the shared `generic_vendor_parser.py`.
- **Vriens & Partners** (`vrienspartners.com`) — domain verified via web search (real staff
  emails found), but **no real email has arrived yet** — watch for the first one.

## Missing older emails, fixed 2026-08-05 (lookback-window gap, not a folder problem)

User reported missing 7/29-7/31 "EG Daily Brief" rows and many Signum emails, and suspected it was
because they'd manually moved a lot of vendor emails into a custom Outlook folder ("External
Advisor Briefs"). Diagnosed with a read-only Outlook COM scan: that folder IS reached and scanned
correctly (confirmed 318 vendor-domain matches inside it) — **the real cause was the lookback
window**, not the folder. `run_vendor_email_update.ps1` defaulted to `-DaysBack 3`, so any run only
ever looks back 3 days from "now" — a vendor email older than that, at the time of every run since
the tool started, simply never gets scanned, and no future run will ever go back and pick it up
either (each one only looks forward from its own "now"). Confirmed via the same scan: the mailbox
actually has Signum back to 2026-05-15 and Eurasia Group back to 2026-05-13.

Fixed two ways (both per user confirmation):
1. **One-time backfill**: exported vendor emails from 2026-07-22 through today (covering the
   reported gap with margin) and appended them — 85 raw emails, 85 new rows added (existing dedup
   correctly skipped 28 already-scraped ones), workbook went from 158 → 243 rows. Verified 7/29-
   7/31 specifically: all three days' "EG Daily Brief" emails now present and correctly split (8
   sections each on 7/30 and 7/31), plus the previously-missing Signum emails.
2. **Widened the daily default from 3 → 7 days** in `run_vendor_email_update.ps1` (nothing hardcodes
   `-DaysBack` anywhere else — not `run_daily_full_update.ps1`, not the registered scheduled task —
   so this one default change takes effect immediately for future runs). Costs nothing extra since
   `append_rows()` already dedupes by URL/title+source+date — a wider window just adds resilience
   against a missed day (closed laptop, Outlook not open, etc.) without ever double-adding content.

If a similar "missing content" report comes up again: check the actual date range of matching
emails across the WHOLE mailbox first (a wide, no-date-filter Outlook COM scan, same technique as
this diagnosis), don't assume it's a folder-visibility problem — the folder walk in
`export_vendor_emails.ps1` already recurses through every folder in every account/store, skipping
only "Sync Issues".

## Workbook schema (as of 2026-08-05, changed from the original build)

Leftmost columns, in order: **Report** ("Both" / "Spotlight only" — which of the user's weekly
reports a source feeds) → **Scraping-Type** ("Web"/"Email") → **Source** (organization) →
**Newsletter Name** (the specific named publication/series — e.g. "Dispatches", "EG Update").

This is DIFFERENT from the original build: "Source"/"Newsletter Name" were swapped and a genuinely
new "Newsletter Name" column was added, per the user's boss's feedback. If you see old references
to "Source" meaning Web/Email, that's the STALE pre-2026-08-05 meaning.

## Known real bugs found and fixed today (worth knowing about if similar symptoms recur)

1. **openpyxl silently truncates cells over 32,767 characters** — no exception, no warning, cuts
   mid-word. Now handled: `_fit_to_cell_limit()` in `append_article_repository_workbook.py`
   truncates deliberately with a visible marker. If you ever see article text end abruptly with
   no punctuation, check length — if it's ~32767, this is why (and check the fix is still there).
2. **Windows Task Scheduler rejects a literal colon in `-TaskName`** — fixed in BOTH this tool's
   AND the sibling Iran tool's `Install-Teammate.ps1` (strip the colon before building the name).
3. **Eurasia's multi-topic digest splitter never actually triggered on real data** — it checked
   for "daily brief" in the subject, but the real trigger is "eg update". Fixed; verified against
   every real captured Eurasia email.
4. Several boilerplate-leak fixes (CSIS copyright/disclaimer paragraphs, Eurasia "Read more here"
   CTAs, Middle East Council's subscribe header/address footer) — see memory for details.
5. **Reply/forward emails now excluded entirely** — any vendor email whose subject starts with
   "RE:"/"FW:"/"Fwd:" (personal correspondence, not a newsletter broadcast) is filtered out in
   `parse_and_append_vendor_emails.is_reply_or_forward()`. General across all vendors, not a
   per-vendor special case.
6. **"EG Daily Brief" now splits into per-topic rows, like "EG Update" already did** — it was
   wrongly treated as single-topic (kept as one whole-email row) based on only seeing a trailing
   snippet of the real data; the full bodies show the same rich multi-section structure as "EG
   Update". Fixed the routing in `parse_eurasia_email()`, and fixed a real bug found in the same
   pass: bullet points in this format render as a lone `"*"` on their own line, which an earlier
   fix had wrongly treated as an end-of-section marker — it would have truncated every section to
   just its heading the moment the routing fix was applied. Both fixed together; verified against
   real data (6 and 8 sections from the two real captured samples, no content loss).

## Claude Skill added 2026-08-07 (conversational access, alongside the scheduled task)

`expert-commentary-scraper/.claude/skills/run-update/SKILL.md` — a Claude Code Skill that wraps
the existing scripts (no reimplemented logic) so anyone with Claude Code can just ask, in plain
language, to run the update or check its status, instead of needing to know script names. Ships
inside the same distributable zip, under `.claude/skills/`, so it auto-discovers with **zero
install step** the moment a teammate opens Claude Code with the extracted `expert-commentary-
scraper` folder as their working directory (confirmed via research into Claude Code's actual skill
discovery mechanism: project-scoped `.claude/skills/` works without a git repo, just needs to be
inside the launch directory — see the memory file for the full citation trail).

Runs completely independently of, and never conflicts with, the Windows Scheduled Task — both just
append-and-dedupe into the same shared workbook. Bundles one genuinely new reusable script,
`check_sync.py` (a proper version of the row-by-row canonical/review/team-copy/mirror comparison
that got redone by hand a few times this session), plus `references/troubleshooting.md` — a
condensed, team-facing version of the known-issues knowledge in this file (Chatham House's
standing 403, WEF's low publish frequency vs. the 30-day window, the review-file Excel-lock
pattern, the Signum/Eurasia exclusion rules) so a teammate's Claude session isn't starting from
zero on issues already diagnosed here.

## How to proceed

Start a new conversation in this same project folder — memory should pick up the condensed
context automatically. If anything's unclear, point it at this file or
`expert-commentary-scraper.md` in the memory folder directly.

## Small things still open

- Vriens & Partners: still no real email sample — when one arrives, check the generic parser
  handles it reasonably; may want a dedicated parser module if its format is distinctive
- Two minor single-row CTA leaks not fixed (low priority): Atlantic Council "...sign up here",
  Deloitte Insights "...Subscribe today" — both embedded in otherwise-substantial real content
- CSIS trailing author-bio lines not filtered (judged too name-specific to safely generalize)
- 9am-ET-via-6am-Pacific scheduled task reliability not yet confirmed over multiple real days
  (only just registered today)
- Not yet git-committed (same standing item as the Iran tool)
- The `expert-commentary-scraper.md` memory file has grown to ~500 lines today — worth a
  consolidation pass (like the Iran tool's memory got once) in a calmer future session
