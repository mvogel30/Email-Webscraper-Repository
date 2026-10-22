# Expert Commentary Scraper — set up on your machine

**What this is:** a tool that checks several think tanks/advisors/expert-commentary hubs every
day for new articles (no topic filter — everything they publish), plus vendor emails (Signum,
Eurasia Group, Middle East Council on Global Affairs, Vriens & Partners, whichever ones you
personally receive), and adds them to one shared Excel file the whole team reads. Michael set it
up and is asking a few of us to also run it on our own laptops, so the file keeps updating even
on days he's out. It's completely safe to install — it only reads public websites and your own
Outlook (for senders you already receive), and adds rows to the shared workbook; it doesn't touch
anything else.

**Getting future updates:** this package lives in the same shared folder as the workbook itself —
`Webscraping & Email Scraping tool \ Old Process & Scope \ Setup` (see below for exactly where that
is). This zip is rebuilt automatically every day (as a step in the same daily scheduled task that
runs the scraper itself, on Michael's machine) from whatever the project's code currently is — so
it's always current within a day of any change, with nobody needing to remember a manual rebuild
step. If the tool is ever updated, just grab the current zip from that Setup folder and re-run the
installer (safe to re-run any time — it just re-registers the same daily task with the newer code).

*(2026-08-23: this moved again — the team's visible location is now under the Global Public Policy
& Geopolitics Teams site: **Shared Documents → 03 Cross-Team Projects → 08 - Geopolitics Sensing &
Scraping Tools → Webscraping & Email Scraping tool**. If your own shortcut still points at an older
location, remove it and add a fresh one per "Before you start" below. The older location's files
were left in place, untouched, as a fallback — so nothing was lost in the move.)*

*(2026-10-21: the folder now has two subfolders inside it — **"Old Process & Scope"** (everything
below this point refers to that one - the live pipeline) and **"New Process & Scope"** (reserved
for a future replacement approach, nothing to look at there yet).)*

## Before you start (one-time, 2 minutes)

Make sure the **"Webscraping & Email Scraping tool"** folder itself is synced to your OneDrive.
**Important:** add the shortcut on that specific folder, not a parent folder — adding it at the
wrong level can silently create a local folder that looks right but never actually syncs.

1. Open Teams (or SharePoint) and find the **Global Public Policy & Geopolitics** site.
2. Navigate **into** `Shared Documents` → `03 Cross-Team Projects` →
   `08 - Geopolitics Sensing & Scraping Tools` → `Webscraping & Email Scraping tool` (all the way
   in — don't stop at a parent level).
3. While inside that folder, click **"Add shortcut to OneDrive"** in the toolbar.
4. Wait a minute, then check File Explorer under your own OneDrive folder for a new folder
   (something like **"Webscraping & Email Scraping tool"** — if OneDrive already has a
   differently-located folder of the same name synced, it may add a "1" or "2" suffix; if so,
   just use the exact name that actually appears).

*(Skip this if you already have it synced — most of the team does. If you're not sure, just try
running the installer — it'll tell you clearly if the folder isn't found.)*

## Install (a few minutes)

1. You should have received this whole `expert-commentary-scraper` folder (zip or shared link).
   Unzip it somewhere normal, like your Desktop or Documents.
2. Open the `team-install` folder inside it.
3. Double-click **`Install (double-click).cmd`**.
4. It will ask: *"What time should this run every day?"* — type a time in 24-hour format, e.g.
   `08:30` for 8:30 AM. Pick whatever suits your normal morning (no need to coordinate with
   anyone — everyone can use a different time).
5. It checks/installs anything it needs (Python, a few small libraries, and a small headless
   browser used only for one source), registers the daily task, and runs once immediately so you
   can see it work. **This takes noticeably longer than a typical installer the first time** —
   there's a real ~300MB browser download involved (see "Why the extra download?" below).
6. When you see **"All set."**, you're done — nothing else to do, ever, unless you want to change
   the time later (just run the installer again with a new time).

### Why the extra download?

One of the sources (World Economic Forum) actively blocks normal web requests, but loads fine
through a real browser — so this tool drives an actual (headless, no visible window) Chromium
browser just for that one source. Everything else works exactly like a normal web request. This
download only happens once, not on every daily run.

## Where to see the results

The shared file everyone reads is:
```
Webscraping & Email Scraping tool \ Old Process & Scope \ Expert Commentary Repository (review).xlsx
```
(under wherever the "Webscraping & Email Scraping tool" shortcut landed in your own OneDrive)
Open it anytime in Excel — new articles simply appear as rows, with a `Link` column for the URL,
and a `Report` column on the far left tagging which of the team's weekly reports each row feeds
("Both" or "Spotlight only").

There's also a second copy in the same folder, **`Expert Commentary Repository (team copy).xlsx`**
— same content, refreshed on every run just like the review file. If you're viewing files through
Teams rather than a synced local OneDrive folder and the review file looks stale, check this one
instead: since nobody keeps it open in Excel, it isn't blocked from syncing up to Teams the way a
file someone has open locally can be.

## Running it manually (optional)

If you ever want to trigger an update yourself instead of waiting for your scheduled time,
double-click **`Run Expert Commentary Scraper.cmd`** in the main project folder (one level up
from `team-install`). Running it extra times is always safe — it never adds the same article
twice.

## Running it through Claude Code (optional, no extra setup)

If you have Claude Code, this folder already includes a Claude Skill (`.claude/skills/run-update/`)
that lets you just ask for it in plain language instead of remembering a script name. Open Claude
Code with this `expert-commentary-scraper` folder as your working directory, then just say
something like "run the expert commentary scraper" or "is the workbook up to date?" — Claude picks
it up automatically (no install step, no configuration). It runs the exact same scripts as the
manual/scheduled paths above; nothing extra to maintain.

## If something goes wrong

- The installer prints clear error messages — re-read what it said; the most common issue is the
  shared library not being synced yet (see "Before you start" above).
- Each run's raw results are saved in `Documents\Expert Commentary Repository\daily_runs\` on
  your machine, so nothing is ever lost even if a particular run has trouble writing to the
  shared file — it can always be re-applied later.
- You can ping Michael (mivogel@deloitte.com) if you're stuck.
