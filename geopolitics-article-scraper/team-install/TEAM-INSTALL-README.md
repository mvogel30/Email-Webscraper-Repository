# Geopolitics Article Scraper — set up on your machine

**What this is:** a tool that checks several think tanks and advisory firms every day for new Iran
articles, and adds them to one shared Excel file the whole team reads. Michael set it up and is
asking a few of us (some in the US, some in India) to also run it on our own laptops, so the file
keeps updating even on days he's out. It's completely safe to install — it only reads public
websites and adds rows to the shared workbook; it doesn't touch anything else.

**Getting future updates:** this package lives in the same shared folder as the workbook itself —
`06 Sensing (Inputs) \ 10 - Geopolitics Sensing & Scraping Tools \ Iran Scraping example \
Setup`. If the tool is
ever updated, a newer zip will appear there; just grab it and re-run the installer (safe to
re-run any time — it just re-registers the same daily task with the newer code).

## Before you start (one-time, 2 minutes)

Make sure the **"06 Sensing (Inputs)"** folder is synced to your OneDrive. **Important:** add the
shortcut on that specific folder, not its parent "Global PP+G - Documents" library — adding it at
the wrong level can silently create a local folder that looks right but never actually syncs
(this happened once and took a while to track down).

1. Open Teams (or SharePoint) and find **Global PP+G - Documents**.
2. Navigate **into** `06 Sensing (Inputs)` (don't stop at the library's top level).
3. While inside that folder, click **"Add shortcut to OneDrive"** in the toolbar.
4. Wait a minute, then check File Explorer under your own OneDrive folder for a new folder
   (something like **"Global PP+G - 06 Sensing (Inputs)"**).

*(Skip this if you already have it synced — most of the team does. If you're not sure, just try
running the installer — it'll tell you clearly if the folder isn't found.)*

## Install (2 minutes)

1. You should have received this whole `geopolitics-article-scraper` folder (zip or shared link).
   Unzip it somewhere normal, like your Desktop or Documents.
2. Open the `team-install` folder inside it.
3. Double-click **`Install (double-click).cmd`**.
4. It will ask: *"What time should this run every day?"* — type a time in 24-hour format, e.g.
   `08:30` for 8:30 AM. Pick whatever suits your normal morning (no need to coordinate with
   anyone — everyone can use a different time).
5. It checks/installs anything it needs (Python, a few small libraries), registers the daily
   task, and runs once immediately so you can see it work. This can take a couple of minutes the
   first time.
6. When you see **"All set."**, you're done — nothing else to do, ever, unless you want to change
   the time later (just run the installer again with a new time).

## Where to see the results

The shared file everyone reads is:
```
06 Sensing (Inputs) \ 10 - Geopolitics Sensing & Scraping Tools \ Iran Scraping example \
    Geopolitics Article Repository (review).xlsx
```
(under wherever the "06 Sensing (Inputs)" shortcut landed in your own OneDrive)
Open it anytime in Excel — new articles simply appear as rows, with a `Link` column for the URL.

## Running it manually (optional)

If you ever want to trigger an update yourself instead of waiting for your scheduled time,
double-click **`Run Geopolitics Scraper.cmd`** in the main project folder (one level up from
`team-install`). Running it extra times is always safe — it never adds the same article twice.

## If something goes wrong

- The installer prints clear error messages — re-read what it said; the most common issue is the
  shared library not being synced yet (see "Before you start" above).
- Each run's raw results are saved in `Documents\Geopolitics Article Repository\daily_runs\` on
  your machine, so nothing is ever lost even if a particular run has trouble writing to the shared
  file — it can always be re-applied later.
- You can ping Michael (mivogel@deloitte.com) if you're stuck.
