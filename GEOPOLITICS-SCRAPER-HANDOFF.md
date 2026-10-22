# Geopolitics Article Scraper — Handoff Notes (2026-07-28)

This is a plain-English recap of everything built so far, for picking this up in a new
conversation. Claude's own persistent memory should also surface this context automatically in a
new thread (it's saved and was just cleaned up/condensed) — but this file is here so **you** have
something to read too, and as a durable reference in the repo itself.

## What this tool does

Automatically pulls the newest **Iran-related** articles from several think tanks/advisors, plus
your Signum Global Advisors emails (from Andrew Bishop), into one shared Excel workbook your whole
PP+G team reads.

## Where everything lives

**The code:** `geopolitics-article-scraper/` in this repo, plus a `signum-email-scraper/`
subfolder for the email pipeline.

**The shared Excel workbook (open this one):**
```
C:\Users\mivogel\OneDrive - Deloitte (O365D)\Global PP+G - 06 Sensing (Inputs)\
    10 - Geopolitics Sensing & Scraping Tools\
        Iran Scraping example\
            Geopolitics Article Repository (review).xlsx
```
(2026-08-04: this whole tool's output moved one level deeper, into a new "Iran Scraping example"
subfolder, to make room for a separate "Webscraping & Email Scraping tool" folder sitting next to
it for a different scraping project. The parent folder itself was also renamed from
"10 - Geopolitics Sensing (Think Tank Web Articles)" to "10 - Geopolitics Sensing & Scraping Tools"
the same day. All scripts were repointed to match both changes.)
(Don't open `(auto-managed).xlsx` in that same folder — that's the one the scripts write to
directly; the review copy is the safe, read-only one for people to actually open.)

**Run it manually anytime:** double-click `Run Geopolitics Scraper.cmd` in the project folder.
Safe to run as often as you like — it never adds the same article twice.

**Automated daily run:** Windows Scheduled Task "Geopolitics Article Scraper - Daily 9AM". *Not
100% confirmed reliable yet* — worth checking Task Scheduler occasionally to see if it's actually
firing on its own each morning, versus only running when triggered manually.

**For teammates to install:** two ready-to-send zips sit in a `Setup` folder right next to the
workbook (`...\10 - Geopolitics Sensing & Scraping Tools\Iran Scraping example\Setup\`) —
one pre-set for 9 AM Eastern, one for 6 PM India Standard Time. Point US/India colleagues there.

## Sources currently scraped (topic: Iran)

Atlantic Council, CSIS, CFR, Chatham House, Lowy Institute, Middle East Council on Global Affairs,
China Macro Group — plus your Signum and Eurasia Group email reports. Each has a real "Full Text"
column with the actual article body, not just a short snippet (this was a whole separate fix).

**Eurasia Group (added 2026-07-29):** a third pipeline, `eurasia-email-scraper/`, mirroring the
Signum one. Built on top of the reference newsletter tool your colleague built
(`email-scraper-repository build/backtest_newsletters.py` — its `parse_eurasia_daily_brief`
logic already handled EG's format). Differences from Signum: EG doesn't tag Iran in the subject
line, and sends more than one email format, so the export step pulls every email from any
`@eurasiagroup.net` sender, and the parser itself decides Iran-relevance — for "EG Daily Brief"
emails it splits the brief into its per-country sections and keeps only the Iran-relevant ones;
for any other EG email (e.g. "EG Update") it keeps the whole email only if it mentions Iran.
Writes into the same shared team workbook as Signum and the web scraper (you confirmed this is
fine, same as Signum's actual behavior today despite a stale docstring in its code claiming it's
private-only). Both team-install zips were rebuilt and re-copied to `Setup\` to include it.

## The big saga this session: OneDrive sync

For about a week, the shared workbook path we were all using turned out to be a silently-broken
local folder — it *looked* completely normal (real files, real timestamps) but had lost its actual
connection to OneDrive, so nothing written there ever reached SharePoint or the team. We found and
fixed this by re-adding the "Add shortcut to OneDrive" connection specifically on the
`06 Sensing (Inputs)` folder (not its parent library), which created a genuinely working new local
folder. Everything is now repointed there and verified working with a live round-trip test.

**Leftover:** the old broken folder (`Deloitte (O365D)\Global PP+G - Documents\...`) still exists
on disk with about a week of local-only test data that never made it anywhere. Harmless, just
clutter — say the word if you want it cleaned up.

## Also explored: fully cloud-based (no laptop needed)

We looked into running this via Power Automate + Office Scripts instead of a local Windows task,
which would remove the "someone's laptop has to be on" dependency entirely. Blocked by a missing
Power Automate Premium license on your account (confirmed directly — hit the license wall trying
to save a test flow with an HTTP action). Worth revisiting if you ever get that license (via a
free trial or an IT ask) — the architecture is already sketched out (mirrors your boss's own
cloud-flow tool for the newsletter scraper).

## How to proceed

Just start a new conversation in this same project folder — Claude's memory system should pick up
the condensed context automatically. If anything seems off or it asks questions it should already
know the answer to, you can point it at this file directly (`GEOPOLITICS-SCRAPER-HANDOFF.md` in
the repo root) or just paste in whatever's unclear.

## Small things still open

- Confirm the 9 AM automation is actually firing on its own (not just when triggered manually)
- Decide whether to clean up the old broken folder's leftover test data
- The `Ukraine` template in `topics.json` was just an example — untested, would need real URLs
- Never got around to a `git commit` of the project folder (offered a few times)
- `signum-email-scraper/append_signum_workbook.py`'s docstring still claims Signum writes to a
  private workbook, which is stale/wrong (it writes to the shared one, same as Eurasia now) —
  worth fixing the comment sometime, not urgent
- Eurasia Group's Iran-relevance keyword list (`IRAN_KEYWORDS_RE` in `eurasia_parser.py`) is a
  first pass — watch the first few real runs for false negatives (an Iran story that used none of
  the listed terms) or false positives, and tune it
