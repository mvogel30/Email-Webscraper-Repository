# Known issues and how to tell them apart from real problems

## "A source shows 0 results" isn't automatically broken

Check whether the per-source entry in the run output is a plain number (0 is a real, valid count)
or an `{"error": ...}` object (a genuine failure). Two sources are worth knowing about specifically:

- **Chatham House** used to 403 plain HTTP requests intermittently (worked 8/10, 8/12, 8/13;
  failed 8/11, 8/14 - a WAF/bot-detection block, not a real ban). As of 2026-08-14, this source
  fetches via a real headless Chromium session (Playwright) instead of a plain HTTP request - the
  same fix already used for WEF, confirmed to get a clean 200 and 10 real articles every time
  tested, including on a run where the plain-HTTP path was still 403ing. If you ever see an `error`
  entry for this source again, it's a new/different failure, not the old intermittent block - look
  at the actual error message rather than assuming it's the known issue.
- **WEF** can legitimately show `0` (not an error) for days at a time - it only keeps articles with
  a real, parseable publication date, and the site simply doesn't publish very often. Before
  assuming this is broken, check how old WEF's most recent real article actually is versus the
  lookback window (30 days by default) - if the newest one on the page is already older than that,
  0 is correct, not a bug.

## A run added 0 new rows - also not automatically broken

Every run re-checks everything within its lookback window, and existing content is always
correctly skipped as a duplicate (matched by URL, or by source+title+date for emails with no
public URL). 0 added just means nothing genuinely new was found - normal on back-to-back runs, or
if nothing new was published/sent since the last one.

## Review file / team copy not updating that you can see

If a "review copy refresh" or "team copy refresh" step reports a file-lock warning instead of
succeeding: **someone almost certainly has that file open in Excel right now.** OneDrive can't sync
a file that's open, and a script can't overwrite it either. Ask whoever has it open to close it,
then re-run (or just wait for the next scheduled run). This is the single most common thing that
looks like "the tool is broken" and isn't - the canonical file itself was still updated correctly,
only the copy-out step was skipped.

## Newsletter/event exclusions already built in - don't re-flag these as missing content

Some vendor emails are deliberately excluded because they're invitations/announcements, not
research content, not because of a bug:
- Signum: "SIGNUM Zoom", "SIGNUM Investor Trip", "SIGNUM Policy Day", and a general "SIGNUM Event"
  catch-all (anything shaped like `SIGNUM <word(s)>:` - a location/descriptor before the colon,
  as opposed to real reports which are always `SIGNUM: <topic>: ...` with the colon immediately
  after "SIGNUM").
- Eurasia Group: "EG Conference Call" invitations.
- Any vendor email that's a reply/forward (subject starts with "RE:"/"FW:"/"Fwd:") - treated as
  personal correspondence, not a newsletter broadcast, across every vendor.

If a genuinely new kind of announcement/invitation shows up that isn't caught by these, that's a
real gap worth fixing in the relevant vendor parser (`email-scraper/signum_parser.py`,
`email-scraper/eurasia_parser.py`) - not something to work around by hand each time.

## Workbook schema, if column meanings are unclear

Leftmost columns: **Report** ("Both"/"Spotlight only" - which of the team's weekly reports a
source feeds) → **Scraping-Type** ("Web"/"Email") → **Source** (the organization) →
**Newsletter Name** (the specific named publication/series, e.g. "Dispatches", "EG Update", "EG
Daily Brief" - some vendors run several differently-named formats, computed per-email by that
vendor's own parser).

One structural thing worth knowing: Eurasia Group's "EG Update" and "EG Daily Brief" formats are
both rich, multi-section digests that get split into one row PER topic/section (not one row per
email) - a single incoming email can produce 3-8+ rows, each with its own title, geography, and
full text for just that section.
