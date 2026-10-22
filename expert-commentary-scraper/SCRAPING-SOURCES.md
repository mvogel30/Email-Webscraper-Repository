# Scraping Sources & Advisors Reference

Standalone reference document — not read by any script, safe to edit or delete without affecting
the tool. Lists every source the Expert Commentary Scraper pulls from, as configured in
`sources_config.json` and `email-scraper/vendors.json`. No topic/keyword filter is applied on
either side — every article/newsletter found is kept, subject only to the specific exclusions
noted below (invitations/events, reply-forward threads).

Last checked against the live config: 2026-08-17.

## 1. Websites scraped (web sources)

| Source | URL | What's scraped | Notes |
|---|---|---|---|
| Atlantic Council | [atlanticcouncil.org/dispatches](https://www.atlanticcouncil.org/dispatches/) | Title, link, publish date, full article body | "Dispatches" listing |
| CSIS | [csis.org/analysis](https://www.csis.org/analysis) | Title, link, publish date, full article body | "Analysis" listing |
| CFR | [cfr.org/expert-takes](https://www.cfr.org/expert-takes) | Title, link, publish date, full article body | "Expert Takes" listing |
| Chatham House | [chathamhouse.org/search/content_format/Expert Comment](https://www.chathamhouse.org/search/content_format/Expert%20comment) | Title, link, publish date, full article body | Fetched via a real headless browser session (Playwright), not a plain HTTP request — the site blocks plain HTTP clients with an intermittent 403 (bot detection), but not real browsers |
| Lowy Institute | [lowyinstitute.org/the-interpreter (RSS)](https://www.lowyinstitute.org/the-interpreter/rss.xml) | Title, link, publish date, full article body | Via RSS feed |
| China Macro Group | [chinamacro.ch/insights](https://www.chinamacro.ch/insights) | Title, link, listing excerpt (only month-level date for most posts) | "Insights" listing (Wix). Full article body is not reachable via a plain HTTP fetch — the rich content lazy-loads client-side, so only the listing's own excerpt is captured; this is a real ceiling of the current approach, not a bug |
| Deloitte Insights | [deloitte.com/us/en/insights](https://www.deloitte.com/us/en/insights/) (via Deloitte's public sitemap, not this page directly) | Title, link, publish date, full article body | No usable dated listing page exists (client-side rendered), so this source instead reads Deloitte's public XML sitemap and visits each recent `/insights/` candidate page. Only pages confirmed to carry real `Article`/`NewsArticle`/`BlogPosting` structured data are kept — added 2026-08-14 to stop topic/tracker/dashboard pages (e.g. "ConsumerSignals") from being picked up as if they were dated articles |
| World Economic Forum (WEF) | [weforum.org/publications](https://www.weforum.org/publications/) | Title, link, publish date, full text extracted from the linked PDF report | Listing fetched via a real headless browser session (Playwright) — the site blocks plain HTTP clients entirely. As of 2026-08-17, most listing cards have no description at all (a majority, not a rare case — 12 of 19 checked), so "Full Text" now comes from the linked PDF report itself: each article's own page is visited (also via Playwright, same block applies there) to find its "Download PDF" link, then the PDF is downloaded directly (`reports.weforum.org` has no bot-block at all, confirmed 5/5) and its text extracted (`pypdf`). Falls back to the listing description, then a visible "nothing available" placeholder, only if no PDF is found. Reports run 75K-145K real characters — always truncated to Excel's 32,767-char cell limit with a visible marker, same as any other oversized row |

## 2. Advisors scraped via email (vendor newsletters)

Matches ANY sender at the vendor's domain — not specific people, not subject-line tags. Any email
whose subject starts with `RE:` / `FW:` / `Fwd:` is excluded across every vendor below (personal
correspondence, not a newsletter broadcast).

| Advisor | Domain matched | Newsletter formats scraped | Excluded (not scraped) |
|---|---|---|---|
| Signum | `signumglobal.com` | "Signum" (named country/topic reports, e.g. "SIGNUM: Iran: ...") | Event/invitation emails: "SIGNUM Zoom", "SIGNUM Investor Trip", "SIGNUM Policy Day" (any city), and a general "SIGNUM Event" catch-all for the shared `SIGNUM <word(s)>:` invitation shape |
| Eurasia Group | `eurasiagroup.net` | "EG Update" and "EG Daily Brief" — both rich multi-section digests, split into one row per topic/section (a single email can produce several rows); "Daily Research Summary"; generic "Eurasia Group" internal threads | "EG Conference Call" (call invitations) |
| Middle East Council on Global Affairs | `mecouncil.org` | "Daily Brief" (subject pattern `Daily Brief \| <date>`); anything else kept whole under the advisor's own name | None defined yet — only one real email sample seen so far (2026-08-05) |
| Vriens & Partners | `vrienspartners.com` | Kept whole under the advisor's own name (no dedicated newsletter-name detection yet) | None defined yet — domain confirmed correct via web search, but no real email has arrived in this mailbox yet to see what their format actually looks like |

## Keeping this current

This is a snapshot, not a live view. If a source is added, removed, or its scraping behavior
changes, this file needs a manual update — check `sources_config.json` (web sources) and
`email-scraper/vendors.json` + the individual `email-scraper/*_parser.py` modules (advisor
newsletter formats and exclusions) for the current ground truth.

The `.docx` version of this document (`Scraping Sources and Advisors Reference.docx`, same
folder) regenerates automatically whenever this file is saved — a PostToolUse hook
(`.claude/settings.json`) runs `build_scraping_sources_docx.py` after every edit. So: edit this
`.md` file when a source changes, and the `.docx` stays in sync on its own — no separate manual
step needed for that half.
