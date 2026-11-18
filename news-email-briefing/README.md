# Daily News Email Briefing

Reads today's rows from the shared Expert Commentary Repository workbook (built by the sibling
[`expert-commentary-scraper`](../expert-commentary-scraper/) tool), has an LLM group them into an
executive-level briefing, and opens it as an Outlook draft for manual review and send. Full
context, known issues, and current status: see
[`NEWS-EMAIL-BRIEFING-HANDOFF.md`](../NEWS-EMAIL-BRIEFING-HANDOFF.md) at the repo root.

## Setup

```powershell
pip install -r requirements.txt
```

Outlook desktop must be installed, open, and logged in — email is sent via `win32com`, not SMTP.

### Required environment variable

- `OPENAI_API_KEY` — your OpenAI API key.

### Optional environment variables

- `OPENAI_BASE_URL` — defaults to `https://api.openai.com/v1`. Set to
  `https://us.api.openai.com/v1` only if your organization's OpenAI key requires that regional
  host (see the DNS-workaround comment near the top of `daily_news_summary.py` — Deloitte's
  network blocks resolving that host directly; delete that block if you don't need it).
- `OPENAI_MODEL` — defaults to `gpt-5.4`.

### Config block (top of `daily_news_summary.py`)

Edit these for your environment before running:

| Setting | Purpose |
|---|---|
| `EXCEL_PATH` | Full path to the repository workbook to read from |
| `SHEET_NAME` / `HEADER_ROW` | Which worksheet/header row holds the article rows |
| `LOG_PATH` | Where to append the run log |
| `USED_ARTICLES_LOG_PATH` | JSON ledger used for the 1-day leftover carry-forward (see below) |
| `YOUR_EMAIL` | The account that drafts are opened in and reviewed by |
| `BOSS_EMAIL` / `DISTRIBUTION` | CC recipients once the draft is actually sent |
| `MIN_ITEMS_FOR_BRIEFING` | Below this many items for the day, skip the full send and notify `YOUR_EMAIL` only |

## Run it

```powershell
python daily_news_summary.py
```

In production this is driven entirely by a Windows Scheduled Task (Mon–Fri only) that calls
`run_briefing.ps1`, which sets `OPENAI_BASE_URL` and invokes the script — see the handoff doc for
the exact task name/schedule on the current machine.

## How a run behaves

1. Reads today's dated rows from `EXCEL_PATH`, plus any of **yesterday's** rows that weren't
   already recorded as used in a prior day's draft (1-day leftover carry-forward only, with an
   explicit overlap check logged every run).
2. If there are fewer than `MIN_ITEMS_FOR_BRIEFING` rows, skips the real briefing and instead
   emails `YOUR_EMAIL` only, listing the thin content found.
3. Otherwise sends the rows to the LLM to draft a briefing (Executive Summary + Key Themes,
   quotes attributed by the workbook's `Source` column — never `Newsletter Name`, never an
   individual's name).
4. Runs two independent quality gates before anything reaches the full distribution list:
   - a fast structural check (required sections present, not suspiciously short, no leaked
     sentinel text)
   - a second LLM pass that reviews the draft for incoherence, truncation, or anything that would
     look broken to a wide executive audience
   - Either failure sends a private notice (with the raw draft inline) to `YOUR_EMAIL` only — the
     full list never sees a broken briefing.
5. On success, **opens the fully-addressed email as an Outlook draft** (`Display()`, not `Send()`)
   — every day requires a manual click to actually send to the distribution list. Nothing is ever
   auto-sent to anyone but `YOUR_EMAIL`.
6. Records which articles were drafted today so tomorrow's run knows what's a genuine "leftover."

## Data-governance note

This tool sends the day's article text to an external OpenAI API for summarization. Confirm your
organization's policy on routing this content through whichever OpenAI account/key is configured
before pointing it at production data.
