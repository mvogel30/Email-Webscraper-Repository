#!/usr/bin/env python3
"""
Daily Geopolitical News Summary
Reads the repository workbook, groups today's articles by key themes with an LLM, and emails
an HTML briefing via the Outlook desktop app.

Setup (run once):
    pip install pandas openpyxl openai pywin32

Secrets / config are read from environment variables (see README):
    OPENAI_API_KEY   (required)  your OpenAI API key
    OPENAI_BASE_URL  (optional)  defaults to https://api.openai.com/v1
    OPENAI_MODEL     (optional)  defaults to gpt-5.4

Then run:
    python daily_news_summary.py

Windows Task Scheduler handles the daily timing; the script just runs once and exits.
"""

import os
import json
import pandas as pd
from openai import OpenAI
import win32com.client
import time
import datetime
import re

# ─────────────────────────────────────────────────────────────
#  CONFIG — EDIT THESE for your environment
# ─────────────────────────────────────────────────────────────

# Path to the repository workbook the briefing reads from.
EXCEL_PATH = r"C:\Users\mivogel\OneDrive - Deloitte (O365D)\Global PP+G - Webscraping & Email Scraping tool 1\New Process & Scope\Expert Commentary Repository (review).xlsx"
SHEET_NAME = "Repository Rows"   # worksheet holding the articles
HEADER_ROW = 0                   # 0-based header row index passed to pandas (row 1 in Excel)

# Where to write the run log (set to a folder you can write to).
LOG_PATH = r"C:\Users\mivogel\OneDrive - Deloitte (O365D)\Desktop\Tom's News Email Automation Practice\email-briefing\daily_email.log"

# Persisted record of exactly which (Source, Article Title) items were included in each day's
# actually-drafted briefing - lets the next day's run identify yesterday's genuinely-unused
# "leftover" rows (arrived too late to make yesterday's draft) without re-including anything
# already used. Only a 1-day lookback (yesterday only), not an indefinite backlog.
USED_ARTICLES_LOG_PATH = r"C:\Users\mivogel\OneDrive - Deloitte (O365D)\Desktop\Tom's News Email Automation Practice\email-briefing\used_articles_log.json"

# Recipients.
YOUR_EMAIL   = "mivogel@deloitte.com"     # primary To:
BOSS_EMAIL   = ""                         # CC - still blank
DISTRIBUTION = [                          # additional CC recipients
    "stewatkins@deloitte.com",
    "gacorbridge@deloitte.com",
    "jstiller@deloitte.de",
    "victyu@deloitte.com",
    "tsellersclyde@deloitte.es",
    "jeschiu@deloitte.com",
]

# Below this many items, the day's news is too thin for a meaningful briefing - skip the full
# distribution send and just notify yourself instead (see run_daily_job()).
MIN_ITEMS_FOR_BRIEFING = 2

# Secrets & model — from environment variables (never hard-code keys).
OPENAI_API_KEY  = os.environ.get("OPENAI_API_KEY", "")
OPENAI_BASE_URL = os.environ.get("OPENAI_BASE_URL", "https://api.openai.com/v1")
OPENAI_MODEL    = os.environ.get("OPENAI_MODEL", "gpt-5.4")

# Email is sent via the Outlook desktop app (win32com) — no SMTP/password needed.
# Outlook must be open and logged in when the script runs.

# ─────────────────────────────────────────────────────────────
#  Deloitte-network-only workaround (safe to delete elsewhere)
#  Deloitte DNS blocks us.api.openai.com; if you point OPENAI_BASE_URL at that host we
#  route the hostname to a known Cloudflare IP. This only activates for that host, so
#  standard OpenAI users are unaffected. Delete this block if you don't need it.
# ─────────────────────────────────────────────────────────────
if "us.api.openai.com" in OPENAI_BASE_URL:
    import socket
    _orig_getaddrinfo = socket.getaddrinfo
    def _patched_getaddrinfo(host, port, *args, **kwargs):
        if host == "us.api.openai.com":
            host = "162.159.140.245"
        return _orig_getaddrinfo(host, port, *args, **kwargs)
    socket.getaddrinfo = _patched_getaddrinfo


def _load_used_articles_log() -> dict:
    try:
        with open(USED_ARTICLES_LOG_PATH, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _save_used_articles_log(data: dict):
    try:
        with open(USED_ARTICLES_LOG_PATH, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
    except Exception as exc:
        log(f"Could not write used-articles log: {exc!r}")


def _article_key(row_or_df):
    """(Source, Article Title) identity used to track what's been used, as 'Source :: Title'."""
    return row_or_df["Source"].fillna("").astype(str) + " :: " + row_or_df["Article Title"].fillna("").astype(str)


def record_used_articles(date_iso: str, df: pd.DataFrame):
    """Persist which (Source, Article Title) items were actually drafted for date_iso, so the
    next run can identify that date's genuinely-unused leftovers."""
    data = _load_used_articles_log()
    data[date_iso] = sorted(_article_key(df).tolist())
    _save_used_articles_log(data)


def load_todays_news() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, header=HEADER_ROW)

    # New Process & Scope workbook (adopted 2026-11-18) names this column "Publication date"
    # (lowercase "date") instead of the Old Process's "Publication Date" - match case-
    # insensitively and normalize to "Publication Date" so the rest of this function doesn't
    # need to care which workbook schema it's reading.
    _pub_date_col = next(
        (c for c in df.columns if str(c).strip().lower() == "publication date"), None
    )
    if _pub_date_col is not None:
        if _pub_date_col != "Publication Date":
            df = df.rename(columns={_pub_date_col: "Publication Date"})
        df = df.dropna(subset=["Publication Date"])
        df["Publication Date"] = pd.to_datetime(df["Publication Date"], errors="coerce")
    else:
        # Newer workbook schema (seen 2026-08-20) dropped the combined "Publication Date"
        # column - reconstruct it from the separate day/month/year columns instead. Build it
        # numerically (not via string concatenation) - Excel/pandas sometimes stores the day/
        # year as floats (e.g. 24.0), and astype(str) on those gives "24.0"/"2026.0", which
        # silently fails to match a strict "%b %d %Y" text format and produces all-NaT dates.
        df = df.dropna(subset=["Publication- Date", "Publication Month", "Publication Year"])
        _MONTH_NUM = {
            "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
            "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
        }
        month_num = df["Publication Month"].astype(str).str[:3].map(_MONTH_NUM)
        day_num = pd.to_numeric(df["Publication- Date"], errors="coerce")
        year_num = pd.to_numeric(df["Publication Year"], errors="coerce")
        df["Publication Date"] = pd.to_datetime(
            dict(year=year_num, month=month_num, day=day_num), errors="coerce"
        )

    df = df.dropna(subset=["Publication Date"])
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)

    today_rows = df[df["Publication Date"].dt.date == today].copy()

    # Fold in yesterday's leftovers: rows dated yesterday that weren't part of whatever was
    # actually drafted for yesterday (per the persisted used-articles log). 1-day lookback only
    # - this is a one-day carry-forward, not an indefinitely accumulating backlog.
    used_yesterday = set(_load_used_articles_log().get(yesterday.isoformat(), []))
    yesterday_rows = df[df["Publication Date"].dt.date == yesterday].copy()
    if not yesterday_rows.empty:
        yesterday_rows["_key"] = _article_key(yesterday_rows)
        leftover = yesterday_rows[~yesterday_rows["_key"].isin(used_yesterday)].drop(columns=["_key"])
    else:
        leftover = yesterday_rows

    overlap = used_yesterday & set(_article_key(leftover).tolist()) if not leftover.empty else set()
    if overlap:
        log(f"OVERLAP CHECK FAILED: {len(overlap)} leftover item(s) were already in yesterday's "
            f"used-articles log (should be impossible by construction): {overlap}")
    elif used_yesterday or not yesterday_rows.empty:
        log(f"Overlap check passed: 0 overlap between yesterday's {len(used_yesterday)} used "
            f"item(s) and the {len(leftover)} leftover item(s) (of {len(yesterday_rows)} total "
            f"yesterday-dated row(s)) being folded into today's {len(today_rows)} row(s).")

    return pd.concat([leftover, today_rows]).copy()


def build_news_text(df: pd.DataFrame) -> str:
    lines = []
    for _, row in df.iterrows():
        # Cite by the "Source" column (the organization, e.g. "Eurasia Group", "Middle East
        # Council on Global Affairs") rather than "Newsletter Name" (the specific publication,
        # e.g. "EG Daily Brief", "Daily Brief") - Newsletter Name is too generic/ambiguous on its
        # own (multiple orgs publish something called "Daily Brief").
        block = f"Source: {row.get('Source', '')}"
        if row.get("Geography"):
            block += f" | Geography: {row.get('Geography')}"
        if row.get("Topic"):
            block += f" | Topic: {row.get('Topic')}"
        block += f"\nTitle: {row.get('Article Title', '')}\n{row.get('Full Text', '')}"
        lines.append(block)
    return "\n\n---\n\n".join(lines)


def _get_openai_client() -> OpenAI:
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY environment variable is not set. See README.")
    return OpenAI(api_key=OPENAI_API_KEY, base_url=OPENAI_BASE_URL)


def generate_summary(news_text: str, item_count: int, date_str: str) -> str:
    client = _get_openai_client()

    prompt = f"""You are a senior geopolitical analyst. Below are {item_count} intelligence items \
from {date_str} sourced from newsletters including Bloomberg, Eurasia Group, \
Geopolitical Dispatch, and Signum.

Write a professional daily briefing structured as follows. You MUST include both section \
headings below EXACTLY as written ("**EXECUTIVE SUMMARY**" and "**KEY THEMES**"), each on its \
own line, even when there are very few items or very few themes - never skip straight from the \
executive summary into the theme clusters without the "**KEY THEMES**" heading in between.

**EXECUTIVE SUMMARY**
3-4 sentences covering the most critical developments of the and summarizing business impact.

**KEY THEMES**
Group items into 3-6 thematic clusters (e.g. US Trade & Tariffs, \
European Politics, Emerging Markets, China & Asia-Pacific, \
Energy & Commodities, Regulatory & Tech Policy, etc.) - with as few as 1-2 clusters when there \
are only a handful of items. The "**KEY THEMES**" heading itself still always appears, exactly \
once, immediately before the first theme cluster.
For each theme:
- Bold theme heading
- 2-4 executive level-bullet points summarising the key developments

Tone: concise, analytical, senior-executive audience.
Write in declarative, narrative, executive-facing language. Do not narrate from the perspective of the source set or workflow. Do not write “the packet suggests,” “coverage focused on,” “reporting showed,” or similar.
Prefer concrete wording over abstract framing. Avoid abrupt stacked sentences. Avoid catch-all lists that mix unlike categories without structure. Group consequences clearly and causally. Avoid generic verbs such as “matters” where a more concrete consequence can be stated.

Include strong quotations as evidence to support your key points as sub-bullets where relevant, attributed to the organization name only (e.g. "— Eurasia Group", "— Bloomberg", "— Middle East Council on Global Affairs") — never write the word "Source" or a colon in the attribution itself, never attribute to the specific newsletter/publication name, and never attribute to an individual advisor or person's name.
Distribute quotes evenly across all source organizations present in today's items — do not over-index on any single one. If one source dominates the raw content, still seek representative quotes from the others where available.
Default to advisor-only sourcing. Do not introduce external context into a Summary unless the user explicitly asks for it or approves it. If external context is added, it must not displace the advisor-based framing.

Do not include a signature, sign-off, closing remarks, or any author/analyst name anywhere in the briefing. End the output immediately after the last theme's final bullet point — no "regards," no name, nothing after it.

Before writing anything else: check whether ANY item below contains a genuine, reportable \
geopolitical, economic-policy, security, or market-relevant development. Corporate boilerplate, \
"about us" / company-description pages, marketing copy, and similar non-news content do NOT \
count, even if they are the only item(s) present. If NOT EVEN ONE item qualifies as real news, \
respond with EXACTLY this and nothing else, no explanation, no punctuation added: NO_SUBSTANTIVE_CONTENT

--- TODAY'S EXTERNAL GEOPOLITICAL ADVISOR CONSOLIDATED UPDATE ---
{news_text}
"""

    response = client.chat.completions.create(
        model=OPENAI_MODEL,
        max_completion_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )
    choice = response.choices[0]
    if choice.finish_reason == "length":
        raise RuntimeError(
            "Generated briefing was truncated (hit the max_completion_tokens limit) - "
            "treating as a failed generation rather than sending a cut-off briefing."
        )
    return choice.message.content


# Quick, deterministic checks that don't need a model call - catch obviously broken output
# (missing sections, leftover unconverted markdown, near-empty body) before the more expensive
# LLM review pass runs.
_REQUIRED_SECTION_MARKERS = ("executive summary", "key themes")
_MIN_SUMMARY_CHARS = 200


def check_structural_quality(summary: str) -> tuple[bool, str]:
    lower = summary.lower()
    missing = [m for m in _REQUIRED_SECTION_MARKERS if m not in lower]
    if missing:
        return False, f"Missing expected section(s): {', '.join(missing)}"
    if len(summary.strip()) < _MIN_SUMMARY_CHARS:
        return False, f"Briefing is suspiciously short ({len(summary.strip())} chars)"
    if "NO_SUBSTANTIVE_CONTENT" in summary and summary.strip() != "NO_SUBSTANTIVE_CONTENT":
        return False, "Sentinel text leaked into an otherwise-generated briefing"
    return True, ""


def review_summary_quality(summary: str, date_str: str) -> tuple[bool, str]:
    """Second, independent LLM pass: review the generated briefing for quality before it goes
    to the full distribution list. Catches issues a deterministic check can't - incoherent or
    garbled prose, truncated/incomplete thoughts, repeated boilerplate, hallucinated structure,
    or anything else that would look broken or unprofessional to a wide executive audience.
    """
    client = _get_openai_client()
    review_prompt = f"""You are doing final QA on a daily geopolitical briefing before it is sent \
to a wide distribution list of senior executives. Below is the briefing generated for {date_str}.

Check for ALL of the following:
- Incoherent, garbled, or nonsensical text
- Incomplete or abruptly cut-off sentences/sections
- Missing or malformed EXECUTIVE SUMMARY or KEY THEMES sections
- Repeated or duplicated content
- Placeholder, meta-commentary, or text describing the generation process itself rather than \
real content (e.g. "as an AI...", "I cannot...", references to prompts/instructions)
- Anything else that would look broken, unprofessional, or embarrassing if sent to executives

Respond with EXACTLY "QUALITY_OK" (nothing else) if the briefing has none of these problems.
Otherwise respond with "QUALITY_FAIL: " followed by a short one-sentence reason.

--- BRIEFING TO REVIEW ---
{summary}
"""
    response = client.chat.completions.create(
        model=OPENAI_MODEL,
        max_completion_tokens=200,
        messages=[{"role": "user", "content": review_prompt}]
    )
    verdict = response.choices[0].message.content.strip()
    if verdict == "QUALITY_OK":
        return True, ""
    return False, verdict


def markdown_to_html(text: str) -> str:
    # Bold
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)

    # Section headings — styled as a pill/label bar
    def make_heading(m):
        label = m.group(1).strip()
        return (
            f'<div style="margin:18px 0 6px 0;padding:5px 12px;'
            f'background:#1a3a5c;color:#fff;font-size:11px;font-weight:700;'
            f'letter-spacing:1.2px;text-transform:uppercase;border-radius:3px;'
            f'display:inline-block">{label}</div>'
        )
    text = re.sub(r'^#{1,3} (.+)$', make_heading, text, flags=re.MULTILINE)

    # Bullet points — bullets whose content is a quote (starts with " or “)
    # get rendered as indented italic quote lines with no disc marker.
    # Regular bullets get standard styling.
    def make_bullet(m):
        content = m.group(1)
        is_quote = (
            content.lstrip().startswith(('"', '“', '‘', '’'))
            or re.search(r'[—\-]\s*(Eurasia|Bloomberg|Geopolitical|Signum|Reuters|Daily Brief)', content)
        )
        if is_quote:
            return (
                f'<li style="margin:1px 0 1px 40px;list-style-type:none;'
                f'color:#555;font-style:italic;font-size:12.5px">'
                f'{content}</li>'
            )
        return f'<li style="margin:2px 0">{content}</li>'
    text = re.sub(r'^\s*[-•]\s(.+)$', make_bullet, text, flags=re.MULTILINE)
    text = re.sub(
        r'((?:<li[^>]*>.*?</li>\n?)+)',
        r'<ul style="margin:4px 0 8px 20px;padding-left:16px;list-style-type:disc">\1</ul>',
        text, flags=re.DOTALL
    )

    # Quoted lines ("> text") — indented under the preceding bullet, tighter spacing
    text = re.sub(
        r'^>\s(.+)$',
        r'<div style="margin:1px 0 1px 36px;padding-left:8px;'
        r'border-left:2px solid #ccc;color:#555;font-size:12.5px;'
        r'font-style:italic">\1</div>',
        text, flags=re.MULTILINE
    )

    # Paragraphs — single line-height between sections
    text = re.sub(r'\n{2,}', '</p><p style="margin:6px 0">', text)
    text = text.replace('\n', '<br>')

    today = datetime.date.today().strftime("%B %d, %Y")
    return f"""
<html>
<body style="font-family:'Segoe UI',Arial,sans-serif;font-size:13.5px;
  line-height:1.55;max-width:720px;margin:auto;padding:28px 32px;
  color:#1a1a1a;background:#fafafa">

  <!-- Header bar -->
  <table width="100%" cellpadding="0" cellspacing="0"
         style="border-bottom:2px solid #1a3a5c;margin-bottom:16px;padding-bottom:10px">
    <tr>
      <td style="font-size:13px;font-weight:700;color:#1a3a5c;letter-spacing:0.5px">
        EXTERNAL GEOPOLITICAL ADVISOR DAILY NEWS UPDATE
      </td>
      <td align="right" style="font-size:11px;color:#888">{today}</td>
    </tr>
  </table>

  <!-- Body -->
  <div style="background:#fafafa;padding:4px 0">
    <p style="margin:0 0 6px 0">{text}</p>
  </div>

  <!-- Footer -->
  <p style="margin-top:14px;font-size:10.5px;color:#aaa;text-align:right">
    Auto-generated · External Advisor Newsletter Repository
  </p>

</body>
</html>
"""


def send_email(subject: str, body_html: str, notify_self_only: bool = False):
    """Send via the Outlook desktop app using win32com — no SMTP/password needed.

    notify_self_only=True skips the CC distribution entirely (used for the low-content
    heads-up notice, which should never go to the full list).
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = subject
    mail.HTMLBody = body_html
    mail.To = YOUR_EMAIL
    cc_recipients = [] if notify_self_only else [r for r in ([BOSS_EMAIL] + DISTRIBUTION) if r]
    if cc_recipients:
        mail.CC = "; ".join(cc_recipients)
    mail.Send()
    cc_note = f" (CC: {', '.join(cc_recipients)})" if cc_recipients else ""
    log(f"Email sent to {YOUR_EMAIL}{cc_note}")


def open_email_for_review(subject: str, body_html: str):
    """Compose the full-distribution email (To + CC fully populated, exactly as it would be
    sent) and open it on screen via Outlook's own Display() instead of Send() - every day
    requires you to review it and click Send yourself in Outlook. Nothing goes to the
    distribution list automatically, regardless of how the quality checks came out.
    """
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = subject
    mail.HTMLBody = body_html
    mail.To = YOUR_EMAIL
    cc_recipients = [r for r in ([BOSS_EMAIL] + DISTRIBUTION) if r]
    if cc_recipients:
        mail.CC = "; ".join(cc_recipients)
    mail.Display()
    log(f"Draft opened for your review (To: {YOUR_EMAIL}, CC: {', '.join(cc_recipients)}) - "
        f"NOT sent automatically. Review it in Outlook and click Send yourself.")


def log(message: str):
    line = f"[{datetime.datetime.now().isoformat(timespec='seconds')}] {message}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except Exception:
        pass


def load_todays_news_resilient(attempts: int = 16, delay_seconds: int = 60):
    """Read today's rows, retrying on read errors and on an empty result.

    The window is ~15 min because the repository is often edited on another device/the web and
    OneDrive can take several minutes to sync the updated file down to this machine (it preserves
    the original edit timestamp, so the file can look up to date while the local copy is stale).
    On a normal day rows are found on attempt 1 and the email goes out immediately; the wait only
    happens when today's data isn't present yet (locked, mid-sync, or not yet synced down).
    """
    last_count = 0
    for attempt in range(1, attempts + 1):
        try:
            df = load_todays_news()
            last_count = len(df)
            if not df.empty:
                log(f"Attempt {attempt}/{attempts}: found {len(df)} rows for today.")
                return df
            log(f"Attempt {attempt}/{attempts}: 0 rows for today yet; retrying in {delay_seconds}s.")
        except Exception as exc:
            log(f"Attempt {attempt}/{attempts}: error reading repository ({exc!r}); retrying in {delay_seconds}s.")
        if attempt < attempts:
            time.sleep(delay_seconds)
    log(f"Gave up after {attempts} attempts; last row count = {last_count}.")
    try:
        return load_todays_news()
    except Exception:
        return pd.DataFrame()


def send_skip_notice(reason: str, today_str: str, df: pd.DataFrame):
    """Notify yourself only (never the distribution list) that today's briefing was skipped."""
    titles = "".join(f"<li>{t}</li>" for t in df["Article Title"].fillna("(untitled)"))
    notice_html = (
        f"<html><body style=\"font-family:Arial,sans-serif;font-size:13px\">"
        f"<p>{reason}</p>"
        f"<p>Item(s) in the repository for {today_str} ({len(df)} total):</p><ul>{titles}</ul>"
        f"</body></html>"
    )
    send_email(
        f"Daily News Briefing SKIPPED (low content) - {today_str}",
        notice_html,
        notify_self_only=True,
    )


def run_daily_job():
    today_str = datetime.date.today().strftime("%B %d, %Y")
    log(f"Running daily summary for {today_str}...")

    df = load_todays_news_resilient()

    if df.empty:
        log(f"No entries found for today ({today_str}) after retries. Skipping email.")
        return

    if len(df) < MIN_ITEMS_FOR_BRIEFING:
        log(f"Only {len(df)} item(s) found for {today_str} - too thin for a full briefing. "
            f"Notifying {YOUR_EMAIL} only, skipping the distribution list.")
        send_skip_notice(
            f"Only {len(df)} item(s) were found in the repository for {today_str} - too little "
            f"for a meaningful briefing, so the full distribution send was skipped.",
            today_str, df,
        )
        log(f"Notice sent for {today_str} ({len(df)} item(s), below MIN_ITEMS_FOR_BRIEFING={MIN_ITEMS_FOR_BRIEFING}).")
        return

    log(f"Found {len(df)} items. Generating summary...")
    news_text = build_news_text(df)
    summary   = generate_summary(news_text, len(df), today_str)

    if summary.strip() == "NO_SUBSTANTIVE_CONTENT":
        log(f"Model judged none of today's {len(df)} item(s) for {today_str} to be substantive "
            f"news. Notifying {YOUR_EMAIL} only, skipping the distribution list.")
        send_skip_notice(
            f"None of today's {len(df)} item(s) for {today_str} contained a genuine reportable "
            f"development (e.g. corporate boilerplate rather than real news) - GPT-5.4 judged "
            f"there was nothing substantive to brief, so the full distribution send was skipped.",
            today_str, df,
        )
        log(f"Notice sent for {today_str} (content judged non-substantive).")
        return

    ok, reason = check_structural_quality(summary)
    if not ok:
        log(f"Structural quality check FAILED for {today_str}: {reason}. "
            f"Notifying {YOUR_EMAIL} only, skipping the distribution list.")
        send_skip_notice(
            f"Today's briefing for {today_str} failed an automated quality check before "
            f"sending: {reason}. The full distribution send was skipped so nothing broken "
            f"went out; the raw generated text is below for your review.<br><br>"
            f"<pre style=\"white-space:pre-wrap;font-family:inherit\">{summary}</pre>",
            today_str, df,
        )
        log(f"Notice sent for {today_str} (structural quality check failed: {reason}).")
        return

    log(f"Structural quality check passed for {today_str}. Running LLM quality review...")
    ok, reason = review_summary_quality(summary, today_str)
    if not ok:
        log(f"LLM quality review FAILED for {today_str}: {reason}. "
            f"Notifying {YOUR_EMAIL} only, skipping the distribution list.")
        send_skip_notice(
            f"Today's briefing for {today_str} failed the pre-send quality review: {reason}. "
            f"The full distribution send was skipped so nothing broken went out; the raw "
            f"generated text is below for your review.<br><br>"
            f"<pre style=\"white-space:pre-wrap;font-family:inherit\">{summary}</pre>",
            today_str, df,
        )
        log(f"Notice sent for {today_str} (LLM quality review failed: {reason}).")
        return

    log(f"Quality review passed for {today_str}. Opening draft for your review "
        f"(manual approval required before it goes to the distribution list).")
    html      = markdown_to_html(summary)
    subject   = f"External Geopolitical Advisor Daily News Update - {today_str}"

    open_email_for_review(subject, html)
    record_used_articles(datetime.date.today().isoformat(), df)
    log(f"Draft ready for {today_str} ({len(df)} items) - awaiting your manual Send. "
        f"Recorded {len(df)} item(s) as used for today, so they won't be re-offered as "
        f"tomorrow's leftovers.")


def run_daily_job_with_retry(attempts: int = 3, delay_seconds: int = 120):
    """Run the whole job, retrying on transient failures (OpenAI/network/Outlook not ready).

    send_email() is the last step, so a retry only re-sends if sending itself failed.
    """
    for attempt in range(1, attempts + 1):
        try:
            run_daily_job()
            return
        except Exception as exc:
            log(f"Run attempt {attempt}/{attempts} failed: {exc!r}")
            if attempt < attempts:
                time.sleep(delay_seconds)
    log("All run attempts failed.")
    raise SystemExit(1)


if __name__ == "__main__":
    # Windows Task Scheduler handles the daily timing — just run the job and exit.
    run_daily_job_with_retry()
