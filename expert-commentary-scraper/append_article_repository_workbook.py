"""Append stage: read the fetch JSON and append new article rows to the canonical
Excel repository workbook, skipping duplicates.

Reuses the sibling geopolitics-article-scraper tool's conventions verbatim where they apply:
  - the same repository schema (so both workbooks are compatible/familiar),
  - retry_on_lock() for OneDrive sync/Excel locks,
  - header creation + styling,
  - a Run Log sheet.

Duplicate identity uses two independent signals, either of which alone marks an article as
already-seen: the URL (normalized identically to how sources.py stores it), and Source+Title
(a safety net in case a site ever changes its URL scheme for an already-captured story).

Usage:
    python append_article_repository_workbook.py --input-json run.json --workbook repo.xlsx
"""

import argparse
import collections
import datetime as dt
import json
import re
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import sources  # for normalize_url() - one shared definition of "same URL" for fetch AND dedup
from email_common import COUNTRY_ALIASES  # sources.py already puts email-scraper/ on sys.path


# Mirrors the sibling tool's schema (and, further back, the newsletter tool it originated from)
# so both workbooks stay familiar to the same readers - "Report" is the one column this tool adds
# on top of that shared schema (see REPORT_LABEL_BY_SOURCE below).
# 2026-08-05, per feedback from the user's boss: split the original single "Source" column
# (which held "Web"/"Email") from the organization name (which used to live in "Newsletter
# Name"), and add a genuinely NEW "Newsletter Name" column holding the specific named
# publication/series an article runs under (e.g. Atlantic Council's is "Dispatches", Eurasia
# Group's varies per email - "EG Update", "EG Conference Call", etc. - see
# NEWSLETTER_NAME_BY_SOURCE and each vendor parser's own detection logic).
#
# 2026-08-19, per the user: trimmed down from the original 29-column schema (mirrored from the
# newsletter tool this originated from) to just what's actually used - dropped the combined
# "Publication Date" (redundant once split across "Publication- Date"/Year/Month, which are
# kept), "Industry/Sector", "Geography 1" through "Focus area 2" (never populated by this tool -
# leftover passthrough columns from that original schema). Tested first in a standalone test
# copy before being made permanent here. Every remaining internal reference to a real, sortable
# publication date now reconstructs one from the three split columns instead of reading a single
# "Publication Date" field - see _reconstruct_pub_date() below.
REPOSITORY_COLUMNS = [
    "Report",
    "Scraping-Type", "Source", "Newsletter Name", "Publication- Date", "Publication Year",
    "Publication Month", "Article Title", "Link", "Geography", "Full Text",
]

# 2026-10-21, New Process & Scope (see C:\Users\mivogel\.claude\plans\lively-roaming-moth.md): the
# research-department handover workbook uses this flatter structure instead - one real
# "Publication date" value rather than the split Day/Year/Month columns above, modeled on the
# "new column structure" TEST COPY built and approved earlier (build_new_structure_test.py).
# Column order matches that test copy exactly - Link moves to the end, after Full Text.
NEW_STRUCTURE_COLUMNS = [
    "Report",
    "Scraping-Type", "Source", "Newsletter Name", "Publication date",
    "Article Title", "Geography", "Full Text", "Link",
]

# 2026-09-09, per user request: "Events & Calls" gets its OWN schema, distinct from
# REPOSITORY_COLUMNS above - two single-cell real dates instead of the split day/month/year
# columns, because for an event/call this sheet needs to show TWO different dates that Repository
# Rows never had to distinguish: when the event itself happens ("Event Date") vs. when we learned
# about it/it was announced ("Date Announced" - the email's received date, or the day this scrape
# ran, for a web-sourced event page). Both are written as real Excel date values (not split
# text), which also directly fixes an earlier date-ordering complaint that turned out to be a
# stale AutoFilter range (see autosize() below) rather than a real sort bug - but single real
# date cells are also just more scannable than three split columns for a sheet meant to answer
# "when is this happening", which is exactly what prompted this schema change.
# 2026-09-11, per reviewer feedback: added "Event Time" (a dedicated column - was missing
# entirely before) right after "Event Date".
EVENTS_COLUMNS = [
    "Report",
    "Scraping-Type", "Source", "Newsletter Name", "Event Date", "Event Time", "Date Announced",
    "Article Title", "Link", "Geography", "Full Text",
]

_MONTH_ABBREV_TO_NUM = {
    name.lower(): i for i, name in enumerate(
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1
    )
}


def _reconstruct_pub_date(day, month_abbrev, year):
    """Rebuilds a real date from the three split columns that replaced the old single
    "Publication Date" field, for internal sort/dedup use - never written back to a cell."""
    try:
        if not (day and month_abbrev and year):
            return None
        month_num = _MONTH_ABBREV_TO_NUM.get(str(month_abbrev).strip().lower())
        if not month_num:
            return None
        return dt.date(int(year), month_num, int(day))
    except (ValueError, TypeError):
        return None


def _as_date(value):
    """Coerces a cell value that might be a real datetime/date (normal case, once written as a
    proper Excel date) or a leftover ISO string, to a plain date - or None."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str) and value:
        try:
            return dt.date.fromisoformat(value[:10])
        except ValueError:
            return None
    return None


def _row_date(values, columns):
    """Schema-aware sortable date for one row: reconstructs from the split day/month/year columns
    on REPOSITORY_COLUMNS, reads "Publication date" directly on NEW_STRUCTURE_COLUMNS, or reads
    the single "Event Date" cell directly on EVENTS_COLUMNS - callers (existing_keys,
    sort_by_date_desc, append_rows) work against any of the three schemas through this."""
    if "Event Date" in columns:
        return _as_date(values[columns.index("Event Date")])
    if "Publication date" in columns:
        return _as_date(values[columns.index("Publication date")])
    day_i, year_i, month_i = columns.index("Publication- Date"), columns.index("Publication Year"), columns.index("Publication Month")
    return _reconstruct_pub_date(values[day_i], values[month_i], values[year_i])

# Static per-source newsletter/section name for web sources (one hub page = one fixed named
# section, confirmed against each site's own <h1>/<title> - e.g. csis.org/analysis's own H1 is
# literally "Analysis"). Email vendors compute this per-article instead (see their parser
# modules), since the same vendor can run several differently-named formats.
NEWSLETTER_NAME_BY_SOURCE = {
    "Atlantic Council": "Dispatches",
    "CSIS": "Analysis",
    "CFR": "Expert Takes",
    "Chatham House": "Expert Comment",
    "Lowy Institute": "The Interpreter",
    "China Macro Group": "Insights",
    "Deloitte Insights": "Deloitte Insights",
    "WEF": "Publications",
}

# Which report(s) each source feeds, per the user's own "Weekly Development Summaries" /
# "Weekly Spotlight Analysis" breakdown: sources covered by both weekly reports vs. sources that
# only ever feed the Spotlight Analysis. Keyed on the exact "source" values sources_config.json
# and vendors.json use, so this needs a new entry whenever a new source/vendor is added.
REPORT_LABEL_BY_SOURCE = {
    "Atlantic Council": "Both",
    "CSIS": "Both",
    "CFR": "Both",
    "Chatham House": "Both",
    "Middle East Council on Global Affairs": "Both",
    "Eurasia Group": "Both",
    "Signum": "Both",
    "China Macro Group": "Spotlight only",
    "Deloitte Insights": "Spotlight only",
    "Lowy Institute": "Spotlight only",
    "Vriens & Partners": "Spotlight only",
    "WEF": "Spotlight only",
}
# 2026-09-09: web-scraped events/calendar pages (CSIS/Atlantic Council/CFR/Chatham House/WEF/
# Middle East Council) were added per user request, routing to the "Events & Calls" sheet (see
# is_event handling in main() below) rather than Repository Rows. Their articles' "source" is
# deliberately the plain org name (not a distinct "X Events" name) - see the matching comment in
# each extract_*_events() function in sources.py - so they use the SAME REPORT_LABEL_BY_SOURCE
# entries above as that org's regular content; no new entries needed here.

# 2026-08-18, per the user's manager: a handful of specific Eurasia Group newsletter types feed
# Spotlight only, while the rest of Eurasia Group (EG Daily Brief, EG Update, Daily Research
# Summary, and the generic "Eurasia Group" bucket) stays "Both" per REPORT_LABEL_BY_SOURCE above.
# Checked first in article_to_row() below - only these specific newsletter names are overridden,
# everything else falls through to the normal per-source lookup unchanged.
REPORT_LABEL_OVERRIDE_BY_NEWSLETTER_NAME = {
    "EG Political Risk Monitor": "Spotlight only",
    "EG Global Macro": "Spotlight only",
    "EG Conference Call Notes": "Spotlight only",
    "EG Geo-technology": "Spotlight only",
}

DATA_ROW_HEIGHT = 15  # points - standard single-line height, fixed/uniform for every data row

RUN_LOG_COLUMNS = [
    "Run Timestamp", "Runner", "Source JSON", "Articles In JSON", "Rows Added",
    "Rows Skipped (duplicate)", "Per-Source Counts",
]


def parse_args():
    ap = argparse.ArgumentParser(description="Append scraped article rows to a repository workbook.")
    ap.add_argument("--input-json", required=True)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--runner", default="", help="Machine/person identifier, logged in Run Log.")
    ap.add_argument("--allow-create-new-workbook", action="store_true",
                     help="Allow creating a brand-new empty workbook if none exists yet at "
                          "--workbook. Only pass this for a genuine first-time setup - leaving it "
                          "off (the default) makes a missing production workbook a hard error "
                          "instead of silently replacing it with an empty one.")
    ap.add_argument("--row-schema", choices=["legacy", "new-structure"], default="legacy",
                     help="Which Repository Rows column layout to write - 'legacy' (default,"
                          " REPOSITORY_COLUMNS, split Day/Year/Month) for the Old Process, or"
                          " 'new-structure' (NEW_STRUCTURE_COLUMNS, one real Publication date"
                          " column) for the New Process & Scope. Events & Calls is unaffected"
                          " either way - its schema doesn't change between processes.")
    return ap.parse_args()


def retry_on_lock(operation, what, attempts=10, delay_seconds=25):
    """Run a file op, retrying when the file is briefly locked. The workbook lives in
    OneDrive (possibly shared across several teammates' machines) and may be mid-sync
    or open in Excel at run time; retrying rides out that window before giving up."""
    for attempt in range(1, attempts + 1):
        try:
            return operation()
        except (PermissionError, OSError) as exc:
            if attempt == attempts:
                raise
            print(f"{what}: file locked/busy ({type(exc).__name__}); retry {attempt}/{attempts} in {delay_seconds}s", flush=True)
            time.sleep(delay_seconds)


def load_payload(path: Path) -> dict:
    raw = retry_on_lock(lambda: path.read_text(encoding="utf-8-sig"), "read fetch JSON").strip()
    return json.loads(raw) if raw else {"articles": []}


def get_or_create_workbook(path: Path, allow_create: bool = False):
    """2026-09-18, real incident: a bare `if path.exists()` here had NO retry, unlike every other
    OneDrive file op in this script (see retry_on_lock's own docstring on why those need it) - a
    single transient OneDrive sync gap (the file briefly not materialized locally) made an
    EXISTING production workbook with 1,385+/122 real rows report as missing for one instant,
    which silently created a brand-new EMPTY workbook in its place and saved right over it. No
    exception, no warning - just quiet, total data loss, only caught because someone happened to
    check row counts afterward.

    Fixed two ways: (1) the existence check itself now retries through `retry_on_lock`, same as
    every load/save already does, so a transient sync gap self-resolves instead of being read as
    "doesn't exist"; (2) even after retries genuinely exhaust, creating a brand-new workbook now
    requires the caller to explicitly opt in via `allow_create` - which should only ever be True
    for a genuine first-time setup (a brand new team install that has never had a workbook at this
    path). Every normal daily run leaves this False, so a truly-missing production file now fails
    loudly instead of silently replacing itself with an empty one."""
    def check_exists():
        if not path.exists():
            raise FileNotFoundError(str(path))

    try:
        retry_on_lock(check_exists, "check workbook exists")
    except FileNotFoundError:
        if not allow_create:
            raise RuntimeError(
                f"Workbook not found at {path} after retrying - refusing to silently create a "
                "new empty workbook in its place (this exact gap caused real production data "
                "loss on 2026-09-17/18). If this is genuinely a brand-new setup with no prior "
                "workbook at this path, pass allow_create=True / --allow-create-new-workbook "
                "explicitly."
            ) from None
        print(f"WARNING: no workbook found at {path} after retrying - creating a BRAND NEW EMPTY "
              "workbook here because allow_create was explicitly set. If a workbook was EVER "
              "expected to already exist at this path, stop and investigate instead of "
              "proceeding.", flush=True)
        path.parent.mkdir(parents=True, exist_ok=True)
        return Workbook()

    return retry_on_lock(lambda: load_workbook(path), "open workbook")


def get_or_create_sheet(workbook, title):
    if title in workbook.sheetnames:
        return workbook[title]
    # Reuse the pristine default sheet openpyxl creates with a new workbook (named "Sheet"),
    # but only that one - never an existing named sheet that is merely still empty.
    if workbook.sheetnames == ["Sheet"]:
        default = workbook["Sheet"]
        if default.max_row == 1 and default.max_column == 1 and default["A1"].value is None:
            default.title = title
            return default
    return workbook.create_sheet(title)


def ensure_header(sheet, columns):
    current = [sheet.cell(1, c).value for c in range(1, len(columns) + 1)]
    if current == columns:
        return
    if sheet.max_row == 1 and all(sheet.cell(1, c).value is None for c in range(1, max(sheet.max_column, len(columns)) + 1)):
        for index, name in enumerate(columns, start=1):
            sheet.cell(1, index).value = name
        style_header(sheet, len(columns))
        return
    raise ValueError(f"Sheet {sheet.title!r} does not match the expected header.")


def style_header(sheet, max_col):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in sheet[1][:max_col]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions


def article_key(url, source, title, pub) -> str:
    if url:
        return sources.normalize_url(url).lower()
    pub = pub.isoformat() if isinstance(pub, dt.date) else str(pub or "")[:10]
    return f"{(source or '').strip()}|{(title or '').strip()}|{pub}".lower()


def title_key(source, title) -> str:
    """Secondary duplicate signal: same source + exact same title."""
    return f"title|{(source or '').strip()}|{(title or '').strip()}".lower()


def existing_keys(sheet, columns=REPOSITORY_COLUMNS) -> set:
    """Keys already in the workbook so re-runs don't re-append the same article. `columns` picks
    the schema to read against (REPOSITORY_COLUMNS or EVENTS_COLUMNS) - see _row_date()."""
    link_i = columns.index("Link")
    name_i = columns.index("Source")
    title_i = columns.index("Article Title")
    keys = set()
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not values or len(values) <= title_i:
            continue
        url = values[link_i] if len(values) > link_i else ""
        pub_date = _row_date(values, columns)
        keys.add(article_key(url, values[name_i], values[title_i], pub_date))
        if values[title_i]:
            keys.add(title_key(values[name_i], values[title_i]))
    return keys


def latest_dates_by_url(sheet, columns=REPOSITORY_COLUMNS) -> dict:
    """{normalized URL: (most recent stored date, that row's Full Text)} - lets append_rows()
    recognize a "living page" that's been genuinely republished under the same URL (2026-09-23,
    real case: Deloitte Insights' "A view from London" and "Weekly Global Economic Update" get
    periodically refreshed with new content and a new datePublished, but the URL itself never
    changes). Without this, article_key()'s URL branch ignores date entirely and title_key()
    matches on the unchanged title too, so BOTH dedup signals would silently skip the new edition
    forever - the exact reason those two rows had been stuck on stale August content until this
    fix. Only the LATEST date matters (not full history), since that's the only baseline a future
    run needs to compare against.

    2026-10-21, per user feedback: a later date alone isn't enough to prove a "new edition" -
    real case found investigating a Middle East Council "split-up" report: the same top-feature
    essay (identical body, confirmed byte-for-byte) was captured 3 times across 3 consecutive
    days, because each day's email wraps its link in a DIFFERENT tracking redirect, so even the
    normalized-URL key itself differs run to run and this lookup alone wouldn't have caught it
    either - that duplicate is fixed via title_key() + the full_text equality check in
    append_rows() below instead. The Full Text is carried here too so append_rows() can apply the
    same "identical content isn't a new edition, no matter the date" rule uniformly to both the
    URL and title lookups from one shared place, rather than writing the rule twice."""
    link_i = columns.index("Link")
    text_i = columns.index("Full Text")
    latest = {}
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not values or len(values) <= link_i:
            continue
        url = values[link_i]
        if not url:
            continue
        pub_date = _row_date(values, columns)
        if pub_date is None:
            continue
        normalized = sources.normalize_url(url).lower()
        if normalized not in latest or pub_date > latest[normalized][0]:
            latest[normalized] = (pub_date, values[text_i] if len(values) > text_i else "")
    return latest


def latest_dates_by_title(sheet, columns=REPOSITORY_COLUMNS) -> dict:
    """{(source, title): (most recent stored date, that row's Full Text)} - the email-side
    counterpart to latest_dates_by_url() above, for content with NO url at all (regular Eurasia
    Group emails have none - only event invitations get a registration_link). Real case found
    2026-09-23 while investigating a QA report of "occasional one-day date offsets": Eurasia Group
    runs a genuinely recurring feature, "Chart of the day" (and others), reusing the EXACT same
    title across many real, distinct emails with completely different content each time (confirmed:
    11 real editions found, covering currency intervention, manufacturing jobs, Chinese capital
    markets, Central European stocks...). Since these articles have no URL, article_key() falls
    back to source+title+date - which SHOULD already differ per edition - but title_key() (the
    secondary safety-net signal, title-only, no date) still matched every edition after the very
    first one and silently dropped it forever. The reported "date looks one day off" turned out to
    really be "we kept the OLD edition and dropped the new one, so what's stored looks stale by
    whatever gap separates the two" - not a parsing bug at all. Same fix shape as
    latest_dates_by_url(): track the latest known date per (source, title), and let a genuinely
    newer edition through instead of being silently absorbed by title_key().

    2026-10-21, per user feedback: see the matching comment in latest_dates_by_url() - the Full
    Text is carried here too so append_rows() can refuse to treat a later-dated row as a "new
    edition" when the content is actually byte-identical to what's already stored (a title-only
    recurring MEC feature re-sent on consecutive days, not a real new edition like the Eurasia
    Group cases this function was originally built for)."""
    title_i = columns.index("Article Title")
    name_i = columns.index("Source")
    text_i = columns.index("Full Text")
    latest = {}
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not values or len(values) <= title_i:
            continue
        title = values[title_i]
        if not title:
            continue
        pub_date = _row_date(values, columns)
        if pub_date is None:
            continue
        key = (values[name_i], title)
        if key not in latest or pub_date > latest[key][0]:
            latest[key] = (pub_date, values[text_i] if len(values) > text_i else "")
    return latest


# Excel's real hard limit per cell - openpyxl enforces this SILENTLY (no exception, no warning)
# the moment a value longer than this is assigned to a cell, truncating mid-word/mid-sentence
# with zero indication anything was cut. Found 2026-08-05 during a full-text quality audit: one
# real CFR article (38,240 real chars fetched) had been silently cut to exactly 32,767 in the
# workbook. Truncating deliberately here, with a visible marker, at least makes it obvious a cell
# is incomplete and points the reader to the real source instead of ending mid-sentence unmarked.
_EXCEL_CELL_CHAR_LIMIT = 32767
_TRUNCATION_MARKER = "\n\n[...TRUNCATED - full text exceeds Excel's per-cell character limit. See Link column for the complete article.]"


def _fit_to_cell_limit(text: str) -> str:
    if len(text) <= _EXCEL_CELL_CHAR_LIMIT:
        return text
    keep = _EXCEL_CELL_CHAR_LIMIT - len(_TRUNCATION_MARKER)
    return text[:keep] + _TRUNCATION_MARKER


def _full_text_or_placeholder(article: dict) -> str:
    summary = article.get("summary", "")
    # "Full Text" is the real article body; "Summary2" stays the short blurb. Fall back to the
    # short summary only if fetching the real body genuinely failed.
    raw_full_text = article.get("full_text", "") or summary
    if not raw_full_text:
        # Found 2026-08-17: a WEF listing card with a title but no description at all (every
        # other WEF card has both) produced a silently blank row - nothing wrong with the fetch,
        # there was just nothing there to fall back to. Make the gap visible instead of blank,
        # same philosophy as _fit_to_cell_limit's truncation marker below.
        raw_full_text = (
            f"[No summary or article text was available from {article.get('source') or 'this source'}'s "
            "listing page for this entry - see the Link column for the original.]"
        )
    return _fit_to_cell_limit(raw_full_text)


def _geography_for(article: dict) -> str:
    """2026-09-23, per QA feedback: every one of the 8 web sources' regular content-scraping
    paths (Repository Rows) never called geography inference at all - only the separate Events
    extractors did, from an earlier fix. A production check confirmed this precisely: Lowy, CSIS,
    China Macro, CFR, Atlantic Council, Deloitte Insights, Chatham House, and WEF were ALL 100%
    blank on Geography, while email-sourced rows (which already called infer_geography() in
    their own parser modules) were mostly populated already.

    Rather than patch each of the 8 extractors individually with only a short listing blurb, this
    runs centrally here using the FULL article text - richer signal than any extractor's own
    short summary, and means a retroactive fix needs no re-fetching at all (every already-scraped
    row's Full Text is already sitting in the workbook). Only used as a fallback - an article that
    already set its own geography (every email vendor, every Events extractor) keeps that value
    unchanged.

    2026-09-23, per user follow-up: a genuinely-blank Geography cell reads as ambiguous - did the
    tool never check, or did it check and find nothing? Same "make the gap visible instead of
    blank" philosophy _full_text_or_placeholder() already uses for Full Text below - checked
    every remaining blank row directly (a mix of genuinely topic-only content, admin emails, and
    thin bylines, not a bug) and confirmed a placeholder is honest here, not covering up a real
    gap that should instead be fixed."""
    existing = article.get("geography")
    if existing:
        return existing
    title = article.get("title", "")
    text = article.get("full_text") or article.get("summary") or ""
    if article.get("source") in ("Deloitte Insights", "WEF"):
        return _dominant_geography(title, text)
    inferred = sources.infer_geography(f"{title} {text}")
    return inferred or "[No specific geography identified]"


def _dominant_geography(title: str, text: str, top_n: int = 3) -> str:
    """2026-10-01, per QA feedback: infer_geography()'s "first 3 countries mentioned, in document
    order" works well for the other 6 web sources (each piece IS about a specific country/region),
    but Deloitte Insights and WEF mostly publish thematic, multi-region business/economic reports -
    confirmed live on real examples: "A view from London" tagged "US, Iran, Saudi Arabia" (none of
    which is London), "Risk Management, Reimagined: Outlook to 2035" (a generic global report)
    tagged just "South Africa" off one incidental mention. Neither is about that country at all -
    whatever happens to be named earliest in a long report becomes the tag, regardless of relevance.

    Counts mentions instead of using first-position, and requires the top country/region to clearly
    dominate (>=2x the runner-up's count) before trusting it as the real subject - otherwise the
    piece is genuinely multi-region/thematic and "Global" is the honest answer, not an arbitrary
    pick among several equally-incidental mentions."""
    full = f"{title} {text}"
    counts: dict[str, int] = {}
    first_pos: dict[str, int] = {}
    for candidate, canonical in COUNTRY_ALIASES.items():
        matches = list(re.finditer(rf"\b{re.escape(candidate)}\b", full, re.I))
        if matches:
            counts[canonical] = counts.get(canonical, 0) + len(matches)
            first_pos.setdefault(canonical, matches[0].start())
    if not counts:
        return "[No specific geography identified]"
    ranked = sorted(counts.items(), key=lambda kv: -kv[1])
    if len(ranked) > 2:
        top_count, second_count = ranked[0][1], ranked[1][1]
        if top_count < 2 * second_count:
            return "Global"
        ranked = [(g, c) for g, c in ranked if c >= second_count]
    chosen = sorted((g for g, _ in ranked), key=lambda g: first_pos[g])
    return ", ".join(chosen[:top_n])


def article_to_row(article: dict) -> dict:
    pub = article.get("published", "")
    pub_date = None
    if pub:
        try:
            pub_date = dt.date.fromisoformat(pub[:10])
        except ValueError:
            pub_date = None
    row = {column: "" for column in REPOSITORY_COLUMNS}
    row.update({
        "Report": (
            REPORT_LABEL_OVERRIDE_BY_NEWSLETTER_NAME.get(article.get("newsletter_name", ""))
            or REPORT_LABEL_BY_SOURCE.get(article.get("source", ""), "")
        ),
        "Scraping-Type": article.get("source_kind", "Web"),
        "Source": article.get("source", ""),
        # Web sources have one fixed newsletter/section name (looked up here so callers don't
        # each need to know it); email vendors compute a per-article one themselves (varies by
        # email format) and pass it through directly - explicit "newsletter_name" always wins.
        "Newsletter Name": article.get("newsletter_name") or NEWSLETTER_NAME_BY_SOURCE.get(article.get("source", ""), ""),
        "Publication- Date": f"{pub_date.day:02d}" if pub_date else "",
        "Publication Year": str(pub_date.year) if pub_date else "",
        "Publication Month": pub_date.strftime("%b") if pub_date else "",
        "Article Title": article.get("title", ""),
        "Link": article.get("url", ""),
        "Full Text": _full_text_or_placeholder(article),
        "Geography": _geography_for(article),
    })
    return row


def article_to_row_new_structure(article: dict) -> dict:
    """NEW_STRUCTURE_COLUMNS counterpart to article_to_row() - one real "Publication date" value
    instead of the split Day/Year/Month columns, same shape event_to_row() already uses for
    "Event Date"/"Date Announced". Shares every other field/helper with article_to_row() (Report,
    Newsletter Name, Geography, Full Text logic all unchanged) - only the date handling and column
    layout differ."""
    pub = article.get("published", "")
    pub_date = None
    if pub:
        try:
            pub_date = dt.date.fromisoformat(pub[:10])
        except ValueError:
            pub_date = None
    row = {column: "" for column in NEW_STRUCTURE_COLUMNS}
    row.update({
        "Report": (
            REPORT_LABEL_OVERRIDE_BY_NEWSLETTER_NAME.get(article.get("newsletter_name", ""))
            or REPORT_LABEL_BY_SOURCE.get(article.get("source", ""), "")
        ),
        "Scraping-Type": article.get("source_kind", "Web"),
        "Source": article.get("source", ""),
        "Newsletter Name": article.get("newsletter_name") or NEWSLETTER_NAME_BY_SOURCE.get(article.get("source", ""), ""),
        "Publication date": pub_date,
        "Article Title": article.get("title", ""),
        "Geography": _geography_for(article),
        "Full Text": _full_text_or_placeholder(article),
        "Link": article.get("url", ""),
    })
    return row


def event_to_row(article: dict) -> dict:
    """EVENTS_COLUMNS counterpart to article_to_row() - two real single-cell dates (Event Date,
    Date Announced) instead of the split day/month/year columns. Every event-capable parser
    (signum_parser/eurasia_parser/hub_advisor_events_parser/every extract_*_events() in
    sources.py) sets both "event_date" and "announced_date" on the article dict; `published` is
    only a defensive fallback in case either is ever missing."""
    row = {column: "" for column in EVENTS_COLUMNS}
    row.update({
        "Report": (
            REPORT_LABEL_OVERRIDE_BY_NEWSLETTER_NAME.get(article.get("newsletter_name", ""))
            or REPORT_LABEL_BY_SOURCE.get(article.get("source", ""), "")
        ),
        "Scraping-Type": article.get("source_kind", "Web"),
        "Source": article.get("source", ""),
        "Newsletter Name": article.get("newsletter_name") or NEWSLETTER_NAME_BY_SOURCE.get(article.get("source", ""), ""),
        "Event Date": _as_date(article.get("event_date")) or _as_date(article.get("published")),
        "Event Time": article.get("event_time", ""),
        "Date Announced": _as_date(article.get("announced_date")) or _as_date(article.get("published")),
        "Article Title": article.get("title", ""),
        # "registration_link" (email events only - a real but tracking-wrapped, per-send URL,
        # deliberately kept out of dedup) wins when present; "url" (web-scraped events - a
        # stable, real page link) is the normal case otherwise.
        "Link": article.get("registration_link") or article.get("url", ""),
        "Full Text": _full_text_or_placeholder(article),
        "Geography": _geography_for(article),
    })
    return row


# 2026-09-24, per user feedback: CSIS/Atlantic Council event emails (hub_advisor_events_parser.py)
# re-announce the SAME real event under a changing status prefix as it moves through its
# lifecycle - real subjects seen include "CSIS Event Invite: X", "Watch Live Now: X", "Register
# Now: X", "New Event: X", "POSTPONED: X", "CSIS Event Reminder: X", "TOMORROW: X" - and since
# these emails carry no stable URL, the existing title-based dedup (title_key/article_key) treats
# each announcement as a wholly different event, because the FULL title differs by prefix. That's
# how one real event (H.E. Yassine Jaber's CSIS talk) ended up as ~4 separate rows. Stripping the
# prefix is only done for MATCHING - the stored Article Title always keeps whatever prefix the
# real subject had, so "POSTPONED" etc. stays visible.
_EVENT_STATUS_PREFIX_RE = re.compile(
    r"^(?:csis\s+event\s+(?:invite|reminder)|watch\s+live\s+now|live\s+now|register\s+now|"
    r"new\s+event|postponed|tomorrow)\s*:\s*",
    re.I,
)


def _event_core_title(title: str) -> str:
    return _EVENT_STATUS_PREFIX_RE.sub("", (title or "").strip()).strip().lower()


def upsert_hub_advisor_events(sheet, articles: list, columns=EVENTS_COLUMNS) -> tuple:
    """Splits `articles` (hub_advisor_events_parser.py output only - identified by a "<Source>
    Event" newsletter_name, per that module's own naming) into (to_append, remaining): any article
    whose core title (status prefix stripped) matches an EXISTING row's own core title, for the
    same Source, gets that row UPDATED in place instead - Title/Full Text/Event Time/Link/Event
    Date/Date Announced all refreshed to the incoming (latest-known) values, so the row always
    reflects the event's current status rather than accumulating one row per status change.
    `remaining` (everything that found no match) still needs to go through the normal
    append_rows() path to become new rows for genuinely new events."""
    src_i = columns.index("Source")
    title_i = columns.index("Article Title")

    existing_by_key = {}
    for row_cells in sheet.iter_rows(min_row=2):
        source = row_cells[src_i].value
        title = row_cells[title_i].value
        if not source or not title:
            continue
        existing_by_key[(source, _event_core_title(title))] = row_cells

    remaining, updated = [], 0
    for article in articles:
        if not str(article.get("newsletter_name", "")).endswith(" Event"):
            remaining.append(article)
            continue
        key = (article.get("source", ""), _event_core_title(article.get("title", "")))
        row_cells = existing_by_key.get(key)
        if row_cells is None:
            remaining.append(article)
            continue
        new_row = event_to_row(article)
        for column, value in new_row.items():
            row_cells[columns.index(column)].value = value
        updated += 1
    return remaining, updated


def append_rows(sheet, articles, columns=REPOSITORY_COLUMNS, row_builder=None):
    row_builder = row_builder or article_to_row
    seen = existing_keys(sheet, columns)
    # See latest_dates_by_url()/latest_dates_by_title()'s docstrings - lets a genuinely
    # republished "living page" (same URL) or a genuinely recurring same-titled feature with no
    # URL at all (email content) get added as a NEW row instead of being silently skipped forever
    # by URL/title dedup, both of which are blind to a date change on their own.
    latest_by_url = latest_dates_by_url(sheet, columns)
    latest_by_title = latest_dates_by_title(sheet, columns)
    added, skipped = [], 0
    for article in articles:
        # Dedup keys on Event Date for EVENTS_COLUMNS (matching what existing_keys() reads back
        # from the sheet via _row_date()), Publication Date otherwise.
        pub_for_key = article.get("event_date") if "Event Date" in columns else article.get("published", "")
        url = article.get("url", "")
        source = article.get("source", "")
        title = article.get("title", "")
        url_key = article_key(url, source, title, pub_for_key)
        t_key = title_key(source, title)
        incoming_date = _as_date(pub_for_key)
        row = row_builder(article)
        incoming_text = row.get("Full Text", "")
        # 2026-10-21, per user feedback: a later date alone used to be enough to call this a "new
        # edition" (see the docstrings above) - real case found investigating a Middle East
        # Council "split-up" report: the same top-feature essay, byte-identical, got captured 3
        # times across 3 consecutive days, since each day's email just re-links the same still-
        # current feature with a fresh send date. That's not a new edition, so it must ALSO have
        # genuinely different content before the later date is trusted - true new editions
        # (Deloitte Insights' "A view from London", Eurasia Group's "Chart of the day") always do.
        is_new_edition = False
        if url:
            stored = latest_by_url.get(sources.normalize_url(url).lower())
            is_new_edition = bool(stored and incoming_date and stored[0] and incoming_date > stored[0] and incoming_text != stored[1])
        else:
            stored = latest_by_title.get((source, title))
            is_new_edition = bool(stored and incoming_date and stored[0] and incoming_date > stored[0] and incoming_text != stored[1])
        if not is_new_edition and (url_key in seen or t_key in seen):
            skipped += 1
            continue
        seen.add(url_key)
        seen.add(t_key)
        if is_new_edition:
            if url:
                latest_by_url[sources.normalize_url(url).lower()] = (incoming_date, incoming_text)
            else:
                latest_by_title[(source, title)] = (incoming_date, incoming_text)
        values = [row.get(column, "") for column in columns]
        sheet.append(values)
        added.append(article)
    return added, skipped


def sort_by_date_desc(sheet, columns=REPOSITORY_COLUMNS):
    """Re-order every data row so the newest date is at the top (Publication Date for
    REPOSITORY_COLUMNS, Event Date for EVENTS_COLUMNS - see _row_date()). Rows with no date are
    kept at the bottom (relative order preserved among themselves)."""
    rows = list(sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True))
    if len(rows) < 2:
        return

    def sort_key(row):
        pub_date = _row_date(row, columns)
        if pub_date is not None:
            return (0, -pub_date.toordinal())
        return (1, 0)

    rows.sort(key=sort_key)
    sheet.delete_rows(2, sheet.max_row - 1)
    for row_values in rows:
        sheet.append(row_values)


def autosize(sheet, columns=REPOSITORY_COLUMNS):
    # Keyed by column name (not letter) so it stays correct if the schema changes.
    widths = {
        "Report": 14, "Scraping-Type": 14, "Source": 20, "Newsletter Name": 20,
        "Article Title": 60, "Link": 55, "Full Text": 80,
        "Event Date": 14, "Event Time": 22, "Date Announced": 14,
        "Publication date": 14,
    }
    for name, width in widths.items():
        if name in columns:
            letter = get_column_letter(columns.index(name) + 1)
            sheet.column_dimensions[letter].width = width

    # 2026-09-09, real bug found: `style_header()` sets auto_filter.ref ONCE, the moment a sheet's
    # header is first created (ensure_header() only calls it when the sheet is still empty) - at
    # that instant sheet.dimensions is just the header row itself, so the filter range gets
    # permanently baked in as e.g. "A1:K1" (or, on Repository Rows, a stale "A1:AC489" left over
    # from BEFORE the 29-to-11 column schema trim) and never grows as rows get appended on every
    # later run. Excel then filters/sorts against that frozen, wrong range instead of the sheet's
    # real current data - confirmed as the actual cause of what looked like broken date ordering
    # on Events & Calls. autosize() runs on every save in every append pipeline, so refreshing the
    # filter range here keeps it permanently correct instead of stuck at whatever it was on day one.
    sheet.auto_filter.ref = sheet.dimensions


def finalize_sheet(sheet, columns=REPOSITORY_COLUMNS):
    """Post-processing applied after every append+sort, in every pipeline: column widths/filter
    range (autosize), Full Text wrap/height, and - for EVENTS_COLUMNS - a real date number format
    on Event Date/Date Announced so Excel displays them as dates rather than raw serial numbers.
    Re-applied on every run (not just once) since sort_by_date_desc() rebuilds every row from
    plain values, which resets any per-cell formatting a previous run had set."""
    autosize(sheet, columns)
    text_i = columns.index("Full Text")
    date_col_indices = [columns.index(name) for name in ("Event Date", "Date Announced", "Publication date") if name in columns]
    for row_num in range(2, sheet.max_row + 1):
        row = sheet[row_num]
        row[text_i].alignment = Alignment(wrap_text=True, vertical="top")
        # Fixed single-line height regardless of content length (Full Text varies from ~100
        # chars to 30,000+); wrap_text stays on so the text is still fully in the cell.
        sheet.row_dimensions[row_num].height = DATA_ROW_HEIGHT
        for idx in date_col_indices:
            row[idx].number_format = "yyyy-mm-dd"


def main():
    args = parse_args()
    input_json = Path(args.input_json)
    workbook_path = Path(args.workbook)

    payload = load_payload(input_json)
    all_articles = payload.get("articles", [])
    # 2026-09-09: web-scraped events/calendar pages (CSIS/Atlantic Council/CFR/Chatham
    # House/WEF/Middle East Council) route to "Events & Calls" instead of Repository Rows, same
    # is_event split already used by the vendor-email pipeline (parse_and_append_vendor_emails.py).
    articles = [a for a in all_articles if not a.get("is_event")]
    event_articles = [a for a in all_articles if a.get("is_event")]

    repo_columns = NEW_STRUCTURE_COLUMNS if args.row_schema == "new-structure" else REPOSITORY_COLUMNS
    repo_row_builder = article_to_row_new_structure if args.row_schema == "new-structure" else article_to_row

    workbook = get_or_create_workbook(workbook_path, allow_create=args.allow_create_new_workbook)
    repo_sheet = get_or_create_sheet(workbook, "Repository Rows")
    events_sheet = get_or_create_sheet(workbook, "Events & Calls")
    run_log_sheet = get_or_create_sheet(workbook, "Run Log")
    ensure_header(repo_sheet, repo_columns)
    ensure_header(events_sheet, EVENTS_COLUMNS)
    ensure_header(run_log_sheet, RUN_LOG_COLUMNS)

    added, skipped = append_rows(repo_sheet, articles, repo_columns, repo_row_builder)
    sort_by_date_desc(repo_sheet, repo_columns)
    events_added, events_skipped = append_rows(events_sheet, event_articles, EVENTS_COLUMNS, event_to_row)
    sort_by_date_desc(events_sheet, EVENTS_COLUMNS)

    run_timestamp = dt.datetime.now().isoformat(timespec="seconds")
    run_log_sheet.append([
        run_timestamp,
        args.runner,
        str(input_json),
        len(all_articles),
        len(added),
        skipped,
        json.dumps(payload.get("per_source", {})),
    ])

    finalize_sheet(repo_sheet, repo_columns)
    finalize_sheet(events_sheet, EVENTS_COLUMNS)

    retry_on_lock(lambda: workbook.save(workbook_path), "save workbook")

    counts = collections.Counter(a.get("source", "") for a in added)
    events_counts = collections.Counter(a.get("source", "") for a in events_added)
    print(json.dumps({
        "workbook": str(workbook_path),
        "articles_in_json": len(all_articles),
        "rows_added": len(added),
        "rows_skipped_duplicate": skipped,
        "added_by_source": dict(counts),
        "events_calls_rows_added": len(events_added),
        "events_calls_rows_skipped_duplicate": events_skipped,
        "events_calls_added_by_source": dict(events_counts),
    }, indent=2))


if __name__ == "__main__":
    main()
