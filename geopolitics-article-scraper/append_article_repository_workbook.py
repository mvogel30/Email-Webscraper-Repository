"""Append stage: read the fetch JSON and append new article rows to the canonical
Excel repository workbook, skipping duplicates.

Reuses the newsletter tool's conventions verbatim where they apply:
  - the same 26-column repository schema (so both workbooks are compatible),
  - retry_on_lock() for OneDrive sync/Excel locks,
  - header creation + styling,
  - a Run Log sheet.

Duplicate identity uses two independent signals, either of which alone marks an article as
already-seen: the URL (normalized identically to how sources.py stores it, so a trailing-slash
or tracking-param variant can never slip past as "new"), and Source+Title (a safety net in case
a site ever changes its URL scheme for an already-captured story).

Usage:
    python append_article_repository_workbook.py --input-json run.json --workbook repo.xlsx
"""

import argparse
import collections
import datetime as dt
import json
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import sources  # for normalize_url() — one shared definition of "same URL" for fetch AND dedup


# Mirrors the newsletter tool's schema, plus a dedicated "Link" column right after the title
# so the article URL is obvious and always present (the inherited "Source Link2" also holds it).
REPOSITORY_COLUMNS = [
    "Source", "Newsletter Name", "Publication Date", "Publication- Date", "Publication Year",
    "Publication Month", "Article Title", "Link", "Industry/Sector", "Geography", "Full Text",
    "Geography 1", "4 Geopolitical Trends", "4 Geopolitical Scenarios", "Topic",
    "Tom Prioritization", "Steve Prioritization", "Total Score", "Match with Feedly",
    "Date", "Title", "Summary2", "Source Name2", "Source Link2", "Full Content",
    "Focus area 1", "Focus area 2",
]

DATA_ROW_HEIGHT = 15  # points — standard single-line height, fixed/uniform for every data row (see main())

RUN_LOG_COLUMNS = [
    "Run Timestamp", "Runner", "Topic", "Source JSON", "Articles In JSON", "Rows Added",
    "Rows Skipped (duplicate)", "Per-Source Counts",
]


def parse_args():
    ap = argparse.ArgumentParser(description="Append scraped article rows to a repository workbook.")
    ap.add_argument("--input-json", required=True)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--runner", default="", help="Machine/person identifier, logged in Run Log. Several teammates may write to the same shared workbook, so this makes it possible to tell whose run added what.")
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


def get_or_create_workbook(path: Path):
    if path.exists():
        return retry_on_lock(lambda: load_workbook(path), "open workbook")
    path.parent.mkdir(parents=True, exist_ok=True)
    return Workbook()


def get_or_create_sheet(workbook, title):
    if title in workbook.sheetnames:
        return workbook[title]
    # Reuse the pristine default sheet openpyxl creates with a new workbook (named "Sheet"),
    # but only that one — never an existing named sheet that is merely still empty.
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
        # Normalized identically to sources.make_article(), so a URL stored years ago (or
        # under a slightly different trailing-slash/tracking-param form) always matches a
        # freshly-fetched one for the same article. This is the fix for a real recurring bug:
        # before this, a trailing-slash variant of an already-seen URL silently created a
        # duplicate row every time the site happened to serve the other form.
        return sources.normalize_url(url).lower()
    pub = pub.isoformat() if isinstance(pub, dt.date) else str(pub or "")[:10]
    return f"{(source or '').strip()}|{(title or '').strip()}|{pub}".lower()


def title_key(source, title) -> str:
    """Secondary duplicate signal: same source + exact same title. A safety net for the rare
    case where a site changes its URL scheme for a story we've already captured — two distinct
    articles from the same source sharing an identical title is effectively never legitimate."""
    return f"title|{(source or '').strip()}|{(title or '').strip()}".lower()


def existing_keys(sheet) -> set:
    """Keys already in the workbook so re-runs don't re-append the same article."""
    link_i = REPOSITORY_COLUMNS.index("Source Link2")
    name_i = REPOSITORY_COLUMNS.index("Newsletter Name")
    title_i = REPOSITORY_COLUMNS.index("Article Title")
    date_i = REPOSITORY_COLUMNS.index("Publication Date")
    keys = set()
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not values or len(values) <= title_i:
            continue
        url = values[link_i] if len(values) > link_i else ""
        keys.add(article_key(url, values[name_i], values[title_i], values[date_i]))
        if values[title_i]:
            keys.add(title_key(values[name_i], values[title_i]))
    return keys


def article_to_row(article: dict) -> dict:
    pub = article.get("published", "")
    pub_date = None
    if pub:
        try:
            pub_date = dt.date.fromisoformat(pub[:10])
        except ValueError:
            pub_date = None
    summary = article.get("summary", "")
    # "Full Text" is the real article body; "Summary2" stays the short blurb. Fall back to the
    # short summary only if fetching the real body genuinely failed (e.g. the article page was
    # unreachable) — a row should never be left completely empty.
    full_text = article.get("full_text", "") or summary
    row = {column: "" for column in REPOSITORY_COLUMNS}
    row.update({
        "Source": "Think Tank",
        "Newsletter Name": article.get("source", ""),
        "Publication Date": pub_date if pub_date else "",
        "Publication- Date": f"{pub_date.day:02d}" if pub_date else "",
        "Publication Year": str(pub_date.year) if pub_date else "",
        "Publication Month": pub_date.strftime("%b") if pub_date else "",
        "Article Title": article.get("title", ""),
        "Link": article.get("url", ""),
        "Full Text": full_text,
        "Geography": article.get("geography", ""),
        "Geography 1": article.get("geography_code", ""),
        "Topic": article.get("topic", ""),
        "Date": pub_date.isoformat() if pub_date else "",
        "Title": article.get("title", ""),
        "Summary2": summary,
        "Source Name2": article.get("source", ""),
        "Source Link2": article.get("url", ""),
    })
    return row


def append_rows(sheet, articles):
    seen = existing_keys(sheet)
    added, skipped = [], 0
    for article in articles:
        url_key = article_key(article.get("url", ""), article.get("source", ""), article.get("title", ""), article.get("published", ""))
        t_key = title_key(article.get("source", ""), article.get("title", ""))
        if url_key in seen or t_key in seen:
            skipped += 1
            continue
        seen.add(url_key)
        seen.add(t_key)
        row = article_to_row(article)
        values = []
        for column in REPOSITORY_COLUMNS:
            value = row.get(column, "")
            if column == "Publication Date" and isinstance(value, dt.date):
                value = dt.datetime(value.year, value.month, value.day)
            values.append(value)
        sheet.append(values)
        added.append(article)
    return added, skipped


def sort_by_date_desc(sheet):
    """Re-order every data row so the newest Publication Date is at the top. Rows with no
    date are kept at the bottom (relative order preserved among themselves) rather than
    guessed into an arbitrary position. Runs every time, not just when rows were added, so
    the sheet is always in this order regardless of how it got here."""
    date_i = REPOSITORY_COLUMNS.index("Publication Date")
    rows = list(sheet.iter_rows(min_row=2, max_row=sheet.max_row, values_only=True))
    if len(rows) < 2:
        return

    def sort_key(row):
        value = row[date_i]
        if isinstance(value, dt.date):  # dt.datetime is a subclass of dt.date
            return (0, -value.toordinal())
        return (1, 0)

    rows.sort(key=sort_key)
    sheet.delete_rows(2, sheet.max_row - 1)
    for row_values in rows:
        sheet.append(row_values)


def autosize(sheet):
    # Keyed by column name (not letter) so it stays correct if the schema changes.
    widths = {
        "Source": 14, "Newsletter Name": 20, "Publication Date": 14, "Article Title": 60,
        "Link": 55, "Full Text": 80, "Topic": 16, "Source Link2": 55,
    }
    for name, width in widths.items():
        if name in REPOSITORY_COLUMNS:
            letter = get_column_letter(REPOSITORY_COLUMNS.index(name) + 1)
            sheet.column_dimensions[letter].width = width


def main():
    args = parse_args()
    input_json = Path(args.input_json)
    workbook_path = Path(args.workbook)

    payload = load_payload(input_json)
    articles = payload.get("articles", [])

    workbook = get_or_create_workbook(workbook_path)
    repo_sheet = get_or_create_sheet(workbook, "Repository Rows")
    run_log_sheet = get_or_create_sheet(workbook, "Run Log")
    ensure_header(repo_sheet, REPOSITORY_COLUMNS)
    ensure_header(run_log_sheet, RUN_LOG_COLUMNS)

    added, skipped = append_rows(repo_sheet, articles)
    sort_by_date_desc(repo_sheet)

    run_timestamp = dt.datetime.now().isoformat(timespec="seconds")
    run_log_sheet.append([
        run_timestamp,
        args.runner,
        payload.get("topic", ""),
        str(input_json),
        len(articles),
        len(added),
        skipped,
        json.dumps(payload.get("per_source", {})),
    ])

    autosize(repo_sheet)
    date_i = REPOSITORY_COLUMNS.index("Publication Date")
    text_i = REPOSITORY_COLUMNS.index("Full Text")
    for row_num in range(2, repo_sheet.max_row + 1):
        row = repo_sheet[row_num]
        row[date_i].number_format = "m/d/yyyy"
        row[text_i].alignment = Alignment(wrap_text=True, vertical="top")
        # Full Text now holds the real article body, so length varies from ~100 chars to
        # 30,000+. Left to Excel's own auto-fit, that made every row a different height. A
        # fixed single-line height instead keeps every row the same regardless of content
        # length; the full text is still entirely in the cell (wrap_text is still on), just
        # read a specific article via the formula bar or expand that one row manually.
        repo_sheet.row_dimensions[row_num].height = DATA_ROW_HEIGHT

    retry_on_lock(lambda: workbook.save(workbook_path), "save workbook")

    counts = collections.Counter(a.get("source", "") for a in added)
    print(json.dumps({
        "workbook": str(workbook_path),
        "articles_in_json": len(articles),
        "rows_added": len(added),
        "rows_skipped_duplicate": skipped,
        "added_by_source": dict(counts),
    }, indent=2))


if __name__ == "__main__":
    main()
