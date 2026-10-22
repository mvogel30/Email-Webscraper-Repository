"""Append stage for vendor emails (Signum, Eurasia Group, and any future one listed in
vendors.json): reads the combined export JSON, parses each email with its vendor's parser
module, and appends new rows to the SAME shared workbook the web scraper writes to.

Every vendor parser module exposes a uniform parse_email(subject, body, received_date, vendor) ->
list[dict] entry point (see signum_parser.py / eurasia_parser.py), so adding a third vendor
later only means: a new parser module implementing that one function, plus a line in
vendors.json - not a new copy of this script. `vendor` is passed through mainly for
generic_vendor_parser.py, which is shared across multiple vendors with no dedicated module of
their own yet, and uses it as the newsletter-name fallback when no known format is recognized.

Usage:
    python parse_and_append_vendor_emails.py --input-json export.json --workbook repo.xlsx
"""

import argparse
import collections
import datetime as dt
import importlib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))
import append_article_repository_workbook as appender


# 2026-08-05, per user request: a reply/forward is inherently a personal correspondence thread
# (someone at the vendor answering the user's own email), never an original newsletter broadcast
# - true for every vendor, not just one, so this is checked here rather than duplicated in each
# vendor's own parser module. Covers both previously-flagged cases in one general fix: Eurasia's
# "RE: New EG communications" and Signum's "RE: Request for Signum Speaker: ...". Matches one or
# more chained prefixes (e.g. "Re: Fwd:", "RE: RE:").
_REPLY_FORWARD_PREFIX_RE = re.compile(r"^(?:\[EXT\]\s*)?(?:(?:re|fw|fwd)\s*:\s*)+", re.I)


def is_reply_or_forward(subject: str) -> bool:
    return bool(_REPLY_FORWARD_PREFIX_RE.match((subject or "").strip()))


def parse_args():
    ap = argparse.ArgumentParser(description="Append parsed vendor emails to the shared repository workbook.")
    ap.add_argument("--vendors-config", default=str(Path(__file__).with_name("vendors.json")))
    ap.add_argument("--input-json", required=True)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--runner", default="")
    ap.add_argument("--allow-create-new-workbook", action="store_true",
                     help="Allow creating a brand-new empty workbook if none exists yet at "
                          "--workbook. Only pass this for a genuine first-time setup - see "
                          "append_article_repository_workbook.get_or_create_workbook().")
    ap.add_argument("--row-schema", choices=["legacy", "new-structure"], default="legacy",
                     help="Same meaning as append_article_repository_workbook.py's own "
                          "--row-schema - which Repository Rows column layout to write.")
    return ap.parse_args()


def load_vendor_parsers(vendors_config_path: Path) -> dict:
    """{vendor name: parser module}, importing each vendor's module by name from vendors.json."""
    config = json.loads(vendors_config_path.read_text(encoding="utf-8"))
    parsers = {}
    for vendor in config["vendors"]:
        parsers[vendor["name"]] = importlib.import_module(vendor["parser_module"])
    return parsers


def emails_to_articles(emails: list, parsers: dict) -> list:
    articles = []
    skipped_unknown_vendor = collections.Counter()
    skipped_reply_forward = collections.Counter()
    for email in emails:
        vendor = email.get("vendor", "")
        parser = parsers.get(vendor)
        if parser is None:
            skipped_unknown_vendor[vendor] += 1
            continue
        if is_reply_or_forward(email.get("subject", "")):
            skipped_reply_forward[vendor] += 1
            continue
        for parsed in parser.parse_email(
            email["subject"], email["body"], email["received_date"], vendor,
            sender_name=email.get("sender_name", ""), sender_email=email.get("sender_email", ""),
        ):
            full_text = parsed["full_text"]
            articles.append({
                "source": vendor,
                "source_kind": "Email",
                "title": parsed["title"],
                # 2026-10-22, real bug found (reported as "MEC articles look like just headlines"):
                # this was hardcoded to "" for every vendor from back before any parser computed a
                # real per-article link - generic_vendor_parser.parse_mec_daily_brief has returned
                # a real `url` (the wire-story/mecouncil.org link, see its own `teaser_links`
                # system) since 2026-09-24/2026-10-21, but this call site kept discarding it, so
                # the Link column stayed blank for 100% of MEC rows even when the full-text fetch
                # itself succeeded. Most other parsers simply don't set "url" at all, so
                # parsed.get("url", "") is a no-op for them - dedup still falls back to
                # source+title+date wherever a vendor has no url, exactly as before.
                "url": parsed.get("url", ""),
                "published": parsed["published"],
                "summary": full_text[:300],
                "full_text": full_text,
                "geography": parsed["geography"],
                "geography_code": parsed["geography_code"],
                "newsletter_name": parsed.get("newsletter_name", ""),
                # 2026-09-02: event/call invitations (SIGNUM Zoom/Investor Trip/Policy Day/Event,
                # EG Conference Call) are real content again, but routed to a separate "Events &
                # Calls" sheet in main() below rather than mixed into the main research feed - see
                # EVENT_NEWSLETTER_NAMES in signum_parser.py / eurasia_parser.py.
                "is_event": bool(parsed.get("is_event")),
                # 2026-09-09: only meaningful when is_event - the real event date and the date
                # this invitation was received, set by every event-capable parser (signum_parser/
                # eurasia_parser/hub_advisor_events_parser) - see extract_event_date() in
                # email_common.py for why this differs from `published` above.
                "event_date": parsed.get("event_date", ""),
                "announced_date": parsed.get("announced_date", ""),
                # 2026-09-11, per reviewer feedback: dedicated Event Time + a real registration
                # link. "registration_link" is deliberately separate from "url" above (which stays
                # "") - the link is real and clickable but uniquely tracking-wrapped per email
                # send, so it must never affect dedup (see event_to_row() in
                # append_article_repository_workbook.py, which reads it for the "Link" column
                # without going anywhere near article_key()).
                "event_time": parsed.get("event_time", ""),
                "registration_link": parsed.get("registration_link", ""),
            })
    if skipped_unknown_vendor:
        print(f"WARNING: skipped emails from unrecognized vendor tag(s): {dict(skipped_unknown_vendor)}", flush=True)
    if skipped_reply_forward:
        print(f"Skipped reply/forward threads (personal correspondence, not newsletter content): {dict(skipped_reply_forward)}", flush=True)
    return articles


def main():
    args = parse_args()
    input_json = Path(args.input_json)
    workbook_path = Path(args.workbook)
    parsers = load_vendor_parsers(Path(args.vendors_config))

    raw = appender.retry_on_lock(lambda: input_json.read_text(encoding="utf-8-sig"), "read export JSON").strip()
    emails = json.loads(raw) if raw else []
    if isinstance(emails, dict):
        emails = [emails]
    all_articles = emails_to_articles(emails, parsers)
    articles = [a for a in all_articles if not a.get("is_event")]
    event_articles = [a for a in all_articles if a.get("is_event")]

    repo_columns = appender.NEW_STRUCTURE_COLUMNS if args.row_schema == "new-structure" else appender.REPOSITORY_COLUMNS
    repo_row_builder = appender.article_to_row_new_structure if args.row_schema == "new-structure" else appender.article_to_row

    workbook = appender.get_or_create_workbook(workbook_path, allow_create=args.allow_create_new_workbook)
    repo_sheet = appender.get_or_create_sheet(workbook, "Repository Rows")
    events_sheet = appender.get_or_create_sheet(workbook, "Events & Calls")
    run_log_sheet = appender.get_or_create_sheet(workbook, "Run Log")
    appender.ensure_header(repo_sheet, repo_columns)
    appender.ensure_header(events_sheet, appender.EVENTS_COLUMNS)
    appender.ensure_header(run_log_sheet, appender.RUN_LOG_COLUMNS)

    added, skipped = appender.append_rows(repo_sheet, articles, repo_columns, repo_row_builder)
    appender.sort_by_date_desc(repo_sheet, repo_columns)
    # 2026-09-24, per user feedback: CSIS/Atlantic Council event emails re-announce the SAME event
    # under a changing status prefix ("Live Now:", "New Event:", "POSTPONED:", ...) with no stable
    # URL to dedup on, which used to create one row per status change. Collapse those into an
    # in-place update on the matching existing row FIRST - only whatever's left (genuinely new
    # events, or non-hub_advisor events, which this leaves untouched) goes through the normal
    # append path below.
    event_articles, events_updated = appender.upsert_hub_advisor_events(events_sheet, event_articles)
    events_added, events_skipped = appender.append_rows(
        events_sheet, event_articles, appender.EVENTS_COLUMNS, appender.event_to_row
    )
    appender.sort_by_date_desc(events_sheet, appender.EVENTS_COLUMNS)

    counts = collections.Counter(a.get("source", "") for a in added)
    events_counts = collections.Counter(a.get("source", "") for a in events_added)
    run_timestamp = dt.datetime.now().isoformat(timespec="seconds")
    run_log_sheet.append([
        run_timestamp, args.runner, str(input_json),
        len(all_articles), len(added), skipped,
        json.dumps(dict(counts)),
    ])

    appender.finalize_sheet(repo_sheet, repo_columns)
    appender.finalize_sheet(events_sheet, appender.EVENTS_COLUMNS)

    appender.retry_on_lock(lambda: workbook.save(workbook_path), "save workbook")

    print(json.dumps({
        "workbook": str(workbook_path),
        "emails_in_json": len(emails),
        "rows_added": len(added),
        "rows_skipped_duplicate": skipped,
        "added_by_vendor": dict(counts),
        "events_calls_rows_added": len(events_added),
        "events_calls_rows_skipped_duplicate": events_skipped,
        "events_calls_rows_updated_in_place": events_updated,
        "events_calls_added_by_vendor": dict(events_counts),
    }, indent=2))


if __name__ == "__main__":
    main()
