"""Check whether the canonical workbook, its review copy, the team copy, and (if present) a
personal mirror all hold the same data.

Written 2026-08-07 after several ad-hoc versions of this same check got run by hand during a
support session (someone asked "did it update?" / "is everything synced?" more than once). All of
those checks did the same thing: compare row-by-row content across the file copies, since a raw
file-size/hash comparison is misleading (opening a workbook in Excel and re-saving it changes its
internal zip/XML bytes even with zero data changes - a real false alarm hit during that session).
This script exists so that check doesn't need to be rewritten from scratch again, by a person or
by Claude - see the "run-update" skill (.claude/skills/run-update/), which calls this directly.

Usage:
    python check_sync.py
    python check_sync.py --shared-library-root "D:\\OneDrive\\Global PP+G - Webscraping & Email Scraping tool"
"""

import argparse
import sys
from pathlib import Path

import append_article_repository_workbook as appender

DEFAULT_CANONICAL = "DO NOT OPEN - Reserved for Automated Updates - Expert Commentary Repository.xlsx"
DEFAULT_REVIEW = "Expert Commentary Repository (review).xlsx"
DEFAULT_TEAM_COPY = "Expert Commentary Repository (team copy).xlsx"


def parse_args():
    ap = argparse.ArgumentParser(description="Compare the canonical workbook against its review/team-copy/mirror copies.")
    # 2026-08-18: SharePoint site restructured (a real move, not a rename) - shortcut is now
    # added directly at the "Webscraping & Email Scraping tool" folder level, so no subfolder
    # underneath it is needed anymore. See the matching comment in run_articles_update.ps1.
    #
    # 2026-08-23: moved again - see the matching comment in run_articles_update.ps1. Old location
    # left untouched as a revert point; drop the " 1" suffix to point back at it.
    #
    # 2026-10-21, per user request: the shared folder now splits into "Old Process & Scope" (the
    # live pipeline) and "New Process & Scope" (reserved for whatever replaces it).
    ap.add_argument("--shared-library-root", default=str(Path.home() / "OneDrive - Deloitte (O365D)" / "Global PP+G - Webscraping & Email Scraping tool 1"))
    ap.add_argument("--shared-subfolder", default="Old Process & Scope")
    ap.add_argument("--canonical-file-name", default=DEFAULT_CANONICAL)
    ap.add_argument("--review-file-name", default=DEFAULT_REVIEW)
    ap.add_argument("--team-copy-file-name", default=DEFAULT_TEAM_COPY)
    ap.add_argument("--local-mirror-path", default=str(Path.home() / "OneDrive - Deloitte (O365D)" / "Documents" / "Expert Commentary Repository" / "Expert Commentary Repository (review).xlsx"))
    return ap.parse_args()


# 2026-09-09: added "Events & Calls" (holds event/webinar/call-invitation rows, kept separate
# from the main research feed) - checked here too now, not just Repository Rows, so a sync issue
# on this sheet specifically doesn't go unnoticed just because the main sheet still matches.
SHEETS_TO_CHECK = ("Repository Rows", "Events & Calls")


def read_sheets(path: Path):
    if not path.exists():
        return None
    wb = appender.get_or_create_workbook(path)
    return {
        sheet_name: [tuple(r) for r in wb[sheet_name].iter_rows(values_only=True)]
        for sheet_name in SHEETS_TO_CHECK
        if sheet_name in wb.sheetnames
    }


def main():
    args = parse_args()
    shared_folder = Path(args.shared_library_root) / args.shared_subfolder

    paths = {
        "canonical": shared_folder / args.canonical_file_name,
        "review": shared_folder / args.review_file_name,
        "team_copy": shared_folder / args.team_copy_file_name,
        "personal_mirror": Path(args.local_mirror_path),
    }

    sheets_by_file = {}
    for name, path in paths.items():
        result = read_sheets(path)
        if result is None:
            print(f"{name}: MISSING - {path}")
            continue
        sheets_by_file[name] = result
        counts = ", ".join(f"{s}: {len(rows) - 1}" for s, rows in result.items())
        print(f"{name}: {counts} - {path}")

    if "canonical" not in sheets_by_file:
        print("\nCan't compare - canonical workbook not found/readable.")
        sys.exit(1)

    base = sheets_by_file["canonical"]
    all_synced = True
    print()
    for name in ("review", "team_copy", "personal_mirror"):
        if name not in sheets_by_file:
            continue
        other = sheets_by_file[name]
        for sheet_name in SHEETS_TO_CHECK:
            if sheet_name not in base or sheet_name not in other:
                continue
            base_rows, other_rows = base[sheet_name], other[sheet_name]
            same_length = len(base_rows) == len(other_rows)
            diffs = [i for i in range(min(len(base_rows), len(other_rows))) if base_rows[i] != other_rows[i]]
            in_sync = same_length and not diffs
            all_synced = all_synced and in_sync
            status = "IN SYNC" if in_sync else "OUT OF SYNC"
            print(f"{name} [{sheet_name}] vs canonical: {status} (same row count: {same_length}, differing rows: {len(diffs)})")
            if diffs[:3]:
                print(f"  first differing row index: {diffs[0]}")

    print("\n" + ("Everything is synced." if all_synced else "NOT everything is synced - see above."))
    sys.exit(0 if all_synced else 1)


if __name__ == "__main__":
    main()
