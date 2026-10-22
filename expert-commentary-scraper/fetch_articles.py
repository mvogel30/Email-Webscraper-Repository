"""Fetch stage: read sources_config.json, scrape every configured source, and write a
per-run JSON file of normalized articles. No topic/keyword filtering - every article found
is kept (these hub pages are already curated expert-commentary feeds).

Usage:
    python fetch_articles.py --output-json run.json
    python fetch_articles.py --output-json run.json --lookback-days 60
    python fetch_articles.py --output-json run.json --all      # no recency window
"""

import argparse
import datetime as dt
import json
import time
from pathlib import Path

import sources


def parse_args():
    ap = argparse.ArgumentParser(description="Scrape expert-commentary articles from all configured sources.")
    ap.add_argument("--config", default=str(Path(__file__).with_name("sources_config.json")))
    ap.add_argument("--output-json", required=True)
    ap.add_argument("--lookback-days", type=int, help="Only keep articles published within N days (overrides config).")
    ap.add_argument("--all", action="store_true", help="Ignore the recency window; keep every article found.")
    return ap.parse_args()


def within_window(published_iso: str, cutoff: dt.date) -> bool:
    """Keep an article if it has no parseable date (never silently drop content)
    or its date is on/after the cutoff."""
    if not published_iso:
        return True
    try:
        return dt.date.fromisoformat(published_iso[:10]) >= cutoff
    except ValueError:
        return True


def main():
    args = parse_args()
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))

    lookback_days = None if args.all else (args.lookback_days if args.lookback_days is not None else config.get("lookback_days"))
    cutoff = (dt.date.today() - dt.timedelta(days=lookback_days)) if lookback_days else None

    all_articles = []
    per_source = {}
    for source in config["sources"]:
        name = source["name"]
        try:
            articles = sources.fetch_source(source)
        except Exception as exc:  # one dead source must not sink the whole run
            # Printed to stdout, not stderr: under PowerShell 5.1 with
            # $ErrorActionPreference = "Stop" (and Start-Transcript active), stderr text from a
            # native process can itself become a script-terminating error even though this
            # process still exits 0.
            print(f"WARNING: source {name!r} failed: {type(exc).__name__}: {exc}", flush=True)
            per_source[name] = {"error": f"{type(exc).__name__}: {exc}", "count": 0}
            continue

        found = len(articles)
        if cutoff is not None:
            articles = [a for a in articles if within_window(a["published"], cutoff)]

        # Always pull the REAL full article body (distinct from the short listing summary),
        # so "Full Text" in the workbook is genuinely the full text.
        for article in articles:
            article["full_text"] = sources.fetch_full_text(article["url"], name)
            time.sleep(0.3)

        per_source[name] = len(articles)
        kept = len(articles)
        print(f"{name}: {kept} kept (of {found} found)", flush=True)
        all_articles.extend(articles)
        time.sleep(sources.REQUEST_PAUSE_SECONDS)

    payload = {
        "fetched_at": dt.datetime.now().isoformat(timespec="seconds"),
        "lookback_days": lookback_days,
        "cutoff": cutoff.isoformat() if cutoff else None,
        "count": len(all_articles),
        "per_source": per_source,
        "articles": all_articles,
    }

    output_path = Path(args.output_json)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")

    print(json.dumps({"total": len(all_articles), "lookback_days": lookback_days, "per_source": per_source, "output": str(output_path)}, indent=2))


if __name__ == "__main__":
    main()
