"""Fetch stage: read topics.json, scrape each source for the active topic, and
write a per-run JSON file of normalized articles.

Analogous to export_outlook_newsletters.ps1 in the newsletter tool, but the source
is the web instead of Outlook. The appender reads the JSON this produces.

Usage:
    python fetch_thinktank_articles.py --output-json run.json
    python fetch_thinktank_articles.py --output-json run.json --topic Ukraine
    python fetch_thinktank_articles.py --output-json run.json --lookback-days 60
    python fetch_thinktank_articles.py --output-json run.json --all      # no recency window
"""

import argparse
import datetime as dt
import json
import time
from pathlib import Path

import sources


def parse_args():
    ap = argparse.ArgumentParser(description="Scrape think-tank articles for a configured topic.")
    ap.add_argument("--config", default=str(Path(__file__).with_name("topics.json")))
    ap.add_argument("--output-json", required=True)
    ap.add_argument("--topic", help="Override active_topic from the config.")
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

    topic_name = args.topic or config.get("active_topic")
    if not topic_name or topic_name not in config.get("topics", {}):
        available = ", ".join(config.get("topics", {}).keys())
        raise SystemExit(f"Topic {topic_name!r} not found in config. Available: {available}")

    profile = config["topics"][topic_name]
    keywords = profile.get("keywords", [])
    require_keywords = profile.get("require_keywords", False)

    lookback_days = None if args.all else (args.lookback_days if args.lookback_days is not None else profile.get("lookback_days"))
    cutoff = (dt.date.today() - dt.timedelta(days=lookback_days)) if lookback_days else None

    all_articles = []
    per_source = {}
    for source in profile["sources"]:
        name = source["name"]
        try:
            articles = sources.fetch_source(source, topic_name)
        except Exception as exc:  # one dead source must not sink the whole run
            # Deliberately printed to stdout, not stderr: under PowerShell 5.1 with
            # $ErrorActionPreference = "Stop" (and Start-Transcript active), stderr text
            # from a native process can itself become a script-terminating error even
            # though this process still exits 0 — which would defeat the whole point of
            # catching this exception per-source instead of crashing the entire run.
            print(f"WARNING: source {name!r} failed: {type(exc).__name__}: {exc}", flush=True)
            per_source[name] = {"error": f"{type(exc).__name__}: {exc}", "count": 0}
            continue

        found = len(articles)
        if cutoff is not None:
            articles = [a for a in articles if within_window(a["published"], cutoff)]
        # Per-source override: a source with its own curated topic page (e.g. CFR's Iran page)
        # trusts the site's curation by default; a source with general, unfiltered content
        # (e.g. a firm's whole "Insights" feed) sets "require_keywords": true in topics.json
        # so only Iran-related hits are kept from it.
        source_requires_keywords = source.get("require_keywords", require_keywords)
        if source_requires_keywords:
            articles = [a for a in articles if sources.matches_keywords(a, keywords)]

        # Optional: enrich articles that lack a short summary by visiting the article page.
        if source.get("fetch_summary"):
            for article in articles:
                if not article["summary"]:
                    article["summary"] = sources.fetch_article_summary(article["url"])
                    time.sleep(0.3)

        # Always: pull the REAL full article body (distinct from the short summary above),
        # so "Full Text" in the workbook is genuinely the full text, not a couple of sentences.
        for article in articles:
            article["full_text"] = sources.fetch_full_text(article["url"], name)
            time.sleep(0.3)

        per_source[name] = len(articles)
        kept = len(articles)
        print(f"{name}: {kept} kept (of {found} found)", flush=True)
        all_articles.extend(articles)
        time.sleep(sources.REQUEST_PAUSE_SECONDS)

    payload = {
        "topic": topic_name,
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

    print(json.dumps({"topic": topic_name, "total": len(all_articles), "lookback_days": lookback_days, "per_source": per_source, "output": str(output_path)}, indent=2))


if __name__ == "__main__":
    main()
