"""Fetch and normalize expert-commentary articles from a fixed set of hub pages.

Sibling tool to geopolitics-article-scraper/, but unfiltered: every article found on each
configured source is kept (no topic/keyword filter), since these hub pages are already curated
expert-commentary feeds rather than general news firehoses.

Mirrors the same two-stage shape: a thin *fetch* stage produces raw normalized article records
(written to JSON by fetch_articles.py), and a separate *append* stage de-duplicates them into the
canonical workbook (append_article_repository_workbook.py).

Each source in sources_config.json names a `parser`:
  - "rss"              : any RSS/Atom feed (Lowy Institute)
  - "cfr"               : Council on Foreign Relations "Expert Takes"
  - "csis"              : CSIS "Analysis" listing
  - "chatham"           : Chatham House "Expert comment" search results (Playwright-rendered -
                          blocks plain HTTP clients with an intermittent 403)
  - "chinamacro"        : China Macro Group, across all 4 of their topic pages (not just
                          "Insights" - see extract_chinamacro), with real per-post JSON-LD dates
  - "ac_dispatches"     : Atlantic Council "Dispatches" listing
  - "deloitte_insights" : Deloitte Insights (sitemap-based, not a hub page - see below)
  - "wef_publications"  : World Economic Forum "Publications" listing (Playwright-rendered)
  - "generic"           : best-effort anchor scrape (fallback for a new site)

A normalized article is a plain dict:
  {source, title, url, published (ISO date str or ""), published_raw, summary}
"""

import datetime as dt
import email.utils
import html
import io
import json
import re
import sys
import time
from zoneinfo import ZoneInfo
from pathlib import Path
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

# 2026-09-11, per reviewer feedback: web-scraped events had no Geography column populated at all
# (this tool's web pipeline never had a geography-inference step - unlike the email pipeline,
# whose own copy lives in email-scraper/email_common.py). Reused directly rather than duplicated,
# since COUNTRY_ALIASES is a real, already-tuned list - not re-implemented here.
sys.path.insert(0, str(Path(__file__).resolve().parent / "email-scraper"))
from email_common import infer_geography  # noqa: E402

import requests
from bs4 import BeautifulSoup

try:
    import feedparser
except ImportError:  # surfaced with a clear message only if an RSS source is used
    feedparser = None

try:
    import pypdf
except ImportError:  # surfaced with a clear message only if WEF's PDF full-text is used
    pypdf = None


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 ExpertCommentaryScraper/1.0"
)
REQUEST_TIMEOUT = 40
REQUEST_PAUSE_SECONDS = 1.0  # polite delay between sources within a run


# --------------------------------------------------------------------------- #
# HTTP
# --------------------------------------------------------------------------- #
def fetch_url(url: str, attempts: int = 3, backoff_seconds: float = 2.0) -> str:
    """GET a URL as text. Honors HTTP(S)_PROXY env vars via requests defaults.

    Retries a couple of times on transient failures (timeouts, connection resets, 5xx)
    since those often clear up on their own within a few seconds. Deliberately does NOT
    retry 4xx responses (403, 404, ...) — retrying wouldn't help, and hammering a site
    that's actively blocking us (e.g. rate-limiting) would only make things worse.
    """
    headers = {"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.9"}
    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            if response.status_code >= 500 and attempt < attempts:
                time.sleep(backoff_seconds * attempt)
                continue
            response.raise_for_status()
            return response.text
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as exc:
            last_exc = exc
            if attempt < attempts:
                time.sleep(backoff_seconds * attempt)
                continue
            raise
    raise last_exc


def make_soup(html: str) -> BeautifulSoup:
    """Prefer lxml (faster/leaner); fall back to the stdlib parser if absent."""
    try:
        return BeautifulSoup(html, "lxml")
    except Exception:
        return BeautifulSoup(html, "html.parser")


# --------------------------------------------------------------------------- #
# Date parsing
# --------------------------------------------------------------------------- #
_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
MONTHS = {}
for _i, _name in enumerate(_MONTH_NAMES, start=1):
    MONTHS[_name.lower()] = _i
    MONTHS[_name[:3].lower()] = _i

_DATE_TEXT_RE = re.compile(r"[A-Z][a-z]+\.?\s+\d{1,2},\s+\d{4}|\d{1,2}\s+[A-Z][a-z]+\.?\s+\d{4}")


def parse_date(value: str):
    """Parse an ISO 8601, 'July 9, 2026', or '21 July 2026' string to a date."""
    if not value:
        return None
    value = value.strip()

    iso = re.match(r"(\d{4})-(\d{2})-(\d{2})", value)
    if iso:
        try:
            return dt.date(int(iso.group(1)), int(iso.group(2)), int(iso.group(3)))
        except ValueError:
            pass

    # "July 9, 2026"
    m = re.search(r"([A-Za-z]+)\.?\s+(\d{1,2}),\s+(\d{4})", value)
    if m and m.group(1).lower() in MONTHS:
        try:
            return dt.date(int(m.group(3)), MONTHS[m.group(1).lower()], int(m.group(2)))
        except ValueError:
            pass

    # "21 July 2026"
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\.?\s+(\d{4})", value)
    if m and m.group(2).lower() in MONTHS:
        try:
            return dt.date(int(m.group(3)), MONTHS[m.group(2).lower()], int(m.group(1)))
        except ValueError:
            pass

    return None


def date_from_text(text: str):
    """Find the first date-like substring in a blob of text and parse it."""
    match = _DATE_TEXT_RE.search(text or "")
    return parse_date(match.group(0)) if match else None


def preferred_time_tag_date(time_tag):
    """2026-09-23, per QA feedback: dates should always be the date the SOURCE itself uses, not
    ours after any conversion. Real bug found across several sources - a `<time datetime="...">`
    attribute is very often server-generated UTC (confirmed live: CFR, Chatham House, Atlantic
    Council, CFR Events all use a "...Z" UTC timestamp), which can land on a different calendar
    day than what the page's own VISIBLE text says once you account for the publisher's real
    timezone (confirmed live on WEF - one real entry had datetime="2026-09-14T22:01:00Z" [UTC]
    while the page's own visible text plainly read "Sep 15, 2026", since 22:01 UTC is already past
    midnight in WEF's home timezone CEST). The visible text is always already in the publisher's
    own intended local day - no ambiguity, no conversion needed - so it's preferred outright over
    the attribute; the attribute is only a fallback for the rare case a tag has no visible text at
    all (matches the OLD behavior for that one case, so nothing regresses)."""
    text = time_tag.get_text(strip=True) if time_tag is not None else ""
    if text:
        return text
    return _get_ci(time_tag, "datetime") if time_tag is not None else None


def local_date_from_iso(value: str, tz_name: str):
    """For a source with NO visible date text anywhere on the page (Deloitte Insights, China
    Macro Group - both confirmed live: their real article pages render the byline entirely
    client-side or not at all, so there's no human-readable fallback to prefer) - the only date
    available is a machine JSON-LD `datePublished`, always UTC ("...Z" or "+00:00"). Reading that
    literally is wrong whenever the real local publish time crosses a UTC-day boundary - confirmed
    live and real, not theoretical: a Deloitte article at "2026-09-16T03:00:00Z" is 11pm the
    PREVIOUS evening in its real home timezone (America/New_York), and a systematic check of every
    already-scraped Deloitte/China Macro row found 8/35 and 11/80 real mismatches respectively,
    every one of them exactly this off-by-one-day shift. Converts properly via zoneinfo (handles
    DST automatically, unlike a fixed UTC offset) to recover the calendar date the publisher's own
    clock would have shown."""
    if not value:
        return None
    try:
        parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        return parsed.date()
    return parsed.astimezone(ZoneInfo(tz_name)).date()


# --------------------------------------------------------------------------- #
# Normalization
# --------------------------------------------------------------------------- #
def compact(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


_TRACKING_PARAMS = {
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "fbclid", "gclid", "mc_cid", "mc_eid",
}


_PERCENT_ENCODED_RE = re.compile(r"%[0-9A-Fa-f]{2}")
# Matches the literal "." form and its percent-encoded form - this runs AFTER _PERCENT_ENCODED_RE
# has already uppercased any encoding, so "%2E" (never "%2e") is the only encoded form to expect.
_INDEX_PHP_PREFIX_RE = re.compile(r"^/index(?:\.|%2E)php(?=/)", re.I)


def normalize_url(url: str) -> str:
    """Canonicalize a URL for stable de-duplication: lowercase scheme/host, drop a
    trailing slash, drop known tracking query params.

    2026-09-24, per a broad quality sweep: found real CSIS events duplicated under 3-4 URL
    variants of the exact same page - "/events/<slug>" and "/index.php/events/<slug>" both serve
    the identical content (a Drupal quirk, not CSIS-specific), and the site's own listing pages
    inconsistently percent-encode the literal "." in "index.php" as "%2E" vs "%2e" depending on
    which page rendered the link. Neither was being canonicalized, so the same real event kept
    reappearing as a "new" one. Both fixes are conservative: uppercasing percent-encoded hex
    digits is a pure RFC 3986 no-op for any URL that was already consistent, and stripping an
    "/index.php" prefix only matches that exact literal segment, never arbitrary path content."""
    if not url:
        return url
    parts = urlsplit(url.strip())
    path = _PERCENT_ENCODED_RE.sub(lambda m: m.group(0).upper(), parts.path)
    path = _INDEX_PHP_PREFIX_RE.sub("", path)
    path = path.rstrip("/") or "/"
    query_pairs = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True) if k.lower() not in _TRACKING_PARAMS]
    return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), path, urlencode(query_pairs), ""))


def make_article(source, title, url, published, summary):
    return {
        "source": source,
        "title": compact(title),
        "url": normalize_url(url),
        "published": published.isoformat() if isinstance(published, dt.date) else "",
        "published_raw": published.strftime("%b %d, %Y") if isinstance(published, dt.date) else "",
        "summary": compact(summary),
    }


def _get_ci(tag, name: str):
    """Case-insensitive attribute lookup (guards against dateTime vs datetime)."""
    if tag is None:
        return None
    for key, value in tag.attrs.items():
        if key.lower() == name.lower():
            return value
    return None


# --------------------------------------------------------------------------- #
# Per-source extractors
# --------------------------------------------------------------------------- #
def _rss_entry_date(entry):
    """2026-09-23, per QA feedback: dates were consistently one day off for Lowy Institute (and
    presumably any other RSS source). Root cause, confirmed against the real live feed - Lowy
    publishes at "00:00:00 +1000" (Sydney midnight); feedparser's own `published_parsed`/
    `updated_parsed` ALWAYS normalizes to UTC internally, and UTC is behind +1000, so every single
    entry's date rolls back to the PREVIOUS calendar day (verified: 15/15 real entries checked
    were off by exactly one day, not an occasional edge case). Fixed by parsing the RAW date
    string feedparser also exposes (`entry.published`/`entry.updated`, e.g. "Wed, 23 Sep 2026
    00:00:00 +1000") ourselves via `email.utils.parsedate_to_datetime` and taking its `.date()`
    directly - that keeps the calendar date exactly as the source itself printed it, with no UTC
    conversion at all, which is what "the date the source uses for their publication" actually
    means. Falls back to feedparser's own UTC-shifted struct only if the raw string is ever
    missing or unparseable (better than nothing, matches the old behavior in that rare case)."""
    for attr in ("published", "updated"):
        raw = entry.get(attr)
        if raw:
            try:
                return email.utils.parsedate_to_datetime(raw).date()
            except (TypeError, ValueError):
                pass
    for attr in ("published_parsed", "updated_parsed"):
        parsed = getattr(entry, attr, None)
        if parsed:
            return dt.date(parsed.tm_year, parsed.tm_mon, parsed.tm_mday)
    return None


def extract_rss(url, source_name):
    if feedparser is None:
        raise RuntimeError("feedparser is not installed (pip install feedparser) but source uses parser 'rss'")
    feed = feedparser.parse(fetch_url(url))
    articles = []
    for entry in feed.entries:
        published = _rss_entry_date(entry)
        summary = entry.get("summary", "") or ""
        summary = make_soup(summary).get_text(" ", strip=True)
        articles.append(make_article(source_name, entry.get("title", ""), entry.get("link", ""), published, summary))
    return articles


# CFR mixes several content types in identical h3 cards — match any known content type, not
# just "/articles/" (a real report was silently missed once by matching only "/articles/").
_CFR_CONTENT_PATH_PREFIXES = ("/articles/", "/reports/", "/report/", "/event/")


def extract_cfr(html, base_url, source_name):
    soup = make_soup(html)
    articles, seen = [], set()
    for anchor in soup.select("h3 a[href]"):
        href = anchor.get("href", "")
        if not any(prefix in href for prefix in _CFR_CONTENT_PATH_PREFIXES):
            continue
        url = urljoin(base_url, href)
        if url in seen:
            continue
        seen.add(url)
        h3 = anchor.find_parent("h3")
        card = h3.parent if h3 else anchor.parent
        time_tag = (card.find("time") if card else None) or anchor.find_next("time")
        published = parse_date(preferred_time_tag_date(time_tag)) if time_tag is not None else None
        articles.append(make_article(source_name, anchor.get_text(" ", strip=True), url, published, ""))
    return articles


def extract_cfr_reports(html, base_url, source_name):
    """CFR's dedicated /reports listing (cfr.org/reports) - reuses extract_cfr()'s own selector
    and date logic unchanged (it already recognizes "/reports/" as a valid content path via
    _CFR_CONTENT_PATH_PREFIXES), but every article's Source stays plain "CFR" (same cross-
    channel-dedup reasoning the Events extractors elsewhere already use - so a report that also
    happens to get linked from the regular /expert-takes listing dedupes as one row, not two) and
    Newsletter Name is explicitly "Reports", not the "Expert Takes" NEWSLETTER_NAME_BY_SOURCE
    default a regular CFR article would get.

    2026-09-23, per user request, after a real gap was found: a CFR "Expert Takes" summary
    article ("The U.S. Is Highly Vulnerable to Cyber Threats From China...") teases a full
    companion report ("The Cyber Gap") that was completely invisible to this tool - neither
    captured as its own row nor merged into the summary. A live check of 15 real Expert Takes
    articles found this is genuinely common, not a one-off: 5/15 (about a third) link out to a
    separate report; the rest are standalone original commentary with nothing missing.

    Deliberately kept as fully SEPARATE rows from whatever summary article references a report,
    not merged into it - a summary can reference multiple reports, or cite THIRD-PARTY reports
    that aren't CFR's own at all (confirmed real on this exact article: it also linked out to a
    DNI assessment, a CrowdStrike report, and a NIST publication - none of which should ever be
    pulled in), and a report can also be published with no summary article ever linking to it."""
    articles = extract_cfr(html, base_url, "CFR")
    for article in articles:
        article["newsletter_name"] = "Reports"
    return articles


def extract_csis(html, base_url, source_name):
    soup = make_soup(html)
    articles, seen = [], set()
    for art in soup.find_all("article"):
        anchor = art.select_one('a.hocus-headline[href*="/analysis/"]') or art.select_one('a[href*="/analysis/"]')
        if anchor is None:
            continue
        url = urljoin(base_url, anchor.get("href", ""))
        if url in seen:
            continue
        seen.add(url)
        teaser = art.select_one(".teaser")
        summary = teaser.get_text(" ", strip=True) if teaser else ""
        published = date_from_text(art.get_text(" ", strip=True))
        articles.append(make_article(source_name, anchor.get_text(" ", strip=True), url, published, summary))
    return articles


def extract_chatham(html, base_url, source_name):
    soup = make_soup(html)
    articles, seen = [], set()
    for art in soup.select("article.teaser"):
        anchor = art.select_one("a.teaser__link[href]") or art.select_one("a[href]")
        if anchor is None:
            continue
        href = anchor.get("href", "")
        if not re.search(r"/20\d{2}/|/publications/", href):  # articles & publications, not events
            continue
        url = urljoin(base_url, href)
        if url in seen:
            continue
        seen.add(url)
        title_el = art.select_one(".teaser__title")
        title = title_el.get_text(" ", strip=True) if title_el else anchor.get_text(" ", strip=True)
        summary_el = art.select_one(".teaser__summary")
        summary = summary_el.get_text(" ", strip=True) if summary_el else ""
        time_tag = art.select_one("time")
        published = parse_date(preferred_time_tag_date(time_tag)) if time_tag is not None else None
        articles.append(make_article(source_name, title, url, published, summary))
    return articles


# Chatham House blocks plain HTTP clients (requests/curl get an intermittent 403 - confirmed
# 2026-08-14 via the Run Log alternating success/403 across consecutive days) on BOTH the search
# listing AND individual article pages (verified separately - a plain fetch_url() to a real
# article URL 403s the same way). This is a WAF/bot-detection block, not a real IP/UA ban: a real
# headless Chromium session (same technique already used for WEF) gets a clean 200 every time,
# with identical markup to what a successful plain request would have returned - confirmed
# directly against both the listing (extract_chatham() parses it unchanged) and an article page
# (the existing ".body-content" selector in FULL_TEXT_SELECTORS still applies unchanged too, see
# fetch_full_text() below). Used for both, rather than retrying the blocked plain request, which
# only ever had a chance of working.
def _fetch_html_via_browser(url: str, wait_ms: int = 1500) -> str:
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        page.wait_for_timeout(wait_ms)
        html = page.content()
        browser.close()
    return html


def extract_chatham_via_browser(url, source_name):
    return extract_chatham(_fetch_html_via_browser(url), url, source_name)


# China Macro Group publishes across several separate topic pages, not just the "Insights" page
# this source was originally pointed at - found 2026-08-19: /macro-policy alone has ~3x more
# posts than /insights, with almost no overlap (45 unique posts across all 4 pages combined, 33
# of them never appearing on /insights at all). Every one of these pages is visited for candidate
# URLs now, not just the configured one.
_CHINAMACRO_LISTING_PATHS = ("insights", "macro-policy", "sector-trends", "geopolitics")

# 2026-08-19: visiting every candidate's own page for its JSON-LD date (see extract_chinamacro)
# takes ~3-4 minutes across ~80 posts, EVERY run, since there's no cheap "lastmod"-style signal
# to prefilter by the way Deloitte Insights' sitemap has. A post's real datePublished can't
# change once fetched, so it's cached here by URL - a run only pays the individual-page-visit
# cost for candidates it has genuinely never resolved before (new posts, normally 0-2 per run).
# The listing pages themselves are still ALWAYS re-scanned in full on every run (cheap, 4 fetches)
# so a newly-published post is guaranteed to be seen and added to the cache - caching only skips
# re-fetching pages already known, never skips discovering new ones.
_CHINAMACRO_DATE_CACHE_PATH = Path(__file__).with_name("chinamacro_date_cache.json")


def _load_chinamacro_date_cache() -> dict:
    try:
        return json.loads(_CHINAMACRO_DATE_CACHE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_chinamacro_date_cache(cache: dict) -> None:
    try:
        _CHINAMACRO_DATE_CACHE_PATH.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass  # best-effort - a cache write failure should never break a real run


def extract_chinamacro(base_url, source_name):
    """China Macro Group listings (Wix), across all known topic pages - self-fetching (see
    _SELF_FETCHING_PARSERS), not given pre-fetched HTML, since it needs to visit several pages.

    Dates are NOT derived from the listing anymore. The old approach parsed a "Month Year | ..."
    prefix out of the listing teaser text, but that prefix is inconsistent even on /insights and
    absent entirely on the other three pages (confirmed directly - no such text exists in their
    card markup at all), and unparsed dates were treated as "keep forever, never expires" rather
    than dropped. That silently let real, old content get stuck permanently: one already-captured
    row that showed as blank/undated in the workbook turned out to have a real datePublished of
    2025-03-19 on its own page - over a year old, and never correctly aged out because it looked
    undated to the old logic. Every candidate's own page is now visited instead, for its real
    schema.org/BlogPosting JSON-LD datePublished - the same approach already used for Deloitte
    Insights, confirmed present and reliable on every real post checked."""
    site_root = f"{urlsplit(base_url).scheme}://{urlsplit(base_url).netloc}"

    candidates = {}
    for path in _CHINAMACRO_LISTING_PATHS:
        listing_url = f"{site_root}/{path}"
        try:
            listing_html = fetch_url(listing_url)
        except Exception:
            continue
        listing_soup = make_soup(listing_html)
        for card in listing_soup.select('[data-hook="post-list-item"]'):
            anchor = card.select_one('a[href*="/post/"]')
            if anchor is None:
                continue
            url = urljoin(listing_url, anchor.get("href", ""))
            if url in candidates:
                continue
            title_el = card.select_one('[data-hook="post-title"]') or anchor
            candidates[url] = title_el.get_text(" ", strip=True) if title_el else ""

    cache = _load_chinamacro_date_cache()
    cache_changed = False
    articles = []
    for url, listing_title in candidates.items():
        cached = cache.get(url)
        if cached is not None:
            title = cached.get("title") or listing_title
            published = dt.date.fromisoformat(cached["published"]) if cached.get("published") else None
            if title:
                articles.append(make_article(source_name, title, url, published, ""))
            continue

        try:
            post_html = fetch_url(url)
        except Exception:
            continue
        post_soup = make_soup(post_html)
        title, published = listing_title, None
        for script in post_soup.find_all("script", {"type": "application/ld+json"}):
            try:
                data = _json.loads(script.string or "")
            except Exception:
                continue
            for item in (data if isinstance(data, list) else [data]):
                if isinstance(item, dict) and item.get("@type") in ("BlogPosting", "Article", "NewsArticle"):
                    title = html.unescape(item.get("headline") or "") or title
                    # China Macro's own pages render no visible byline date at all (confirmed
                    # live) - JSON-LD datePublished is the only date available, always UTC, so it
                    # must be converted to China Macro's real home timezone (see
                    # local_date_from_iso's docstring - a systematic check found 11/80 already-
                    # scraped rows off by exactly one day from this exact conversion gap).
                    published = local_date_from_iso(item.get("datePublished", ""), "Asia/Shanghai") or published
        if not title:
            continue
        articles.append(make_article(source_name, title, url, published, ""))
        cache[url] = {"title": title, "published": published.isoformat() if published else ""}
        cache_changed = True
        time.sleep(0.3)

    if cache_changed:
        _save_chinamacro_date_cache(cache)
    return articles


def extract_ac_dispatches(html, base_url, source_name):
    """Atlantic Council "Dispatches" listing — a different template from the RSS-based
    region/topic feeds the sibling tool uses (this one has real per-article h2/time markup)."""
    soup = make_soup(html)
    articles, seen = [], set()
    for art in soup.select("article.article-item, article.featured-article"):
        anchor = art.select_one(".article-content h2 a[href]") or art.select_one("h2 a[href], h3 a[href]")
        if anchor is None:
            continue
        url = urljoin(base_url, anchor.get("href", ""))
        if url in seen:
            continue
        seen.add(url)
        summary_el = art.select_one(".article-content p") or art.select_one("p")
        summary = summary_el.get_text(" ", strip=True) if summary_el else ""
        time_tag = art.select_one("time.article-date") or art.select_one("time")
        published = parse_date(preferred_time_tag_date(time_tag)) if time_tag is not None else None
        articles.append(make_article(source_name, anchor.get_text(" ", strip=True), url, published, summary))
    return articles


import json as _json

# Deloitte Insights has no dated listing page (the real hub page's article feed is
# client-side-rendered, and every static "topic hub" URL is an evergreen collection with no
# dates at all) - but every individual article DOES carry a real datePublished, just via
# schema.org JSON-LD, not visible listing markup. So instead of a hub page, this source is
# pointed at Deloitte's public sitemap, which has a real <lastmod> per URL; candidates are
# narrowed by path + lastmod first (avoids visiting all ~2,000 Insights URLs every run), then
# each candidate is visited once to pull its real title/date from JSON-LD.
_DELOITTE_INSIGHTS_PATH = "/us/en/insights/"
# NOT the tool's real lookback window (fetch_articles.py's own cutoff, applied afterward using
# each article's real JSON-LD published date, is what actually decides what's kept) - this is
# just how far back to bother visiting candidate pages at all. Deloitte publishes across every
# topic constantly, so even 30-45 days here means fetching hundreds of individual pages every
# run - a couple of days' buffer past the daily run cadence (same idea as DaysBack on the vendor
# email export scripts) is enough to never miss anything, without that blowing up run time.
_DELOITTE_SITEMAP_PREFILTER_DAYS = 5


def extract_deloitte_insights(sitemap_url, source_name):
    xml = fetch_url(sitemap_url)
    cutoff = dt.date.today() - dt.timedelta(days=_DELOITTE_SITEMAP_PREFILTER_DAYS)
    candidates = []
    for loc, lastmod in re.findall(r"<url>\s*<loc>([^<]+)</loc>\s*<lastmod>([^<]+)</lastmod>\s*</url>", xml):
        if _DELOITTE_INSIGHTS_PATH not in loc:
            continue
        lastmod_date = parse_date(lastmod)
        if lastmod_date and lastmod_date < cutoff:
            continue
        candidates.append((loc, lastmod_date))

    articles = []
    for url, lastmod_date in candidates:
        try:
            soup = make_soup(fetch_url(url))
        except Exception:
            continue
        title, published, is_article = "", lastmod_date, False
        for script in soup.find_all("script", {"type": "application/ld+json"}):
            try:
                data = _json.loads(script.string or "")
            except Exception:
                continue
            for item in (data if isinstance(data, list) else [data]):
                if isinstance(item, dict) and item.get("@type") in ("Article", "NewsArticle", "BlogPosting"):
                    is_article = True
                    title = html.unescape(item.get("headline") or "") or title
                    # 2026-10-01, per QA feedback: the America/New_York conversion below (added
                    # for an earlier "8/35 off by one" finding) was itself re-checked against the
                    # sitemap's own <lastmod> (an independent, Deloitte-generated signal the
                    # conversion never touches) - confirmed live on the QA-flagged example
                    # (coo-agentic-ai-transformation.html, datePublished "2026-09-16T03:00:00Z"):
                    # lastmod reads "2026-09-16T11:24:26Z", i.e. the SAME UTC calendar day as
                    # datePublished, not the day earlier the timezone conversion produces. A
                    # broader sample (12 random current articles) found the conversion only ever
                    # disagrees with the raw UTC date for this same narrow pattern (a datePublished
                    # in the first few hours of the UTC day), and always in the wrong direction.
                    # Reading the UTC calendar date directly (no conversion) matches both lastmod
                    # and what a reviewer sees literally reading the timestamp's date portion.
                    published = parse_date(item.get("datePublished", "")) or published
        if not is_article:
            # Not actually an article - a topic/tracker/hub landing page under the same
            # /insights/ path (e.g. "ConsumerSignals"), which has a title but no real body and
            # a <lastmod> that just reflects a routine data refresh, not new commentary. Skip it
            # rather than accept a title-only stub.
            continue
        if not title:
            og = soup.select_one('meta[property="og:title"]')
            title = (og.get("content") if og else "") or (soup.title.get_text(strip=True) if soup.title else "")
        if not title:
            continue
        articles.append(make_article(source_name, title, url, published, ""))
        time.sleep(0.3)
    return articles


# WEF actively blocks plain HTTP clients: requests AND curl (different TLS stacks entirely -
# rules out a library-specific fix) both get an identical Akamai "Access Denied" on every path,
# including /robots.txt itself, which fires before any page-specific logic runs. A real browser
# on this same machine/network loads the page fine, dates and all, which is what a real headless
# Chromium (Playwright) reproduces here - confirmed directly (status 200, real content) before
# committing to this dependency. This is the only source in the tool that needs it.
def extract_wef_publications(url, source_name):
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)  # lets the client-rendered publication cards finish painting
        html = page.content()
        browser.close()

    soup = make_soup(html)
    articles, seen = [], set()
    for art in soup.find_all("article"):
        # WEF's own CSS classes are build-hashed (Chakra UI/Emotion, e.g. "wef-1qbpo0l") and
        # change across deploys - the `lineclamp` attribute (a literal HTML attribute, not a
        # class) is the stable hook: the title paragraph carries lineclamp="4", the description
        # div right after it carries lineclamp="3".
        title_p = art.select_one("p[lineclamp]")
        anchor = title_p.select_one('a[href*="/publications/"]') if title_p else None
        if anchor is None:
            continue
        link = urljoin(url, anchor.get("href", ""))
        if link in seen:
            continue
        seen.add(link)
        time_tag = art.select_one("time[datetime]")
        # Unlike every other source, a missing date here does NOT mean "kept anyway, date
        # unknown" - the page mixes in undated "Series" tiles (e.g. "Global Risks Report" as a
        # recurring series name, not a specific dated report) that would otherwise slip past the
        # recency window forever (blank dates always pass it) as permanently empty rows. WEF's
        # own JSON confirmed these truly have no date anywhere, so skipping them here is correct,
        # not a case of discarding real content with an unparsed date.
        if time_tag is None:
            continue
        published = parse_date(preferred_time_tag_date(time_tag))
        summary_el = art.select_one("div[lineclamp]")
        summary = summary_el.get_text(" ", strip=True) if summary_el else ""
        articles.append(make_article(source_name, anchor.get_text(" ", strip=True), link, published, summary))
    return articles


# --------------------------------------------------------------------------- #
# Events extractors (2026-09-09, per user request) - each org's regular commentary is already
# scraped above; these hit each org's own SEPARATE events/calendar page and feed the workbook's
# "Events & Calls" sheet instead of "Repository Rows" - see the `is_event` flag set on every
# article below, checked in append_article_repository_workbook.py's main(). One-time manual
# capture (2026-09-09, back to 2026-08-27) was verified against these exact pages/structures
# before wiring this in as permanent daily automation.
# --------------------------------------------------------------------------- #

_WEF_DATE_RANGE_RE = re.compile(r"(\d{1,2})\s*[–—-]\s*\d{1,2}\s+([A-Za-z]+)\s+(\d{4})")

# 2026-09-11, per reviewer feedback: a real event (CSIS's "Nigeria 2027: Can Nigerians Trust the
# 2027 Election?", Sept 10) was confirmed missing from the workbook - root-caused directly: CSIS
# renders TWO differently-sized card variants on the SAME single page load, `.ts-card-event-sm`
# (a small ~4-item "featured" strip) and `.ts-card-event-md` (the real, much fuller listing - 12
# real events found in one check, covering both upcoming and recently-past). The old code only
# ever selected `.ts-card-event-sm`, silently missing everything only present as `.ts-card-event-
# md`. Fixed by selecting both. This also makes the separate `?past_events=1` fetch pointless (it
# was already unreliable - see the 403 handling that used to live here) since a single page load
# now covers both eras via the `-md` cards alone; removed.
_CSIS_CARD_SELECTOR = ".ts-card-event-sm, .ts-card-event-md"

# Same explicit-timezone-whitelist approach as email_common.py's own event-time regex (see that
# module's docstring for why a bare-letters version once falsely matched "21-24 Septe" out of a
# WEF date RANGE as if it were a time) - duplicated here rather than imported since sources.py and
# email_common.py are otherwise independent modules with their own separate regexes already.
_WEB_TZ_ABBREV = r"(?:ET|EDT|EST|PT|PDT|PST|CT|CDT|CST|MT|MDT|MST|GMT|BST|CET|CEST|UTC|SGT|GST|IST|AEST|AEDT|HKT|JST|KST)"
# 2026-09-23, per QA feedback: CFR's individual event pages write out a real time as "4:15 p.m.
# (EDT)" - periods in "p.m." and a parenthesized zone, neither of which the plain "am|pm" /
# bare-abbreviation shape below matched, silently leaving CFR's Event Time blank. Also fixes the
# same "p.m." gap independently found in a real Atlantic Council event email's body ("7:00 p.m.
# ET"). _WEB_TZ lets the abbreviation be optionally wrapped in parens; _WEB_AMPM accepts an
# optional period after each letter.
_WEB_AMPM = r"(?:[ap]\.?m\.?)"
_WEB_TZ = rf"\(?{_WEB_TZ_ABBREV}\)?"
_WEB_TIME_SEP = r"(?:-|–|to)"
_WEB_EVENT_TIME_RE = re.compile(
    rf"\d{{1,2}}(?::\d{{2}})?\s*{_WEB_AMPM}?\s*{_WEB_TIME_SEP}\s*\d{{1,2}}(?::\d{{2}})?\s*{_WEB_AMPM}?\s*{_WEB_TZ}"
    rf"(?:\s*[/|]\s*\d{{1,2}}(?::\d{{2}})?\s*{_WEB_AMPM}?\s*{_WEB_TIME_SEP}?\s*\d{{0,2}}:?\d{{0,2}}\s*{_WEB_AMPM}?\s*{_WEB_TZ})*"
    rf"|\d{{1,2}}(?::\d{{2}})?\s*{_WEB_AMPM}\s*{_WEB_TZ}(?:\s*/\s*\d{{1,2}}(?::\d{{2}})?\s*{_WEB_AMPM}\s*{_WEB_TZ})*"
    # 2026-09-23, per QA feedback (Signum): a real single-city, in-person roundtable invite states
    # only "Wednesday, October 14 12:00 PM" - no timezone at all. Kept as its own fallback since
    # colon-separated minutes + a mandatory explicit am/pm is specific enough to trust without a
    # timezone (unlike the bare "N-M TZ" range above, which needs the timezone as its only signal).
    rf"|\d{{1,2}}:\d{{2}}\s*{_WEB_AMPM}",
    re.I,
)


def _extract_web_event_time(text: str) -> str:
    match = _WEB_EVENT_TIME_RE.search(text or "")
    if not match:
        return ""
    value = re.sub(r"\s+", " ", match.group(0)).strip()
    # Same plausibility guard as email_common.extract_event_time (2026-09-23, per QA feedback,
    # traced to a real garbled "3-0pt" cell) - the bare-range branch can match an ordinary "N-M"
    # number pair immediately followed by a real timezone abbreviation with nothing else marking
    # it as a time. A genuine event time always has minutes (":00") or an explicit am/pm (with or
    # without periods).
    if not re.search(r":|[ap]\.?m\.?", value, re.I):
        return ""
    return value


_CSIS_JSONLD_RE = re.compile(
    r'<script type="application/ld\+json">(.*?)</script>', re.S
)


def _csis_event_page_datetime(link: str):
    """Last-resort fallback for CSIS cards with neither a `data-start-date` attribute nor a
    parseable date in the card text (e.g. "Changing the Narrative on African Investment" -
    confirmed via live check to have no date anywhere on its listing card, yet its own event page
    embeds a full schema.org Event via JSON-LD with a real startDate/endDate). Fetches the
    individual event page and pulls the date/time straight out of that structured data. Returns
    (date_or_None, time_str) - never raises; any fetch/parse failure just yields (None, "")."""
    try:
        page_html = fetch_url(link)
    except Exception:
        return None, ""
    for match in _CSIS_JSONLD_RE.finditer(page_html):
        try:
            data = json.loads(match.group(1))
        except ValueError:
            continue
        nodes = data.get("@graph", [data]) if isinstance(data, dict) else data
        for node in nodes if isinstance(nodes, list) else []:
            if not isinstance(node, dict) or node.get("@type") != "Event":
                continue
            start = node.get("startDate", "")
            if not start:
                continue
            try:
                event_dt = dt.datetime.fromisoformat(start)
            except ValueError:
                continue
            time_str = ""
            if event_dt.hour or event_dt.minute:
                hour12 = event_dt.hour % 12 or 12
                ampm = "AM" if event_dt.hour < 12 else "PM"
                time_str = f"{hour12}:{event_dt.minute:02d}{ampm}"
            return event_dt.date(), time_str
    return None, ""


def _csis_event_page_time(link: str) -> str:
    """2026-09-24, per a 2-week quality sweep: found real CSIS events (confirmed on 3 live pages)
    with NO time on the listing card at all, even though the event's own page always shows one in
    a single clean text node ("September 11, 2026 • 9:00 – 9:30 am EDT") inside
    `.flex-shrink-0` (confirmed unique on the page, not reused elsewhere). Never raises - a fetch/
    parse failure just yields ""."""
    try:
        html = fetch_url(link)
    except Exception:
        return ""
    soup = make_soup(html)
    el = soup.select_one(".flex-shrink-0")
    if el is None:
        return ""
    return _extract_web_event_time(el.get_text(" ", strip=True))


def extract_csis_events(url, source_name):
    """Self-fetching only because the past behavior needed two fetches - now a single plain fetch
    of the configured URL is enough (see _CSIS_CARD_SELECTOR's comment above). Real dates come
    from each card's own `data-start-date` attribute (ISO, more reliable than the human-readable
    text next to it) where present, falling back to the visible date text, and finally (rare) to
    the individual event page's own JSON-LD Event data via `_csis_event_page_datetime` when a card
    exposes neither.

    NOTE: `source_name` (the config entry's own "CSIS Events" name, used for per-source run
    logging) is deliberately NOT used as the article's "source" - that stays plain "CSIS", the
    same value the regular CSIS Analysis scrape AND the email-based CSIS event capture
    (hub_advisor_events_parser.py) both use, so a CSIS event announced via email AND later
    appearing on this page dedupes as the same row instead of a false duplicate."""
    articles, seen = [], set()
    html = fetch_url(url)
    soup = make_soup(html)
    for card in soup.select(_CSIS_CARD_SELECTOR):
        h3 = card.find("h3")
        anchor = h3.find("a") if h3 else None
        if anchor is None:
            continue
        link = urljoin(url, anchor.get("href", ""))
        if link in seen:
            continue
        seen.add(link)
        card_text = card.get_text(" ", strip=True)
        status_div = card.select_one(".event-status")
        start = _get_ci(status_div, "data-start-date") if status_div else None
        published = parse_date(start) or date_from_text(card_text)
        fallback_time = ""
        if published is None:
            published, fallback_time = _csis_event_page_datetime(link)
        title = anchor.get_text(" ", strip=True)
        # "Hosted by X" is real content (which program ran it), kept in Full Text; the date/time
        # prefix visible in card_text is now redundant with the dedicated columns, so full_text
        # here is just the "Hosted by ..." line where present, not the whole raw card_text dump.
        hosted_match = re.search(r"Hosted by\s+.+$", card_text)
        blurb = hosted_match.group(0) if hosted_match else ""
        article = make_article("CSIS", title, link, published, blurb)
        article["is_event"] = True
        # 2026-09-09, per user request: "published" above IS the real event date for
        # every web-scraped events page (extracted directly from each site's own
        # listing) - "announced_date" is the day THIS scrape ran, since these pages
        # don't expose when the org itself first posted the listing.
        article["event_date"] = article["published"]
        article["announced_date"] = dt.date.today().isoformat()
        article["event_time"] = _extract_web_event_time(card_text) or fallback_time or _csis_event_page_time(link)
        article["geography"] = infer_geography(f"{title} {blurb}")
        article["newsletter_name"] = "CSIS Events"
        articles.append(article)
    return articles


def _ac_event_page_details(link: str):
    """Visits the event's OWN page for two things the listing card often doesn't carry at all:
    a real one-line description (the page's `.gta-site-banner--content` summary paragraph - the
    listing card only ever has a status/date and, sometimes, region/topic tags, never a summary)
    and a normalized date/time line (the first `<li>` of `.gta-event-site-banner--event-meta`),
    which is consistently present even for events whose listing card shows no time at all.
    Confirmed live 2026-09-23, per QA feedback, against a real event (the card had neither a
    description nor a time; its own page had both). Never raises - a fetch/parse failure just
    yields ("", ""), and the caller falls back to whatever the listing card already gave."""
    try:
        html = fetch_url(link)
    except Exception:
        return "", ""
    soup = make_soup(html)
    description = ""
    banner = soup.select_one(".gta-site-banner--content")
    if banner is not None:
        description = compact(banner.get_text(" ", strip=True))
    time_str = ""
    meta_li = soup.select_one(".gta-event-site-banner--event-meta li")
    if meta_li is not None:
        time_str = _extract_web_event_time(meta_li.get_text(" ", strip=True))
    return description, time_str


# Real "Load More" pagination confirmed live 2026-09-23, per QA feedback that ~7 real events were
# silently missing from a date range the site does actually list: both the upcoming and
# `?ac-timing=past` variants page 4 events at a time via `&ac-page=N`, and the extractor below
# previously only ever fetched page 1 of each, silently dropping everything past the 4th most
# recent/upcoming event. PAST_LOOKBACK_DAYS bounds the past variant (which otherwise paginates
# back through the site's entire history) to a window wide enough to cover this tool's own
# lookback_days (see sources_config.json) with margin; MAX_PAGES is a hard safety cap against a
# runaway fetch loop if a page's dates are ever unparseable.
_AC_EVENTS_PAST_LOOKBACK_DAYS = 45
_AC_EVENTS_MAX_PAGES = 15


def extract_ac_events(url, source_name):
    """Self-fetching: visits both the default (upcoming) page and its `?ac-timing=past` variant,
    each paginated (see _AC_EVENTS_MAX_PAGES above), and merges by URL. Article "source" stays
    plain "Atlantic Council" (not the config's "Atlantic Council Events" name) for the same
    cross-channel dedup reason documented in extract_csis_events above.

    Each card's own region/issue tags (`.gta-embed--terms`, e.g. "Africa", "Politics &
    Diplomacy") are a real Geography signal, kept for that even though the real description now
    comes from the event's own page (see _ac_event_page_details) rather than from these tags."""
    articles, seen = [], set()
    base_variants = [(url, False), (url.rstrip("/") + "/?ac-timing=past", True)]
    for base_url, is_past in base_variants:
        page = 1
        while page <= _AC_EVENTS_MAX_PAGES:
            page_url = base_url if page == 1 else f"{base_url}{'&' if '?' in base_url else '?'}ac-page={page}"
            html = fetch_url(page_url)
            soup = make_soup(html)
            anchors = soup.select("a.gta-event-embed--link[href]")
            if not anchors:
                break
            new_link_found = False
            oldest_on_page = None
            for anchor in anchors:
                link = urljoin(url, anchor.get("href", ""))
                if link in seen:
                    continue
                seen.add(link)
                new_link_found = True
                title_el = anchor.select_one(".gta-event-embed--title")
                date_el = anchor.select_one(".gta-event-embed--heading--date")
                title = title_el.get_text(" ", strip=True) if title_el else anchor.get_text(" ", strip=True)
                date_text = date_el.get_text(" ", strip=True) if date_el else ""
                published = date_from_text(date_text)
                if published and (oldest_on_page is None or published < oldest_on_page):
                    oldest_on_page = published
                tags = [t.get_text(strip=True) for t in anchor.select(".gta-embed--terms a")]
                page_description, page_time = _ac_event_page_details(link)
                blurb = page_description or (("Tags: " + ", ".join(tags)) if tags else "")
                article = make_article("Atlantic Council", title, link, published, blurb)
                article["is_event"] = True
                # 2026-09-09, per user request: "published" above IS the real event date for
                # every web-scraped events page (extracted directly from each site's own
                # listing) - "announced_date" is the day THIS scrape ran, since these pages
                # don't expose when the org itself first posted the listing.
                article["event_date"] = article["published"]
                article["announced_date"] = dt.date.today().isoformat()
                article["event_time"] = page_time or _extract_web_event_time(date_text)
                article["geography"] = infer_geography(f"{title} {' '.join(tags)} {page_description}")
                article["newsletter_name"] = "Atlantic Council Events"
                articles.append(article)
            if not new_link_found:
                break
            if is_past and oldest_on_page and (dt.date.today() - oldest_on_page).days > _AC_EVENTS_PAST_LOOKBACK_DAYS:
                break
            has_next_page = soup.select_one(f'a[href*="ac-page={page + 1}"]')
            if not has_next_page:
                break
            page += 1
    return articles


def _cfr_event_page_details(link: str):
    """Visits the event's OWN page for its real time and description - the listing card carries
    neither (only a one-word type tag, see below). Confirmed live 2026-09-23, per QA feedback
    that CFR's Event Time was blank for every row: the page's "Event date" block has several
    `.type-mono-ui3` <time> tags (weekday, calendar date, and the actual clock time in no fixed
    position), so each is tried through the same time-extraction regex rather than assumed to be
    the last one; the description is the page's own `.rich-text.type-serif-body4` paragraph.
    Never raises - a fetch/parse failure just yields ("", "")."""
    try:
        html = fetch_url(link)
    except Exception:
        return "", ""
    soup = make_soup(html)
    time_str = ""
    for time_el in soup.select("time.type-mono-ui3"):
        time_str = _extract_web_event_time(time_el.get_text(" ", strip=True))
        if time_str:
            break
    description = ""
    desc_el = soup.select_one(".rich-text.type-serif-body4")
    if desc_el is not None:
        description = compact(desc_el.get_text(" ", strip=True))
    return description, time_str


def extract_cfr_events(html, base_url, source_name):
    """CFR's /events page is a retrospective archive (confirmed by the user directly, not just
    inferred) - it lists recently-past events, not genuinely upcoming ones, and no other CFR
    channel (checked: every real captured CFR email, plus an /events/upcoming URL) surfaces
    anything newer. Kept anyway since recently-past events are still real content worth having.
    Article "source" stays plain "CFR" for the same cross-channel dedup reason as CSIS/AC above."""
    soup = make_soup(html)
    articles, seen = [], set()
    for anchor in soup.select('a[href^="/event/"]'):
        link = urljoin(base_url, anchor.get("href", ""))
        if link in seen:
            continue
        seen.add(link)
        card = anchor
        for _ in range(6):
            if card.parent is None:
                break
            card = card.parent
            if card.find("h3") and card.find("time"):
                break
        h3 = card.find("h3")
        title = h3.get_text(" ", strip=True) if h3 else anchor.get_text(" ", strip=True)
        if not title:
            continue
        time_tag = card.find("time")
        published = parse_date(preferred_time_tag_date(time_tag)) if time_tag is not None else None
        # The only other real text on this card is a one-word type tag ("Meeting", "Webinar",
        # etc., confirmed 2026-09-09) - no separate summary paragraph exists on this listing, so
        # the real description (and real time) come from the event's own page instead.
        type_span = card.find("span")
        page_description, page_time = _cfr_event_page_details(link)
        blurb = page_description or (type_span.get_text(strip=True) if type_span else "")
        article = make_article("CFR", title, link, published, blurb)
        article["is_event"] = True
        # 2026-09-09, per user request: "published" above IS the real event date for
        # every web-scraped events page (extracted directly from each site's own
        # listing) - "announced_date" is the day THIS scrape ran, since these pages
        # don't expose when the org itself first posted the listing.
        article["event_date"] = article["published"]
        article["announced_date"] = dt.date.today().isoformat()
        article["event_time"] = page_time or _extract_web_event_time(card.get_text(" ", strip=True))
        article["geography"] = infer_geography(f"{title} {page_description}")
        article["newsletter_name"] = "CFR Events"
        articles.append(article)
    return articles


def extract_chatham_events(url, source_name):
    """Chatham House blocks plain HTTP clients the same way its regular Expert Comment page does
    (see extract_chatham_via_browser above) - uses the same Playwright fetch. Same card markup as
    extract_chatham() above (`.teaser__title`/`.teaser__summary` inside `a.teaser__link`), just
    filtered to `/events/all/` links instead of `/20\\d{2}/|/publications/`. Only shows genuinely
    upcoming events (no past-events variant needed/available). Article "source" stays plain
    "Chatham House" for the same cross-channel dedup reason as CSIS/AC/CFR above."""
    html = _fetch_html_via_browser(url, wait_ms=2500)
    soup = make_soup(html)
    articles, seen = [], set()
    for anchor in soup.select("a.teaser__link[href]"):
        href = anchor["href"]
        if "/events/all/" not in href:
            continue
        link = urljoin(url, href)
        if link in seen:
            continue
        seen.add(link)
        title_el = anchor.select_one(".teaser__title")
        title = title_el.get_text(" ", strip=True) if title_el else anchor.get_text(" ", strip=True)
        summary_el = anchor.select_one(".teaser__summary")
        summary = summary_el.get_text(" ", strip=True) if summary_el else ""
        card_text = anchor.get_text(" ", strip=True)
        published = date_from_text(card_text)
        article = make_article("Chatham House", title, link, published, summary)
        article["is_event"] = True
        # 2026-09-09, per user request: "published" above IS the real event date for
        # every web-scraped events page (extracted directly from each site's own
        # listing) - "announced_date" is the day THIS scrape ran, since these pages
        # don't expose when the org itself first posted the listing.
        article["event_date"] = article["published"]
        article["announced_date"] = dt.date.today().isoformat()
        article["event_time"] = _extract_web_event_time(card_text)
        article["geography"] = infer_geography(f"{title} {summary}")
        article["newsletter_name"] = "Chatham House Events"
        articles.append(article)
    return articles


def extract_wef_meetings(url, source_name):
    """WEF blocks plain HTTP clients the same way its Publications page does (see
    extract_wef_publications above) - uses the same Playwright fetch. Only lists WEF's own
    flagship annual meetings (a handful per year, not frequent webinars), each with a real
    "D-D Month YYYY" date range in the surrounding card text - only the START day/month/year is
    kept as the published date. Article "source" stays plain "WEF" (matching the regular WEF
    Publications source) since nothing else currently captures WEF events via another channel."""
    html = _fetch_html_via_browser(url, wait_ms=2500)
    soup = make_soup(html)
    articles, seen = [], set()
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"]
        if "/meetings/" not in href or not re.search(r"/meetings/[a-z0-9-]+-\d{4}/?$", href):
            continue
        link = urljoin(url, href)
        if link in seen:
            continue
        seen.add(link)
        card = anchor
        card_text = None
        for _ in range(5):
            if card.parent is None:
                break
            card = card.parent
            text = card.get_text(" ", strip=True)
            if _WEF_DATE_RANGE_RE.search(text):
                card_text = text
                break
        if not card_text:
            continue
        m = _WEF_DATE_RANGE_RE.search(card_text)
        published = parse_date(f"{m.group(1)} {m.group(2)} {m.group(3)}") if m else None
        # The card's own <a> wraps only its thumbnail image (no text) - the real title is the
        # text immediately BEFORE the "D-D Month YYYY" date range in the card's combined text
        # (confirmed 2026-09-09: card_text is exactly "<title> <date range> <location>").
        title = card_text[:m.start()].strip() if m else card_text[:120]
        # 2026-09-11, per reviewer feedback ("event date, time and location collated under full
        # text"): card_text is exactly "<title> <date range> <location>" (see comment above) - the
        # location is whatever's left AFTER the date range, kept as Full Text on its own instead
        # of the whole raw blob repeating the title/date already in their own columns. WEF's real
        # dates are multi-day RANGES (a conference spanning several days), not a single start
        # time, so there's no separate "Event Time" to extract here - it stays blank.
        location = card_text[m.end():].strip() if m else ""
        article = make_article("WEF", title, link, published, location)
        article["is_event"] = True
        # 2026-09-09, per user request: "published" above IS the real event date for
        # every web-scraped events page (extracted directly from each site's own
        # listing) - "announced_date" is the day THIS scrape ran, since these pages
        # don't expose when the org itself first posted the listing.
        article["event_date"] = article["published"]
        article["announced_date"] = dt.date.today().isoformat()
        article["geography"] = infer_geography(f"{title} {location}")
        article["newsletter_name"] = "WEF Meetings"
        articles.append(article)
    return articles


# China Macro Group's own /webinar page (given directly by the user) is a hub with no dated
# listings of its own - just 3 category "threads", each behind a "Read More" link. Checked all of
# them directly 2026-09-09: Thread 1 ("Staying in Dialogue with China") and Thread 3 ("China's
# Foreign Affairs and Geopolitics") share the SAME real listing page (both "Read More" links on
# the hub point at the identical URL - confirmed by fetching it, its real webinars include both
# domestic-policy and geopolitics/Trump/Taiwan topics), and Thread 2 ("Sector Dialogues") has its
# own separate page. These two are the real, permanent listing pages - hardcoded here the same
# way extract_chinamacro() above hardcodes its 4 known topic-page paths.
_CHINAMACRO_WEBINAR_PAGES = ("staying-in-dialgoue-with-china-and-flash", "sector-dialogues")

# Real dates on this site render two different ways depending on the webinar's age - "28 August,
# 2026" (comma before year, no ordinal) for newer ones, "5th April, 2024" / "19th June 2024" /
# "11th September 2024" (ordinal suffix, comma optional) for the 2024 archive - confirmed against
# real cards, not assumed. Older cards ALSO sometimes fold date+time into one field instead of two
# separate ones ("5th April, 2024 9-945am CEST | 3-345pm CST"), which shifts every field after it
# over by one - so title is found by CONTENT (the first field that isn't the date and doesn't look
# like a bare time range), not by a fixed index.
_CHINAMACRO_ORDINAL_RE = re.compile(r"(\d)(?:st|nd|rd|th)\b", re.I)
_CHINAMACRO_TIME_RANGE_RE = re.compile(r"^\d{1,2}[-:]\d+\s*(?:am|pm)", re.I)


def _chinamacro_webinar_date(text: str):
    return parse_date(_CHINAMACRO_ORDINAL_RE.sub(r"\1", text).replace(",", ""))


def extract_chinamacro_webinars(url, source_name):
    """Self-fetching (plain HTTP, no WAF block found). Each webinar is a Wix "repeater" card
    (`.wixui-repeater__item`) whose fields are a sequence of `p.font_8` paragraphs - confirmed
    directly against real cards on both pages, but the shape isn't perfectly uniform (see
    _CHINAMACRO_ORDINAL_RE's comment above), so the date field is found by trying to parse EACH
    field rather than assuming a fixed index, and the title is the first remaining field that
    isn't a bare time range."""
    articles, seen = [], set()
    for page in _CHINAMACRO_WEBINAR_PAGES:
        page_url = f"https://www.chinamacro.ch/{page}"
        html = fetch_url(page_url)
        soup = make_soup(html)
        for card in soup.select(".wixui-repeater__item"):
            link_tag = card.select_one('a[href*="/post/"]')
            if link_tag is None:
                continue
            link = urljoin(page_url, link_tag["href"])
            if link in seen:
                continue
            seen.add(link)
            fields = [p.get_text(" ", strip=True) for p in card.select("p.font_8")]
            if len(fields) < 2:
                continue
            published, date_idx = None, None
            for i, field in enumerate(fields):
                candidate = _chinamacro_webinar_date(field)
                if candidate:
                    published, date_idx = candidate, i
                    break
            # This site writes time compactly without a colon on the second number ("9-945am
            # CEST" = 9:00-9:45am) - a shape _WEB_EVENT_TIME_RE doesn't parse (built around other
            # sites' "9:00-9:45am" style) - so the time field is identified by POSITION here (via
            # the same _CHINAMACRO_TIME_RANGE_RE already used to skip it for title-detection)
            # rather than re-extracted with a regex, and used as-is.
            title, event_time = None, ""
            for field in fields[(date_idx + 1 if date_idx is not None else 0):]:
                if _CHINAMACRO_TIME_RANGE_RE.match(field):
                    event_time = field
                    continue
                title = field
                break
            if not title:
                continue
            # 2026-09-11, per reviewer feedback: keep the real descriptive fields (title, any
            # Chinese-language subtitle, speaker) in Full Text but drop the date/time field(s)
            # now that they live in their own dedicated columns - date_idx marks whichever field
            # was consumed as the date; a bare time-range field (if date+time were split into two
            # fields) is dropped the same way title-detection above skips it.
            other_fields = [
                f for i, f in enumerate(fields)
                if i != date_idx and not _CHINAMACRO_TIME_RANGE_RE.match(f)
            ]
            blurb = " ".join(other_fields)
            # "source" stays plain "China Macro Group" (not the config's own "China Macro Group
            # Webinars" name) - matches the regular chinamacro source, same cross-channel-
            # consistency reasoning as CSIS/AC/CFR/Chatham House/MEC above.
            article = make_article("China Macro Group", title, link, published, blurb)
            article["is_event"] = True
            article["event_date"] = article["published"]
            article["announced_date"] = dt.date.today().isoformat()
            article["event_time"] = event_time
            article["geography"] = infer_geography(f"{title} {blurb}") or "China"
            article["newsletter_name"] = "China Macro Group Webinars"
            articles.append(article)
    return articles


# MEC's own time format is unlike every other source - "13:00:00GMT - 16:00:00GMT" (seconds
# always included, no space before the timezone letters) - confirmed 2026-09-11 against a real
# card, handled with its own small pattern rather than forcing it into _WEB_EVENT_TIME_RE (whose
# shape - built around other sites' "9:00-9:45am CEST" style - doesn't fit this one).
_MEC_TIME_RE = re.compile(r"\d{1,2}:\d{2}:\d{2}[A-Za-z]{2,5}\s*-\s*\d{1,2}:\d{2}:\d{2}[A-Za-z]{2,5}")


def extract_mec_events(url, source_name):
    """Middle East Council's events page renders both listings entirely client-side (a plain
    fetch of the base URL returns zero events server-rendered - confirmed 2026-09-09), and its
    own individual event permalinks (`?post_type=events-list&p=...`) redirect-loop for a plain
    HTTP client too - both need a real browser. The page defaults to showing "Previous Events";
    "Upcoming Events" requires clicking its filter tab (`a.filter-btn[data-date='upcoming']`),
    so both are visited in one Playwright session and merged."""
    from playwright.sync_api import sync_playwright

    def extract(html):
        soup = make_soup(html)
        out = []
        for anchor in soup.find_all("a", href=True):
            if "post_type=events-list" not in anchor["href"]:
                continue
            title = anchor.get_text(" ", strip=True)
            if not title:
                continue
            card = anchor
            card_text = None
            for _ in range(5):
                if card.parent is None:
                    break
                card = card.parent
                text = card.get_text(" ", strip=True)
                if date_from_text(text):
                    card_text = text
                    break
            out.append((title, anchor["href"], card_text))
        return out

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)
        past_html = page.content()
        try:
            page.click("a.filter-btn[data-date='upcoming']", timeout=5000)
            page.wait_for_timeout(2500)
            upcoming_html = page.content()
        except Exception:
            upcoming_html = ""
        browser.close()

    articles, seen = [], set()
    for html in (past_html, upcoming_html):
        if not html:
            continue
        for title, href, card_text in extract(html):
            link = urljoin(url, href)
            if link in seen:
                continue
            seen.add(link)
            published = date_from_text(card_text) if card_text else None
            # "source" stays the full org name (matching the email vendor entry in vendors.json)
            # for the same cross-channel dedup reason as CSIS/AC/CFR/Chatham above.
            # 2026-09-11, per reviewer feedback: strip the date/time (own dedicated columns now)
            # and the trailing type tag ("Webinar") out of the stored blurb, keeping just the
            # title itself - card_text here doesn't carry any extra description beyond that.
            time_match = _MEC_TIME_RE.search(card_text) if card_text else None
            event_time = time_match.group(0) if time_match else ""
            article = make_article("Middle East Council on Global Affairs", title, link, published, "")
            article["is_event"] = True
            # 2026-09-09, per user request: "published" above IS the real event date for
            # every web-scraped events page (extracted directly from each site's own
            # listing) - "announced_date" is the day THIS scrape ran, since these pages
            # don't expose when the org itself first posted the listing.
            article["event_date"] = article["published"]
            article["announced_date"] = dt.date.today().isoformat()
            article["event_time"] = event_time
            article["geography"] = infer_geography(title) or "Middle East"
            article["newsletter_name"] = "Middle East Council on Global Affairs Events"
            articles.append(article)
    return articles


def extract_generic(html, base_url, source_name, path_patterns=None):
    """Fallback: collect anchors whose href matches a path pattern and whose text
    is long enough to look like a headline. Dates are left blank."""
    soup = make_soup(html)
    articles, seen = [], set()
    for anchor in soup.find_all("a", href=True):
        href = anchor["href"]
        if path_patterns and not any(p in href for p in path_patterns):
            continue
        title = anchor.get_text(" ", strip=True)
        if len(title) < 20:
            continue
        url = urljoin(base_url, href)
        if url in seen:
            continue
        seen.add(url)
        articles.append(make_article(source_name, title, url, None, ""))
    return articles


# Per-source full-article-body container(s), verified against real article pages from each site.
FULL_TEXT_SELECTORS = {
    "Atlantic Council": [".article-main", ".article-content"],
    "CSIS": [".article--layout-body", ".text-block.ts-friendly"],
    "Chatham House": [".body-content"],
    # Lowy Institute: handled specially in _lowy_full_text() — its body is split across
    # multiple div.prose blocks, and naively taking the first .prose only gets the one-line dek.
    # China Macro Group (Wix): rich content lazy-loads client-side after initial page load, so
    # only a short excerpt is ever present in the static HTML — a real ceiling for a plain HTTP
    # fetch, not a bug. [data-hook='post-page'] is still the right container, it's just short.
    "China Macro Group": ["[data-hook='post-page']"],
    # CFR: no dedicated selector needed — the generic <article> fallback below already gets a
    # clean paragraph list on CFR's article template.
}

_MIN_VIABLE_FULL_TEXT_CHARS = 150

# Deloitte Insights articles run acknowledgments/editorial-credits/cross-promo/legal boilerplate
# after the real body, in this rough order - once any of these is seen, everything from there on
# is stopped rather than individually filtered (more robust than an exact per-line skip list,
# since the exact boilerplate wording varies slightly article to article). Also folded into
# _JUNK_PARAGRAPH_PREFIXES below so the generic (non-Deloitte-specific) full-text fallbacks skip
# these paragraphs too, instead of accepting Deloitte's own site-wide legal footer as "content"
# when a page (e.g. a JS-rendered tracker/hub) has no real .cmp-text body for
# _deloitte_full_text() to find.
_DELOITTE_TAIL_MARKERS = (
    "the authors would like to", "the authors would also like to",
    "this study would not have been possible", "editorial (including production",
    "knowledge services:", "access more insights for",
    "deloitte insights and our research centers", "learn about deloitte",
    "deloitte refers to one or more of deloitte touche tohmatsu",
)

# Paragraphs that are real HTML <p> tags but are boilerplate, not article content.
_JUNK_PARAGRAPH_PREFIXES = (
    "sign up to receive", "sign up for", "subscribe to", "related experts:",
    "read more here", "share this", "reach out to discuss:",
    "for personalized content and settings",
    # Found 2026-08-05 during a full-text quality audit: CSIS runs a standalone copyright
    # paragraph, plus (on their "Commentary" template specifically) a standalone "who publishes
    # this" disclaimer paragraph, as the last 1-2 <p> tags of the real content container.
    "© ", "commentary is produced by the center for strategic",
    # 2026-09-24, per user feedback: a real, recurring CSIS lead-image credit line ("Photo: Houthi
    # Media Center/Handout via Xinhua/Getty Images") was showing up glued onto the start of the
    # real text. Confirmed real across 48 rows, ALWAYS its own short (40-101 char) standalone
    # paragraph, never a real sentence that happens to start this way - safe to drop as a whole
    # paragraph, same as the copyright/disclaimer lines above.
    "photo: ",
) + _DELOITTE_TAIL_MARKERS

# Paragraphs whose OWN class marks them as non-content (author byline blocks, etc.).
_JUNK_PARAGRAPH_CLASSES = {"author-teaser__meta"}

# 2026-09-23, per QA feedback ("missing CFR subheadings, text between subheadings not captured
# in full"): real bug, and real across every source using _extract_paragraphs() (or the Lowy/
# Deloitte equivalents below), not just CFR - confirmed live on 4 real articles. The old code only
# ever called container.find_all("p"), which structurally cannot see an <h2>/<h3>/<h4> tag at
# all - Chatham House alone was found to be silently dropping real, substantial subheadings
# ("Deepening military ties", "Mutually beneficial", "Battlefield experience", "Ramifications
# for East Asia") from a single article, with the paragraphs before and after each one silently
# concatenated into one continuous blob with nothing marking where one section ended and the next
# began. That's very likely also the real explanation for testers separately describing the text
# BETWEEN subheadings as sometimes "not accurate" - it's not that the words themselves were
# wrong, but that two genuinely different subsections' content was blending together unmarked
# with no indication a real section break belonged there.
#
# Headings are kept even when shorter than min_paragraph_chars (a real subheading is very often
# under that, e.g. "Mutually beneficial" - only 20 chars), prefixed with "### " (a plain,
# unambiguous visual marker in a plain-text cell) so a reader can tell heading from body text.
# Bare numbering artifacts confirmed real on Atlantic Council (lone "<h4>2.</h4>", "<h4>3.</h4>" -
# footnote/citation markers, not real subheadings) are filtered out as noise.
_BARE_HEADING_NUMBER_RE = re.compile(r"^\d+\.?$")
_JUNK_HEADING_TEXTS = {
    "available downloads",  # CSIS: a UI section label, not real content
    "information",  # CFR: a bare section label immediately preceding the Creative Commons
                     # footer (see _GENERIC_TAIL_MARKERS below) - not real content either.
    "programs & projects",  # CSIS: a generic nav/taxonomy label, same reasoning as "tags" below -
                             # skipped individually (never a stop-trigger, see why below).
    "tags",  # CSIS: skipped individually - NOT a stop-trigger. Real counter-example found live
             # 2026-09-23 during the false-positive sweep this fix's own docstring insists on: one
             # real article's heading order is author-names -> Programs & Projects -> Tags ->
             # Appendix (a REAL, substantive methodology section - genuine dollar figures and data
             # notes) -> Related Content -> unrelated-article carousel. Treating "Tags" as a stop
             # point had silently cut the real Appendix - exactly the over-deletion risk this task
             # was explicitly warned against. "Tags" is just a generic nav label that can appear
             # BEFORE real content, not reliably at the true end, so it's now skipped (not
             # trusted as a boundary) rather than treated as "everything after this is junk".
    # 2026-09-24, per user feedback: found live, immediately before the "Additional Contributions"
    # credits block (see _EXACT_TAIL_MARKERS below), a run of bare CFR mini-toc headings with NO
    # paragraph text under any of them at all - confirmed on a real article (nothing followed
    # these three before hitting "Additional Contributions"). Skipped individually rather than
    # treated as a stop point, same reasoning as "tags" above.
    "colophon", "staff writers", "data visualization",
}

# 2026-09-23, per QA feedback ("remove extraneous elements such as author names and references"):
# checked real, already-scraped Full Text across every source before touching anything (per this
# project's own established practice) rather than guessing, and re-verified with a broader
# false-positive sweep across 30 more real rows before trusting any of this (see "tags" above for
# the one real problem that sweep caught and how it was corrected). What actually stayed reliable
# after that sweep:
#   - Chatham House: EVERY one of 6 real samples checked ends with the identical bare sequence
#     "Topics" / "Regions" / "Departments" - checking for the first one alone is enough.
#   - CSIS: "Related Content" reliably marks the start of a genuinely unrelated-article carousel
#     (each entry has its own distinct title + byline + date, clearly a different piece) - and,
#     confirmed directly on the one article that exposed the "Tags" problem, "Related Content"
#     itself sits AFTER the real Appendix section, never before it - unlike "Tags", which sits
#     before real content in that same article. Only "Related Content" is trusted as a stop point;
#     "Tags" is deliberately NOT (see _JUNK_HEADING_TEXTS above).
# Deliberately EXACT match, not a prefix/startswith check like _GENERIC_TAIL_MARKERS below - these
# are short, generic words a genuine, longer real subheading could plausibly start with (e.g. a
# real "Topics for the Next Administration to Consider" heading must NOT be truncated just because
# it starts with "Topics"), so only a heading/paragraph that is EXACTLY one of these bare labels
# and nothing else triggers the stop.
#   - CFR: "Additional Contributions" (a real <h3>, confirmed live 2026-09-24) reliably marks the
#     start of a production-credits block - a list of staff names/titles (Editorial Director,
#     Senior Photo Editor, Data Visualization, ...) on CFR's multimedia-heavy features, followed
#     by an "Additional Photo Credits" note. Real, attributable people, but not article content -
#     same reasoning as "Related Content" above, and it only ever appears after the genuine body
#     text ends on every real example checked.
_EXACT_TAIL_MARKERS = {"related content", "topics", "additional contributions"}

# A byline masquerading as a heading - two real shapes confirmed, not hypothetical: "By experts
# and staff" / "By James M. Lindsay" (CFR), and a bare name with nothing else, e.g. "Ksenia
# Svetlova" (Chatham House), "William Alan Reinsch" (CSIS).
#
# The "By <Name>" shape is safe to check anywhere in a document - a real subheading essentially
# never starts with the literal word "By" followed by a capitalized name.
#
# The bare-name shape (no "By" prefix) is a different risk profile: it reuses the same structural
# test already proven safe on 37 real MEC emails (_looks_like_mec_author_name in
# generic_vendor_parser.py - every word capitalized, 2-4 words, no digits/colon), but that test
# alone COULD misfire on a genuine short, title-case section heading with no digits, e.g. "USMCA
# Review" or "Key Takeaways" (a real subheading like "Mutually beneficial"/"Deepening military
# ties" already fails it safely, since those aren't all-words-capitalized - but a short, generic,
# title-case heading might still exist somewhere and would look identical in shape to a name).
# Every real bare-name-as-heading case actually found appeared in an unbroken RUN of such
# headings starting from the very first one (right after the lead paragraph) - a real CSIS
# multi-author brief confirmed live has three, one per co-author ("Yasir Atalan" / "Erik
# Tiersten-Nyman" / "Benjamin Jensen") back to back before any real subheading. So this shape is
# only ever checked while still inside that leading run, never later in the document once a
# non-name heading has broken it - keeps the false-positive risk near zero rather than merely low.
_BYLINE_HEADING_RE = re.compile(r"^By\s+[A-Z]", re.I)
_BARE_NAME_HEADING_RE = re.compile(r"^[A-ZÀ-ÖØ-Þ][\wÀ-ÖØ-öø-ÿ'.-]*(?:\s+[A-ZÀ-ÖØ-Þ][\wÀ-ÖØ-öø-ÿ'.-]*){1,3}$")


def _is_byline_heading(text: str, in_leading_run: bool) -> bool:
    if _BYLINE_HEADING_RE.match(text):
        return True
    if not in_leading_run:
        return False
    if ":" in text or any(ch.isdigit() for ch in text) or len(text) > 40:
        return False
    return bool(_BARE_NAME_HEADING_RE.match(text))

# 2026-09-23, found live while verifying the subheadings fix above: a real CFR article's generic
# <article> container also runs real content straight into trailing boilerplate with no
# separation at all - a "Creative Commons: Some rights reserved." licensing footer, a bare "Tags"
# label + the tag list, and then (confirmed real, not hypothetical) a completely UNRELATED
# "related article" teaser for a different piece entirely ("Xi Jinping Visits the White House" /
# "By James M. Lindsay") - all of which were being silently included as if they were part of
# THIS article's own Full Text. Very likely part of what testers meant by the text "not being
# accurate". "Creative Commons" is CFR's own standard, recurring license footer (a nonprofit,
# applied site-wide) - once seen, everything from there on is stopped rather than filtered line
# by line, same truncate-at-first-match approach _deloitte_full_text() already uses for its own
# tail boilerplate.
_GENERIC_TAIL_MARKERS = ("creative commons",)

# 2026-09-23: footnote reference markers - confirmed live on a real Deloitte Insights article
# ("...creator and social platforms,<sup><a href="#endnote-2">2</a></sup> the brand-creator...")
# were being read straight into the sentence as a stray floating digit ("...platforms, 2 the
# brand-creator..."), since get_text() has no way to know a <sup> is a citation marker rather than
# real words. Stripped at the DOM level before text extraction (the only safe way - a regex
# cleanup afterward couldn't reliably tell an injected footnote "3" apart from a real "3" that
# belongs in the sentence, e.g. "3 million people"). Scoped narrowly to a <sup> that contains
# ONLY a footnote-shaped anchor link (href starting with "#endnote"/"#fn"/"#footnote"/"#cite"), not
# every <sup> tag on the page - a genuine superscript used for something else (e.g. a trademark
# symbol or a chemical formula) would not match this and is left untouched.
_FOOTNOTE_HREF_RE = re.compile(r"^#(endnote|fn|footnote|cite)", re.I)


def _strip_footnote_markers(container) -> None:
    for sup in container.find_all("sup"):
        anchor = sup.find("a", href=True)
        if anchor is not None and _FOOTNOTE_HREF_RE.match(anchor["href"]) and sup.get_text(strip=True) == anchor.get_text(strip=True):
            sup.decompose()


def _is_junk_paragraph(text: str) -> bool:
    lower = text.lower()
    return any(lower.startswith(prefix) for prefix in _JUNK_PARAGRAPH_PREFIXES)


def _is_real_heading(text: str) -> bool:
    return bool(text) and not _BARE_HEADING_NUMBER_RE.match(text) and text.lower() not in _JUNK_HEADING_TEXTS


def _extract_paragraphs(container, min_paragraph_chars: int):
    _strip_footnote_markers(container)
    paragraphs = []
    in_leading_byline_run = True
    for tag in container.find_all(["h2", "h3", "h4", "p", "li"]):
        if tag.name == "li":
            # 2026-10-01, per QA feedback: a "decent chunk" of text was silently missing from
            # certain CSIS articles - confirmed live on both named examples that the article uses
            # one or more <ol> (numbered) lists for genuine, substantive enumerated analysis
            # (780-10,676 real characters each), which this loop never looked at before. A <ul>
            # on the same real pages is consistently just a "Programs & Projects" tag/category
            # list, not real content, so only <ol> list items are captured here, same treatment
            # as a real paragraph.
            if tag.find_parent("ol") is None:
                continue
            text = compact(tag.get_text(" ", strip=True))
            if len(text) >= min_paragraph_chars and not _is_junk_paragraph(text):
                paragraphs.append(f"- {text}")
            continue
        if tag.name == "p" and tag.find_parent("li") is not None:
            continue  # already captured whole as part of its enclosing <li> above
        text = compact(tag.get_text(" ", strip=True))
        if text.lower() in _EXACT_TAIL_MARKERS or any(text.lower().startswith(marker) for marker in _GENERIC_TAIL_MARKERS):
            break
        if tag.name != "p":
            is_byline = _is_byline_heading(text, in_leading_byline_run)
            if not is_byline:
                in_leading_byline_run = False
            if _is_real_heading(text) and not is_byline:
                paragraphs.append(f"### {text}")
            continue
        if _JUNK_PARAGRAPH_CLASSES.intersection(tag.get("class") or []):
            continue
        if len(text) >= min_paragraph_chars and not _is_junk_paragraph(text):
            paragraphs.append(text)
    return paragraphs


def _lowy_full_text(soup) -> str:
    """Lowy's article body is split across one or more div.prose blocks (a multi-author
    round-up runs each author's take in its own block). Only the plain
    {"prose", "dark:prose-invert"} blocks are the real body — naively taking the page's first
    .prose returns just the one-sentence dek."""
    blocks = [el for el in soup.select(".prose") if set(el.get("class", [])) == {"prose", "dark:prose-invert"}]
    paragraphs = []
    for block in blocks:
        paragraphs.extend(_extract_paragraphs(block, min_paragraph_chars=40))
    return "\n\n".join(paragraphs)


def _deloitte_full_text(soup) -> str:
    """Deloitte Insights splits an article body across several .cmp-text blocks (not one single
    container like most sites), and appends acknowledgments/editorial-credits/cross-promo/legal
    boilerplate after the real content - truncate at the first _DELOITTE_TAIL_MARKERS match
    rather than filtering each boilerplate line individually (wording varies slightly).

    2026-09-23: also walks h2/h3/h4 alongside p (see _extract_paragraphs' docstring for the full
    story) - confirmed live, a real Deloitte "Weekly Global Economic Update" had 14 real
    subheadings (its per-region sections) previously dropped entirely. Also strips footnote
    markers (see _strip_footnote_markers - this is the exact source a real stray inline "2"/"3"/
    "4" was confirmed on) before reading any text out."""
    paragraphs = []
    in_leading_byline_run = True
    for block in soup.select(".cmp-text"):
        _strip_footnote_markers(block)
        for tag in block.find_all(["h2", "h3", "h4", "p"]):
            if tag.name != "p":
                text = compact(tag.get_text(" ", strip=True))
                is_byline = _is_byline_heading(text, in_leading_byline_run)
                if not is_byline:
                    in_leading_byline_run = False
                if _is_real_heading(text) and not is_byline:
                    paragraphs.append(f"### {text}")
                continue
            text = compact(tag.get_text(" ", strip=True))
            if len(text) < 40:
                continue
            lower = text.lower()
            if any(lower.startswith(marker) for marker in _DELOITTE_TAIL_MARKERS):
                return "\n\n".join(paragraphs)
            if _is_junk_paragraph(text):
                continue
            paragraphs.append(text)
    return "\n\n".join(paragraphs)


def _wef_pdf_full_text(article_url: str) -> str:
    """WEF's real content is the linked PDF report, not the landing page. Found 2026-08-17: most
    current WEF listing cards (12 of 19 checked) have no description at all on the listing page,
    so the old listing-description fallback alone now leaves most WEF rows empty - worth actually
    fetching the PDF. The article page itself is on weforum.org, which blocks plain HTTP clients
    (same block as the listing, see extract_wef_publications), so it's visited via the same
    headless-browser helper already built for Chatham House. The PDF itself is hosted on
    reports.weforum.org, a DIFFERENT domain with no bot-block at all - confirmed directly (5/5 real
    reports downloaded via a plain request, no Playwright needed for that part). Returns "" on any
    failure (no PDF link found, download error, unparseable PDF) - best-effort, same contract as
    every other fetch_full_text() path; caller falls back to the listing summary, then the
    "nothing available" placeholder if that's empty too."""
    if pypdf is None:
        return ""
    try:
        soup = make_soup(_fetch_html_via_browser(article_url))
        pdf_link = None
        for a in soup.find_all("a", href=True):
            if ".pdf" in a["href"].lower():
                pdf_link = urljoin(article_url, a["href"])
                break
        if not pdf_link:
            return ""
        response = requests.get(pdf_link, headers={"User-Agent": USER_AGENT}, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        reader = pypdf.PdfReader(io.BytesIO(response.content))
        pages_text = [page.extract_text() or "" for page in reader.pages]
        return "\n\n".join(t for t in pages_text if t)
    except Exception:
        return ""


def fetch_full_text(url: str, source_name: str, min_paragraph_chars: int = 40) -> str:
    """Best-effort FULL article body (distinct from the short listing summary): visits the
    article's own page and joins the paragraph text from its known content container. Returns
    "" if nothing viable is found (caller falls back to the short summary)."""
    if source_name == "WEF":
        return _wef_pdf_full_text(url)

    try:
        # 2026-10-22, per user report ("many Middle East Council articles seem like just
        # headlines"): MEC's per-headline wire-story links (Reuters/AP/etc., see
        # generic_vendor_parser.parse_mec_daily_brief) were going through the plain-request path
        # below and failing ~66% of the time (64/97 real rows checked) - most wire services block
        # or render nothing useful to a bot-less HTTP client, the same reason Chatham House/WEF
        # already needed the browser path. MEC's own two bylined features (fetched separately via
        # generic_vendor_parser._mec_article_full_text) already used a real browser and succeeded
        # far more often, which is what pointed at this fix.
        #
        # 2026-10-22, per user request: "Atlantic Council Highlights" is the per-item link under
        # Inflection Points' trailing "Highlights from the Atlantic Council" list (see
        # hub_advisor_events_parser._extract_ac_highlights_items) - same SafeLinks+HubSpot
        # tracking-redirect chain as MEC's links above, confirmed live that a plain request
        # returns nothing (0 chars) while the browser path resolves it fine (atlanticcouncil.org
        # itself has no heavy bot-protection, unlike Reuters/AP - the browser is only needed here
        # to walk the multi-hop tracking redirect, not to dodge anti-bot measures).
        #
        # 2026-10-22: CFR's Regional Roundup links briefly had their own sentinel here too, but per
        # user correction those rows no longer fetch the linked article at all (Link is kept,
        # Full Text is just the brief's own blurb) - see hub_advisor_events_parser.py's Regional
        # Roundup handling. Removed rather than left as unreachable dead code.
        if source_name in ("Chatham House", "Chatham House Events", "WEF Meetings", "Middle East Council Events", "MEC Daily Brief Wire Story", "Atlantic Council Highlights"):
            html = _fetch_html_via_browser(url)
        else:
            html = fetch_url(url)
    except Exception:
        return ""

    # 2026-10-22, real bug found: a Cloudflare-style bot-check interstitial (e.g. apnews.com) was
    # being returned by the browser path above and then extracted as if it were the real article -
    # its own generic "security verification" boilerplate ended up stored as Full Text (confirmed
    # live, consistently exactly 242 characters, no two different real articles coincidentally
    # producing the same page length). That's worse than failing outright, since the caller's
    # existing blurb-fallback never kicks in - caught here, before any extraction heuristic runs,
    # by checking for the challenge page's own fixed wording rather than a length/heuristic guess.
    if "performing security verification" in html.lower() or "verifies you are not a bot" in html.lower():
        return ""
    soup = make_soup(html)

    if source_name == "Lowy Institute":
        text = _lowy_full_text(soup)
        if len(text) >= _MIN_VIABLE_FULL_TEXT_CHARS:
            return text

    if source_name == "Deloitte Insights":
        text = _deloitte_full_text(soup)
        if len(text) >= _MIN_VIABLE_FULL_TEXT_CHARS:
            return text

    for selector in FULL_TEXT_SELECTORS.get(source_name) or []:
        container = soup.select_one(selector)
        if container is None:
            continue
        text = "\n\n".join(_extract_paragraphs(container, min_paragraph_chars))
        if len(text) >= _MIN_VIABLE_FULL_TEXT_CHARS:
            return text

    container = soup.find("article")
    if container is not None:
        text = "\n\n".join(_extract_paragraphs(container, min_paragraph_chars))
        if len(text) >= _MIN_VIABLE_FULL_TEXT_CHARS:
            return text

    # Generic heuristic: the div/section whose direct <p> children have the most text.
    best, best_len = None, 0
    for candidate in soup.find_all(["div", "section"]):
        paragraphs = candidate.find_all("p", recursive=False)
        total = sum(len(p.get_text(strip=True)) for p in paragraphs)
        if total > best_len:
            best, best_len = candidate, total
    if best is not None:
        text = "\n\n".join(_extract_paragraphs(best, min_paragraph_chars))
        if len(text) >= _MIN_VIABLE_FULL_TEXT_CHARS:
            return text

    return ""


PARSERS = {
    "rss": extract_rss,
    "cfr": extract_cfr,
    "cfr_reports": extract_cfr_reports,
    "csis": extract_csis,
    "chatham": extract_chatham_via_browser,
    "chinamacro": extract_chinamacro,
    "ac_dispatches": extract_ac_dispatches,
    "deloitte_insights": extract_deloitte_insights,
    "wef_publications": extract_wef_publications,
    "generic": extract_generic,
    "csis_events": extract_csis_events,
    "ac_events": extract_ac_events,
    "cfr_events": extract_cfr_events,
    "chatham_events": extract_chatham_events,
    "wef_meetings": extract_wef_meetings,
    "mec_events": extract_mec_events,
    "chinamacro_webinars": extract_chinamacro_webinars,
}

# Parsers that take (url, name) directly and do their own fetching internally, rather than
# receiving pre-fetched HTML - "rss" parses a feed, "deloitte_insights" parses a sitemap and
# then visits each candidate article page itself, "wef_publications" and "chatham" each drive
# their own Playwright browser instead of the plain fetch_url() helper (both sites block plain
# HTTP clients but not real browsers), "chinamacro" visits several listing pages (not just the
# one configured URL) and then each candidate's own page for its real JSON-LD date.
# "csis_events"/"ac_events" visit two URL variants (upcoming + past) themselves, same reasoning
# as chinamacro. "chatham_events"/"wef_meetings" need Playwright like their sibling regular-
# content parsers above. "mec_events" needs Playwright AND a tab click (its events page renders
# nothing server-side and its own event permalinks redirect-loop for a plain HTTP client).
# "chinamacro_webinars" visits its own 2 known real listing pages (plain HTTP, no WAF block
# found) rather than the configured hub page, same "ignore the configured URL, visit known real
# pages instead" shape as "chinamacro" and "csis_events"/"ac_events".
_SELF_FETCHING_PARSERS = {
    "rss", "deloitte_insights", "wef_publications", "chatham", "chinamacro", "chinamacro_webinars",
    "csis_events", "ac_events", "chatham_events", "wef_meetings", "mec_events",
}


def fetch_source(source: dict):
    """Dispatch one configured source to its extractor and return normalized articles."""
    parser = source.get("parser", "generic")
    url, name = source["url"], source["name"]
    if parser in _SELF_FETCHING_PARSERS:
        return PARSERS[parser](url, name)
    html = fetch_url(url)
    extractor = PARSERS.get(parser, extract_generic)
    if extractor is extract_generic:
        return extract_generic(html, url, name, source.get("path_patterns"))
    return extractor(html, url, name)
