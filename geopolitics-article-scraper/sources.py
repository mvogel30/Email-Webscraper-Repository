"""Fetch and normalize think-tank articles for a configured topic.

Mirrors the newsletter repository tool's shape: a thin *fetch* stage produces
raw normalized article records (written to JSON by fetch_thinktank_articles.py),
and a separate *append* stage de-duplicates them into the canonical workbook
(append_article_repository_workbook.py). Keeping fetch and append separate means
every run's raw articles are saved and the workbook is always rebuildable.

Each source in topics.json names a `parser`:
  - "rss"        : any RSS/Atom feed (Atlantic Council, Lowy Institute use this)
  - "cfr"        : Council on Foreign Relations topic page
  - "csis"       : CSIS region/topic page
  - "chatham"    : Chatham House region/topic page
  - "mecouncil"  : Middle East Council on Global Affairs "Publications" listing
  - "chinamacro" : China Macro Group "Insights" listing (Wix)
  - "generic"    : best-effort anchor scrape (fallback for a new site)

A normalized article is a plain dict:
  {source, title, url, published (ISO date str or ""), published_raw, summary, topic}
"""

import datetime as dt
import re
import time
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import requests
from bs4 import BeautifulSoup

try:
    import feedparser
except ImportError:  # surfaced with a clear message only if an RSS source is used
    feedparser = None


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 GeopoliticsArticleScraper/1.0"
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


# --------------------------------------------------------------------------- #
# Normalization
# --------------------------------------------------------------------------- #
def compact(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


_TRACKING_PARAMS = {
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "fbclid", "gclid", "mc_cid", "mc_eid",
}


def normalize_url(url: str) -> str:
    """Canonicalize a URL for stable de-duplication: lowercase scheme/host, drop a
    trailing slash, drop known tracking query params. Guards against the same
    article being re-added as "new" solely because of a cosmetic URL variation
    (this happened once: the same CSIS article showed up under two URLs)."""
    if not url:
        return url
    parts = urlsplit(url.strip())
    path = parts.path.rstrip("/") or "/"
    query_pairs = [(k, v) for k, v in parse_qsl(parts.query, keep_blank_values=True) if k.lower() not in _TRACKING_PARAMS]
    return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), path, urlencode(query_pairs), ""))


def make_article(source, title, url, published, summary, topic):
    return {
        "source": source,
        "title": compact(title),
        "url": normalize_url(url),
        "published": published.isoformat() if isinstance(published, dt.date) else "",
        "published_raw": published.strftime("%b %d, %Y") if isinstance(published, dt.date) else "",
        "summary": compact(summary),
        "topic": topic,
    }


def matches_keywords(article: dict, keywords) -> bool:
    """True if any keyword appears in title/summary/url (case-insensitive)."""
    if not keywords:
        return True
    haystack = f"{article['title']} {article['summary']} {article['url']}".lower()
    return any(str(k).lower() in haystack for k in keywords)


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
def extract_rss(url, source_name, topic):
    if feedparser is None:
        raise RuntimeError("feedparser is not installed (pip install feedparser) but source uses parser 'rss'")
    feed = feedparser.parse(fetch_url(url))
    articles = []
    for entry in feed.entries:
        published = None
        for attr in ("published_parsed", "updated_parsed"):
            parsed = getattr(entry, attr, None)
            if parsed:
                published = dt.date(parsed.tm_year, parsed.tm_mon, parsed.tm_mday)
                break
        summary = entry.get("summary", "") or ""
        summary = make_soup(summary).get_text(" ", strip=True)
        articles.append(make_article(source_name, entry.get("title", ""), entry.get("link", ""), published, summary, topic))
    return articles


# CFR's Iran page mixes several content types in identical h3 cards — not just "/articles/".
# The original selector only matched "/articles/" and silently dropped everything else (a real
# report, https://www.cfr.org/reports/conflict-driven-chokepoint-disruptions, was missed this
# way). Match any known CFR content type instead of just one.
_CFR_CONTENT_PATH_PREFIXES = ("/articles/", "/reports/", "/report/", "/event/")


def extract_cfr(html, base_url, source_name, topic):
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
        published = None
        if time_tag is not None:
            published = parse_date(_get_ci(time_tag, "datetime") or time_tag.get_text(strip=True))
        articles.append(make_article(source_name, anchor.get_text(" ", strip=True), url, published, "", topic))
    return articles


def extract_csis(html, base_url, source_name, topic):
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
        articles.append(make_article(source_name, anchor.get_text(" ", strip=True), url, published, summary, topic))
    return articles


def extract_chatham(html, base_url, source_name, topic):
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
        published = None
        if time_tag is not None:
            published = parse_date(_get_ci(time_tag, "datetime") or time_tag.get_text(strip=True))
        articles.append(make_article(source_name, title, url, published, summary, topic))
    return articles


def extract_mecouncil(html, base_url, source_name, topic):
    """Middle East Council on Global Affairs — the formal "Publications" listing
    (issue briefs/analyses). Chosen over their "Afkar" opinion-blog listing because
    this one carries a real per-article date; Afkar's listing has none."""
    soup = make_soup(html)
    articles, seen = [], set()
    for card in soup.select(".post-type-item"):
        anchor = card.select_one("a.post-type-img[href]") or card.select_one("a[href]")
        if anchor is None:
            continue
        url = urljoin(base_url, anchor.get("href", ""))
        if url in seen:
            continue
        seen.add(url)
        title_el = card.select_one(".meta-bottom .author") or card.select_one(".meta .author")
        title = title_el.get_text(" ", strip=True) if title_el else anchor.get_text(" ", strip=True)
        date_el = card.select_one(".publication-date .date")
        published = parse_date(date_el.get_text(strip=True)) if date_el else None
        articles.append(make_article(source_name, title, url, published, "", topic))
    return articles


def extract_chinamacro(html, base_url, source_name, topic):
    """China Macro Group — general "Insights" listing (Wix). Not Iran-specific
    content, so pair with a source-level require_keywords override in topics.json.
    The listing prefixes each summary with "Month Year | ..." — only month-level
    date precision is available (no day)."""
    soup = make_soup(html)
    articles, seen = [], set()
    for card in soup.select('[data-hook="post-list-item"]'):
        anchor = card.select_one('a[href*="/post/"]')
        if anchor is None:
            continue
        url = urljoin(base_url, anchor.get("href", ""))
        if url in seen:
            continue
        seen.add(url)
        title_el = card.select_one('[data-hook="post-title"]') or anchor
        title = title_el.get_text(" ", strip=True) if title_el else ""
        desc_el = card.select_one('[data-hook="post-description"]')
        summary = desc_el.get_text(" ", strip=True) if desc_el else ""
        published = None
        month_year = re.match(r"^([A-Za-z]+)\s+(\d{4})\s*\|\s*(.*)$", summary)
        if month_year and month_year.group(1).lower() in MONTHS:
            published = dt.date(int(month_year.group(2)), MONTHS[month_year.group(1).lower()], 1)
            summary = month_year.group(3)
        articles.append(make_article(source_name, title, url, published, summary, topic))
    return articles


def extract_generic(html, base_url, source_name, topic, path_patterns=None):
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
        articles.append(make_article(source_name, title, url, None, "", topic))
    return articles


# Per-source full-article-body container(s), verified against real article pages from each
# site. A source may need more than one selector because it uses different page templates for
# different content types (e.g. CSIS's "report" pages use a different layout than its regular
# commentary pages) — each is tried in order, first one yielding enough text wins. Falls back
# to the semantic <article> tag, then a generic "largest paragraph cluster" heuristic, then
# gives up (letting the caller fall back to the short listing/meta-description summary) rather
# than ever store a stub page's leftover tag-list text as if it were the real article.
FULL_TEXT_SELECTORS = {
    "Atlantic Council": [".ac-single-post--content"],
    "CSIS": [".article--layout-body", ".text-block.ts-friendly"],
    "Middle East Council on Global Affairs": [".single-news"],
    "Chatham House": [".body-content"],
    # Lowy Institute: handled specially in _lowy_full_text() — its body is split across
    # multiple div.prose blocks, and naively taking the first .prose only gets the one-line dek.
    # China Macro Group (Wix): rich content lazy-loads client-side after initial page load, so
    # only a short excerpt is ever present in the static HTML — a real ceiling for a plain HTTP
    # fetch, not a bug. [data-hook='post-page'] is still the right container, it's just short.
    "China Macro Group": ["[data-hook='post-page']"],
}

_MIN_VIABLE_FULL_TEXT_CHARS = 150

# Paragraphs that are real HTML <p> tags but are boilerplate, not article content — newsletter
# signup CTAs and "related experts" tag lists were seen bleeding into otherwise-good extractions.
_JUNK_PARAGRAPH_PREFIXES = (
    "sign up to receive", "sign up for", "subscribe to", "related experts:",
    "read more here", "share this", "reach out to discuss:",
)

# Paragraphs whose OWN class marks them as non-content (author byline blocks, etc.) rather than
# needing a text-prefix guess — more reliable when a site consistently tags these distinctly.
_JUNK_PARAGRAPH_CLASSES = {"author-teaser__meta"}


def _is_junk_paragraph(text: str) -> bool:
    lower = text.lower()
    return any(lower.startswith(prefix) for prefix in _JUNK_PARAGRAPH_PREFIXES)


def _extract_paragraphs(container, min_paragraph_chars: int):
    paragraphs = []
    for p in container.find_all("p"):
        if _JUNK_PARAGRAPH_CLASSES.intersection(p.get("class") or []):
            continue
        text = compact(p.get_text(" ", strip=True))
        if len(text) >= min_paragraph_chars and not _is_junk_paragraph(text):
            paragraphs.append(text)
    return paragraphs


def _lowy_full_text(soup) -> str:
    """Lowy's article body is split across one or more div.prose blocks (a multi-author
    round-up runs each author's take in its own block). The lede/dek and author-bio blocks
    carry extra utility classes; only the plain {"prose", "dark:prose-invert"} blocks are the
    real body — naively taking the page's first .prose returns just the one-sentence dek."""
    blocks = [el for el in soup.select(".prose") if set(el.get("class", [])) == {"prose", "dark:prose-invert"}]
    paragraphs = []
    for block in blocks:
        paragraphs.extend(_extract_paragraphs(block, min_paragraph_chars=40))
    return "\n\n".join(paragraphs)


def fetch_full_text(url: str, source_name: str, min_paragraph_chars: int = 40) -> str:
    """Best-effort FULL article body (distinct from the short listing/meta-description
    summary): visits the article's own page and joins the paragraph text from its known
    content container. Returns "" if nothing viable is found (caller falls back to the
    short summary) rather than store boilerplate/stub-page leftovers as "the article"."""
    try:
        soup = make_soup(fetch_url(url))
    except Exception:
        return ""

    if source_name == "Lowy Institute":
        text = _lowy_full_text(soup)
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


def fetch_article_summary(url: str) -> str:
    """Visit an article page and pull a short abstract: meta description /
    og:description, else the first substantial paragraph. Best-effort; returns
    "" on any failure so it never breaks a run."""
    try:
        soup = make_soup(fetch_url(url))
    except Exception:
        return ""
    for selector, attr in [
        ('meta[name="description"]', "content"),
        ('meta[property="og:description"]', "content"),
        ('meta[name="twitter:description"]', "content"),
    ]:
        tag = soup.select_one(selector)
        if tag and tag.get(attr):
            text = compact(tag.get(attr))
            if len(text) >= 40:
                return text
    for para in soup.find_all("p"):
        text = compact(para.get_text(" ", strip=True))
        if len(text) >= 80:
            return text
    return ""


PARSERS = {
    "rss": extract_rss,
    "cfr": extract_cfr,
    "csis": extract_csis,
    "chatham": extract_chatham,
    "mecouncil": extract_mecouncil,
    "chinamacro": extract_chinamacro,
    "generic": extract_generic,
}


def infer_parser(source: dict) -> str:
    url = source.get("url", "").lower()
    if url.endswith(("feed/", "/feed", ".xml", ".rss")) or "/feed" in url:
        return "rss"
    return "generic"


def fetch_source(source: dict, topic: str):
    """Dispatch one configured source to its extractor and return normalized articles."""
    parser = source.get("parser") or infer_parser(source)
    url, name = source["url"], source["name"]
    if parser == "rss":
        return extract_rss(url, name, topic)
    html = fetch_url(url)
    extractor = PARSERS.get(parser, extract_generic)
    if extractor is extract_generic:
        return extract_generic(html, url, name, topic, source.get("path_patterns"))
    return extractor(html, url, name, topic)
