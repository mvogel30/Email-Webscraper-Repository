"""Generic parser for a vendor with no dedicated parser module of its own yet (currently Middle
East Council on Global Affairs and Vriens & Partners - both configured with a best-guess domain
in vendors.json, verified correct via web search but not yet confirmed by an email actually
received in this mailbox).

Keeps each email whole as one article (same fallback behavior as
eurasia_parser.parse_eurasia_other_email for a format it doesn't specifically recognize) - no
section-splitting, no vendor-specific sign-off/boilerplate stripping, since only one real sample
exists so far (see detect_newsletter_name). Once more real emails arrive, if a vendor turns out to
have a distinct internal structure worth splitting (like Eurasia Group's multi-topic digest),
promote that vendor to its own parser module - this one is deliberately the safe minimum, not a
placeholder to leave in place forever for a high-volume sender.
"""

import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from email_common import (
    clean_text, clean_newsletter_body, strip_external_prefix, strip_unfamiliar_sender_banner,
    strip_subscribe_header, strip_email_preferences_footer, clean_title, clean_article_text,
    infer_publication_date, infer_geography, alpha2_from_geography, topic_from_geography,
    infer_topic, geography_from_label, compact_text,
    extract_event_date, extract_event_time, extract_registration_link,
)

# 2026-10-22, real bug found (per user audit request "make sure we are not capturing any false
# not wanted email like events or webinar announcements"): this module had NO event detection at
# all, unlike hub_advisor_events_parser.py's EVENT_SUBJECT_RE - a genuine, recurring MEC event
# invite ("Invitation to Attend: Protecting Women's Health in Fragile Settings" - real date, time,
# location, "Register Now" link, confirmed 16 times across saved exports) was being treated as a
# plain article and landing in Repository Rows instead of Events & Calls. Same keyword list as
# hub_advisor_events_parser.EVENT_SUBJECT_RE (proven against real CSIS/CFR/Atlantic Council event
# subjects) plus "Webinar"/"Panel"/"Save the Date", which MEC's own real invite uses ("a public
# panel") and the other three hadn't needed yet. Applies to every vendor this generic module
# serves (MEC today; Vriens & Partners/China Macro Group if either ever sends an event), not just
# MEC specifically - a reasonable general fix, not a one-off patch.
EVENT_SUBJECT_RE = re.compile(
    r"\b(Event|Watch Live|Live Now|Register Now|Invitation|Invite|Join us|RSVP|"
    r"Keynote|Fireside|Awards?|Honorees|POSTPONED|Webinar|Panel|Save the Date)\b",
    re.I,
)


def is_event_subject(subject: str) -> bool:
    return bool(EVENT_SUBJECT_RE.search(strip_external_prefix(clean_text(subject))))

# 2026-10-21: the regular MEC Daily Brief headlines link out to third-party wire services
# (Reuters, AP, etc.), not to mecouncil.org itself - reuses sources.py's own generic,
# no-dedicated-selector extraction path (the same one CFR already relies on) rather than
# building a second one from scratch. sources.py lives one directory up from this file.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import sources  # noqa: E402


# Middle East Council on Global Affairs: one real sample seen 2026-08-05 ("Daily Brief | August
# 5") - a pipe-separated date after "Daily Brief" is distinctive enough to check generically
# here rather than needing a whole dedicated parser module just for this one pattern. Vriens &
# Partners has no real sample yet at all, so falls through to the `vendor` name itself.
def detect_newsletter_name(subject: str, vendor: str) -> str:
    if re.search(r"\bDaily Brief\b\s*\|", subject or "", re.I):
        return "Daily Brief"
    # 2026-10-21, New Process & Scope: China Macro Group's one named newsletter - no real sample
    # seen yet (zero matches in a live 60-day scan of this inbox against their provisional,
    # unverified email domain - see vendors.new.json's _domain_note), so every email from them is
    # labeled this one fixed name, same "safe minimum" approach MEC/Vriens started with before
    # their first real sample arrived.
    if vendor == "China Macro Group":
        return "Main Newsletter"
    return vendor


def _strip_trailing_org_signature(text: str, vendor: str) -> str:
    """Cuts off a trailing address/signature block that starts with the vendor's own org name
    (e.g. "Middle East Council on Global Affairs Burj Al Mana, Zone 60, ..."), a common
    newsletter-platform footer pattern found 2026-08-05. Only trims when that name's LAST
    occurrence is near the very end of the text, so a genuine mid-article mention of the
    vendor's own name (not followed by much else) isn't mistaken for a footer."""
    if not vendor:
        return text
    idx = text.rfind(vendor)
    if idx == -1 or idx < len(text) - 500:
        return text
    return text[:idx].rstrip()


# 2026-09-23, per QA feedback: Middle East Council's "Daily Brief" was grouping its whole
# region-by-region roundup into one giant Full Text cell - too much in one cell. Split into one
# row per country section, same idea as Eurasia Group's "EG Daily Brief" splitter
# (eurasia_parser._daily_brief_sections), but MEC's real heading grammar (verified against 5 real
# captured emails spanning August-September) is genuinely different, so this is its own,
# independently-derived parser rather than a shared one:
#   - A few bylined feature teasers at the very top ("Title" line, then an "Author Name" line on
#     a later line, sometimes with a bare "Situation Assessment" label in between and/or an
#     orphan title with NO author at all - found in 2 of 5 real samples). These have no separate
#     body text anywhere in the plain-text email (the real analysis lives behind the link) - kept
#     as their own thin row (Full Text = "By <author>.") rather than dropped, since a real named
#     analyst byline is still real, distinct content.
#   - ALL-CAPS region headers ("GULF REGION", "MIDDLE EAST", "NORTH AFRICA") - pure grouping,
#     never become their own row.
#   - Title-case country headers ("Saudi Arabia", "Iran", "Palestine-Israel", "Türkiye") each
#     followed by one or more "Headline : blurb" lines - this is the real content, one row per
#     country (its Title is the country's first headline, Full Text joins every headline+blurb
#     for that country, same "derive from first item, keep the rest in the body" shape as EG's
#     bare-label sections).
_MEC_REGION_HEADER_RE = re.compile(r"^[A-Z][A-Z\s]{2,}[A-Z]$")
# Real bug found 2026-09-23: MEC's own template renders these inconsistently - "GULF REGION" in
# some emails, "Gulf Region" (title case) in others, for the SAME literal section. The all-caps
# structural check alone missed the title-case ones, which fell through to the front-matter
# handling below and then mis-read the next real country header ("Saudi Arabia" - two capitalized
# words, indistinguishable from an author name by that heuristic alone) as an author name,
# silently losing that whole country's real content. Checking this explicit, confirmed vocabulary
# case-insensitively first closes the gap; the structural check stays as a fallback for any other
# all-caps region name not yet seen in a real email.
_MEC_KNOWN_REGION_HEADERS = {"gulf region", "middle east", "north africa"}


def _is_mec_region_header(line: str) -> bool:
    if line.strip().lower() in _MEC_KNOWN_REGION_HEADERS:
        return True
    words = line.split()
    return len(words) >= 2 and bool(_MEC_REGION_HEADER_RE.match(line))


def _is_mec_country_header(line: str) -> bool:
    return ":" not in line and len(line) <= 40 and bool(geography_from_label(line))


_MEC_AUTHOR_NAME_RE = re.compile(r"^[A-ZÀ-ÖØ-Þ][\wÀ-ÖØ-öø-ÿ'.-]*(?:\s+[A-ZÀ-ÖØ-Þ][\wÀ-ÖØ-öø-ÿ'.-]*){1,3}$")


def _looks_like_mec_author_name(line: str) -> bool:
    if ":" in line or len(line) > 40 or any(ch.isdigit() for ch in line):
        return False
    # Defense in depth beyond the explicit region-vocabulary check above: a country/region name
    # ("Saudi Arabia", or any future region spelling not yet in _MEC_KNOWN_REGION_HEADERS) is
    # structurally indistinguishable from a real byline ("Khalid Al-Jaber") by word-count/
    # capitalization alone - reject anything geography_from_label recognizes before checking the
    # name-shape pattern.
    if geography_from_label(line):
        return False
    words = line.split()
    return 2 <= len(words) <= 4 and bool(_MEC_AUTHOR_NAME_RE.match(line))


_MEC_HEADLINE_RE = re.compile(r"^(.{5,200}?)\s*:\s*(.{15,})$")


def _mec_daily_brief_sections(body: str) -> list:
    """Returns a list of {"label", "title", "body"} dicts, one per top feature teaser or per
    HEADLINE within a country - same shape eurasia_parser's section dicts use.

    2026-10-01, per QA feedback ("only captures the first title of a particular category"): a
    country section with several distinct headlines used to become a SINGLE entry - only the
    first headline's title survived as the real Title, every later headline in that same country
    got silently folded into that first entry's body as plain "<title>: <blurb>" prose instead of
    becoming its own row. Confirmed live: several real country sections (e.g. multi-story days for
    Libya, Egypt) carry 2-4 distinct headlines, and only the first was ever visible as a Title.
    Now emits one entry per headline, all sharing the same country label (so Geography still
    applies to each), rather than merging everything after the first into one entry's body."""
    lines = [compact_text(line) for line in body.splitlines()]
    lines = [line for line in lines if line]

    sections = []
    current_label = None
    current_entries = None  # list of {"title", "body"} dicts for the country currently open
    pending_feature = None
    seen_region = False

    def flush():
        if current_label is not None:
            for entry in current_entries:
                sections.append({"label": current_label, "title": entry["title"], "body": entry["body"]})

    for line in lines:
        if _is_mec_region_header(line):
            flush()
            current_label, current_entries = None, None
            pending_feature = None
            seen_region = True
            continue
        if not seen_region:
            if line.lower() == "situation assessment":
                pending_feature = None
                continue
            if pending_feature is None:
                pending_feature = line
            elif _looks_like_mec_author_name(line):
                sections.append({"label": "", "title": pending_feature, "body": [f"By {line}."]})
                pending_feature = None
            else:
                pending_feature = line
            continue
        if _is_mec_country_header(line):
            flush()
            current_label, current_entries = line, []
            continue
        if current_entries is not None and current_label:
            # Real quirk found 2026-09-23 ("Libya" section, real captured 8/5 email): a headline
            # can redundantly repeat its own country header as a leading "<Country>: " prefix
            # ("Libya: Calm Returns to Sorman... : Libya's interim Government..."), which would
            # otherwise trick the non-greedy colon-split below into treating "Libya" itself as
            # the title. Stripped first, using the section's own known label - safe because it
            # only ever strips a prefix that exactly matches the country we're already inside.
            line = re.sub(rf"^{re.escape(current_label)}\s*:\s*", "", line, flags=re.I)
        headline_match = _MEC_HEADLINE_RE.match(line)
        if headline_match and current_entries is not None:
            title, blurb = (g.strip() for g in headline_match.groups())
            current_entries.append({"title": title, "body": [blurb]})
            continue
        # A wrapped continuation line of the previous blurb (no colon, not a header) - rare, but
        # seen in real bodies where the plaintext renderer splits one long sentence.
        if current_entries:
            current_entries[-1]["body"][-1] += " " + line
    flush()
    return sections


_MEC_LINK_TAG_RE = re.compile(r"<[a-z][a-z0-9+.\-]*://[^>]*>|<mailto:[^>]+>", re.I)


def _mec_teaser_links(text_with_links: str) -> dict:
    """Maps each line's own link-stripped, compact_text-normalized text to the first link tag
    found on that SAME raw line - used to recover a top feature teaser's real article link (see
    parse_mec_daily_brief below), since _MEC_LINK_TAG_RE.sub() runs before the splitter and
    discards links entirely otherwise. Keyed by TEXT rather than line position deliberately: the
    banner/subscribe/footer stripping steps that run between here and the splitter can shift or
    remove whole lines, which would break a position-based lookup, but they never alter the text
    of a line that survives, so a text-keyed lookup stays correct regardless.

    2026-10-21, per user feedback: this was only ever used to recover a TEASER's link (to fetch
    its real article text) - every other MEC row (every regular country-section headline) ended
    up with a permanently blank Link column, 751/751 real rows checked. That's not just a missing
    field: with no URL to dedupe on, a top-feature teaser MEC re-runs across several consecutive
    days (confirmed real: "Pakistan: A Patron of Last Resort for Middle East Security?" captured
    3 separate times, Sept 23/28/29, all byte-identical) looked like 3 different "new editions" to
    append_rows()'s date-based fallback instead of the same already-captured piece. Also keys by
    just the TITLE half of any headline-shaped line (via _MEC_HEADLINE_RE), not only the whole
    line, since a regular headline's title gets separated from its blurb by the splitter later and
    needs to be looked up by that shorter string."""
    links = {}
    for line in text_with_links.splitlines():
        match = _MEC_LINK_TAG_RE.search(line)
        if not match:
            continue
        link = match.group(0).strip("<>")
        stripped_line = compact_text(_MEC_LINK_TAG_RE.sub("", line))
        if not stripped_line:
            continue
        links.setdefault(stripped_line, link)
        headline_match = _MEC_HEADLINE_RE.match(stripped_line)
        if headline_match:
            title_part = compact_text(headline_match.group(1))
            if title_part:
                links.setdefault(title_part, link)
    return links


def _mec_article_full_text(link: str) -> str:
    """2026-09-24, per user feedback: a bylined top feature teaser's real analysis lives on
    mecouncil.org itself, not in the Daily Brief email body (see parse_mec_daily_brief's
    docstring) - visits that article page for its real text instead of leaving just "By <author>."
    as a stub. Confirmed live that MEC's article pages (unlike its EVENTS pages, which hit a real
    net::ERR_TOO_MANY_REDIRECTS even through a real browser - a genuine site-side limitation, not
    fixable here) render fine via Playwright, including navigating the email's own wrapped
    tracking link directly - the browser's own navigation follows the SafeLinks + click-tracking
    redirect chain transparently, no separate resolution step needed. Never raises - a fetch/parse
    failure just yields "", leaving the caller's existing stub text in place."""
    try:
        from playwright.sync_api import sync_playwright
        from bs4 import BeautifulSoup
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(link, timeout=30000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            html = page.content()
            browser.close()
        soup = BeautifulSoup(html, "html.parser")
        # 2026-09-24: MEC's site has (at least) two content templates for a linked article -
        # "/blog_posts/..." pages add ".single-blog-content" alongside this class, "/publication/
        # ..." pages don't - confirmed real on 3 of 28 real teaser links, all otherwise fetching
        # successfully. ".single-news-body" alone is the one class both templates share.
        container = soup.select_one(".single-news-body")
        if container is None:
            return ""
        return compact_text(container.get_text(" ", strip=True))
    except Exception:
        return ""


def parse_mec_daily_brief(subject: str, body: str, received_date: str) -> list:
    """Split an MEC Daily Brief into one row per top feature teaser / per country section.

    Deliberately does NOT go through clean_newsletter_body()/strip_links() for the initial pass -
    both call compact_text() internally, which collapses every newline into a single space. That
    destroys the line structure this splitter depends on (real bug found 2026-09-23: an early
    version produced 0-2 rows total across 37 real emails because every body had already been
    flattened into one giant line by the time the splitter ran). Strips link tags directly instead
    (newline-preserving), then the existing header/footer helpers below are all safe to run on
    multi-line text (none of them require a single-line string)."""
    text = clean_text(body)
    teaser_links = _mec_teaser_links(text)
    text = _MEC_LINK_TAG_RE.sub("", text)
    text = strip_unfamiliar_sender_banner(text)
    text = strip_subscribe_header(text)
    text = _strip_trailing_org_signature(text, "Middle East Council on Global Affairs")
    text = strip_email_preferences_footer(text)
    published = infer_publication_date(body, received_date)

    # First pass: everything that's pure text processing, no network I/O - fast.
    prepared = []
    for section_data in _mec_daily_brief_sections(text):
        title = clean_title(section_data["title"] or "")
        full_text = clean_article_text(" ".join(section_data["body"]))
        if not (title and full_text):
            continue
        # 2026-10-21, per user request: every row now fetches its OWN full article from its real
        # link, not just the two byline-stub teasers - a regular country-section headline's "full
        # text" used to just be the Daily Brief's own 1-3 sentence blurb, when the real article
        # sits one click away. Two different kinds of link need two different extractors: the two
        # top features are MEC's own writing on mecouncil.org (_mec_article_full_text's known
        # template); every other headline links to a third-party wire story (Reuters, AP, etc.) on
        # a site with no known template, so it goes through sources.py's generic, no-dedicated-
        # selector extraction path instead (same one CFR already relies on). Either way, falls
        # back to the blurb already in hand if the fetch fails - never raises, so a site hiccup or
        # paywall never loses content outright, it just keeps what the brief itself already gave.
        link = teaser_links.get(section_data["title"]) or teaser_links.get(compact_text(section_data["title"] or ""))
        is_teaser = full_text.startswith("By ") and full_text.endswith(".")
        prepared.append({"label": section_data["label"], "title": title, "full_text": full_text, "link": link, "is_teaser": is_teaser})

    # Second pass: the actual network fetches, run CONCURRENTLY rather than one at a time - a
    # typical Daily Brief has 15+ of these, almost entirely waiting on a different third-party
    # site each time, so running them one after another (confirmed live: ~100s per email) wastes
    # nearly all of that time doing nothing but waiting. Each fetcher already never raises
    # (returns "" on any failure - see _mec_article_full_text's and sources.fetch_full_text's own
    # docstrings), so no extra error handling is needed around the parallel calls themselves.
    def _fetch_one(item):
        if not item["link"]:
            return ""
        if item["is_teaser"]:
            return _mec_article_full_text(item["link"])
        # 2026-10-22, per user report - see the matching comment in sources.fetch_full_text():
        # routes through the same real-browser fetch as Chatham House/WEF now, instead of a plain
        # HTTP request most wire-service sites simply block.
        return sources.fetch_full_text(item["link"], "MEC Daily Brief Wire Story")

    with ThreadPoolExecutor(max_workers=8) as pool:
        fetched_texts = list(pool.map(_fetch_one, prepared))

    articles = []
    for item, fetched in zip(prepared, fetched_texts):
        title, full_text, link = item["title"], item["full_text"], item["link"]
        if fetched:
            full_text = fetched
        geography = geography_from_label(item["label"]) or infer_geography(f"{title} {full_text}")
        topic = topic_from_geography(geography) or infer_topic(f"{title} {full_text}")
        articles.append({
            "title": title,
            "url": link or "",
            "geography": geography,
            "geography_code": alpha2_from_geography(geography),
            "full_text": full_text,
            "topic": topic,
            "published": published.isoformat() if published else "",
        })
    return articles


def parse_generic_vendor_email(subject: str, body: str, received_date: str, vendor: str) -> dict:
    """Parse one email into a normalized article dict:
    {title, geography, geography_code, full_text, topic, published, newsletter_name}."""
    newsletter_name = detect_newsletter_name(subject, vendor)
    clean_subject = strip_external_prefix(clean_text(subject))
    title = clean_title(clean_subject)

    text = clean_newsletter_body(body)
    text = strip_unfamiliar_sender_banner(text)
    text = strip_subscribe_header(text)
    text = _strip_trailing_org_signature(text, vendor)
    text = strip_email_preferences_footer(text)
    full_text = clean_article_text(text)

    geography = infer_geography(f"{title} {full_text}")
    topic = topic_from_geography(geography) or infer_topic(f"{title} {full_text}")
    published = infer_publication_date(body, received_date)

    return {
        "title": title,
        "geography": geography,
        "geography_code": alpha2_from_geography(geography),
        "full_text": full_text,
        "topic": topic,
        "published": published.isoformat() if published else "",
        "newsletter_name": newsletter_name,
    }


def parse_email(subject: str, body: str, received_date: str, vendor: str = "",
                 sender_name: str = "", sender_email: str = "") -> list:
    """Uniform entry point every vendor parser module exposes (see
    parse_and_append_vendor_emails.py). MEC's "Daily Brief" specifically routes through the
    country-section splitter above; every other newsletter/vendor (including Vriens & Partners,
    which has no real sample yet at all, and MEC's own other newsletter types like "Monthly
    Newsletter" - unverified structure, left whole) keeps the original single-row behavior.
    `sender_name`/`sender_email` are unused here - accepted for interface consistency with the
    other parser modules, some of which do need them."""
    newsletter_name = detect_newsletter_name(subject, vendor)
    if vendor == "Middle East Council on Global Affairs" and newsletter_name == "Daily Brief":
        articles = parse_mec_daily_brief(subject, body, received_date)
        for article in articles:
            article["newsletter_name"] = newsletter_name
        return articles

    article = parse_generic_vendor_email(subject, body, received_date, vendor)
    if not (article.get("title") and article.get("full_text")):
        return []

    # 2026-10-22, real bug fix (see EVENT_SUBJECT_RE's module note above): route a genuine event/
    # webinar invitation to Events & Calls instead of Repository Rows, same is_event-flag
    # mechanism hub_advisor_events_parser.py's own events already use (parse_and_append_vendor_
    # emails.py routes purely off this flag, regardless of which parser module set it).
    if is_event_subject(subject):
        event_date = extract_event_date(subject, body, received_date)
        article.update({
            "newsletter_name": f"{vendor} Event",
            "is_event": True,
            "event_date": event_date.isoformat() if event_date else "",
            "announced_date": (received_date or "")[:10],
            "event_time": extract_event_time(subject, body),
            "registration_link": extract_registration_link(body),
        })
        return [article]

    return [article]
