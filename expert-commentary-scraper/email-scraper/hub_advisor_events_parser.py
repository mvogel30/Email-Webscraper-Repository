"""Events/calls-only parser for advisors whose regular commentary is ALREADY captured by the web
scraper (sources.py) - CSIS, Atlantic Council, CFR. A 2026-09-09 mailbox scan (all folders,
including Junk, unbounded date range) found real emails from all three going back to their
original subscription in June 2026, most sitting unseen in Junk Email. Checked directly against
each org's live scraped hub page: some individual pieces do overlap (a Mona Yacoubian CSIS piece,
one Michael Froman CFR piece), but several whole series structurally never appear there at all -
Atlantic Council's "Inflection Points" column, CSIS's "Capital Cable" video series and Critical
Minerals program, and CFR's program-specific newsletters (Think Global Health, Religion and
Foreign Policy, Asia Studies, Geoeconomics). Per explicit user decision, capturing that regular
newsletter content is deliberately OUT of scope for now - only genuine event/webinar/call
invitations are kept here; everything else from these three senders is dropped.

This is why parse_email() below returns [] for anything that doesn't look like an event, rather
than keeping it with some other categorization - unlike signum_parser.py/eurasia_parser.py, which
keep everything and only ROUTE event-shaped mail differently, this module keeps ONLY event-shaped
mail. Revisit this module (or replace it with a real per-org parser) if the user later decides to
also capture the regular newsletter content - see EXPERT-COMMENTARY-SCRAPER-HANDOFF.md and this
tool's own memory doc for that discussion.

Detection is subject-only (no body reads needed to tell an invite from a real article) - built and
verified against 189 real captured emails from these three senders, not guessed. Real event
subjects observed: "CSIS Event Invite: AUKUS at Five...", "Watch Live Now: Trump-Kim Summit
2.0?...", "Register Now: CSIS' The Future of U.S-Africa Minerals Diplomacy Event", "New Event: A
Conversation with H.E. Yassine Jaber...", "POSTPONED: A Conversation with H.E. Yassine Jaber...",
"CSIS Keynote Fireside with U.S. Energy Secretary Chris Wright", "Live Now: MENA's Economic
Outlook...", "CSIS Event Reminder: Was This World Cup More Political?...", "Join us virtually at
the Atlantic Council Global Citizen Awards | September 23", "Announcing our newest Atlantic
Council Global Citizen Awards Honorees". No CFR event-shaped subject has been seen yet in the real
corpus (CFR is still registered as a vendor here so a future one is caught automatically, same
reasoning as pre-registering Middle East Council/Vriens & Partners in generic_vendor_parser.py
before their first real sample ever arrived).

Known limitation, accepted for now: matching is necessarily incomplete (e.g. a bare
"This Monday: Lebanon's Strategic Outlook..." reminder for an already-flagged event carries none
of the below keywords and would be missed) - a false NEGATIVE here just means it's silently
dropped, same as today's behavior, not a regression. Revisit EVENT_SUBJECT_RE as more real
examples accumulate, same pattern as SIGNUM Event's catch-all being widened over time.
"""

import re
import sys
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from email_common import (
    clean_text, strip_external_prefix, clean_title, clean_article_text,
    clean_newsletter_body, strip_unfamiliar_sender_banner, infer_publication_date, extract_event_date,
    extract_event_time, extract_registration_link,
    infer_geography, alpha2_from_geography, topic_from_geography, infer_topic, compact_text,
    geography_from_label,
)

# 2026-10-22, per user request: Inflection Points' per-item links under its trailing "Highlights
# from the Atlantic Council" list need their own full-article fetch - same generic, no-dedicated-
# selector pattern MEC's wire-story links already use (see generic_vendor_parser.py), so sources.py
# is reused rather than duplicating the fetch/browser logic here.
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import sources  # noqa: E402


EVENT_SUBJECT_RE = re.compile(
    r"\b(Event|Watch Live|Live Now|Register Now|Invitation|Invite|Join us|RSVP|"
    r"Keynote|Fireside|Awards?|Honorees|POSTPONED)\b",
    re.I,
)


def is_event_subject(subject: str) -> bool:
    return bool(EVENT_SUBJECT_RE.search(strip_external_prefix(clean_text(subject))))


# 2026-10-21, New Process & Scope (see C:\Users\mivogel\.claude\plans\lively-roaming-moth.md): per
# user decision, ahead of a handover to the research department, the regular newsletter content
# this module's own docstring above describes as "deliberately OUT of scope for now" is now IN
# scope for the sources listed here - but ONLY these named newsletters, not everything from the
# sender (unlike signum_parser.py/eurasia_parser.py's "keep everything" vendors). A vendor with no
# key here (Lowy Institute today) gets NO newsletter branch at all - behavior for it is completely
# unchanged, event-or-drop exactly as before. These names are a best-effort guess (the research
# team isn't subscribed yet, and no real sample has been confirmed from this inbox either as of
# this writing) - refine the match once a real subject line is seen, same lifecycle EVENT_SUBJECT_RE
# itself went through above.
NEWSLETTER_NAMES_BY_VENDOR = {
    "CSIS": ["By the Numbers", "Trendlines", "Executive Insights"],
    "CFR": ["The Daily Brief", "The World This Week"],
    "Atlantic Council": ["AC Intel", "Inflection Points", "Fast Thinking", "The Week Ahead"],
}

# 2026-10-21: confirmed live against this inbox's real mail that two of the named newsletters
# above are NOT reliably identifiable from subject text at all - "AC Intel" sends under the
# sender display name "AC Intel" itself (subject is just the day's headline, e.g. "The missing
# piece of economic D-Day"), and "Inflection Points" sends from a dedicated
# inflectionpoints@atlanticcouncil.org address (Fred Kempe's column - subject rarely names it
# either). CFR's real "The Daily Brief" editions also mostly read as bare "Daily Brief: <topic>"
# (no "The"), which the subject-only rule below already catches via a substring match, but sender
# is included too as a more precise, format-proof signal. Checked BEFORE subject-text matching
# since it's the more reliable signal where it applies; vendors/newsletters not listed here simply
# fall through to the subject-text rule unchanged.
NEWSLETTER_SENDER_RULES_BY_VENDOR = {
    "Atlantic Council": [
        ("AC Intel", lambda name, email: name.strip().lower() == "ac intel"),
        ("Inflection Points", lambda name, email: email.strip().lower() == "inflectionpoints@atlanticcouncil.org"),
    ],
    "CFR": [
        ("The Daily Brief", lambda name, email: email.strip().lower() == "dailybrief@cfr.org"),
    ],
}


def _matched_newsletter_name(vendor: str, clean_subject: str, sender_name: str = "", sender_email: str = "") -> str:
    for name, matches in NEWSLETTER_SENDER_RULES_BY_VENDOR.get(vendor, []):
        if matches(sender_name or "", sender_email or ""):
            return name
    for name in NEWSLETTER_NAMES_BY_VENDOR.get(vendor, []):
        if name.lower() in clean_subject.lower():
            return name
    return ""


# 2026-10-22, per user request: "I see many Middle East Council articles that seem just to be
# headlines" led to a fix for MEC's wire-story links; the user then asked for the SAME idea here -
# Inflection Points' Fred Kempe essay is often followed by a short "Highlights from the Atlantic
# Council" list of 2-3 unrelated headlines, each linking to its own full Atlantic Council piece,
# with no blurb at all (unlike MEC's "Headline: blurb" format) - just a title and a link. Checked
# against 6 real captured editions (2026-10-22): only 3 of 6 have this section at all (the other 3
# are podcast-episode/event-announcement editions with no such list), and the section always ends
# with a "Support the Atlantic Council" donation link, which is never real content and must be
# excluded.
_AC_HIGHLIGHTS_HEADING_RE = re.compile(r"highlights from the atlantic council", re.I)
_AC_HIGHLIGHTS_STOP_RE = re.compile(r"support the atlantic council", re.I)
# Real formatting (verified against 3 real samples): each item starts with a literal "*" at the
# start of its own line - sometimes with the title+link on that SAME line, sometimes with the "*"
# alone and the title+link on the line(s) after it. Splitting on the bullet marker itself and then
# compacting each chunk's whitespace handles both shapes uniformly. Shared with CFR's Regional
# Roundup splitter below (both vendors use the same "*" bullet + inline "<url>" tag format).
_BULLET_SPLIT_RE = re.compile(r"(?m)^\*\s*")
_LINK_TAG_RE = re.compile(r"<[a-z][a-z0-9+.\-]*://[^>]*>|<mailto:[^>]+>", re.I)


def _strip_ac_highlights_section(body: str) -> str:
    """Trims the raw email body to everything BEFORE the "Highlights" heading, so the main essay
    row doesn't duplicate content that becomes its own rows below. Returns body unchanged if the
    section isn't present (most editions don't have it - see module note above)."""
    match = _AC_HIGHLIGHTS_HEADING_RE.search(body or "")
    return body[:match.start()] if match else body


def _extract_ac_highlights_items(body: str) -> list:
    """Returns [{"title", "link"}, ...] for each real item under "Highlights from the Atlantic
    Council", stopping before the "Support the Atlantic Council" donation link. Operates on
    clean_text(body) directly (NOT clean_newsletter_body/clean_article_text, which both collapse
    newlines via compact_text - see generic_vendor_parser.parse_mec_daily_brief's docstring for the
    same reasoning) since the bullet structure depends on real line breaks."""
    text = clean_text(body or "")
    heading_match = _AC_HIGHLIGHTS_HEADING_RE.search(text)
    if not heading_match:
        return []
    after = text[heading_match.end():]
    stop_match = _AC_HIGHLIGHTS_STOP_RE.search(after)
    block = after[:stop_match.start()] if stop_match else after

    items = []
    for chunk in _BULLET_SPLIT_RE.split(block):
        link_match = _LINK_TAG_RE.search(chunk)
        if not link_match:
            continue
        link = link_match.group(0).strip("<>")
        title = clean_title(compact_text(_LINK_TAG_RE.sub("", chunk)))
        if title:
            items.append({"title": title, "link": link})
    return items


def _fetch_ac_highlight_full_text(item: dict) -> str:
    if not item.get("link"):
        return ""
    return sources.fetch_full_text(item["link"], "Atlantic Council Highlights")


# 2026-10-22, per user request: AC Intel (unlike Inflection Points) keeps its own main row's Full
# Text completely untouched - no trimming - but every numbered item ending in a "READ <link>"
# action (as opposed to "WATCH"/"REGISTER"/etc., which are never articles) gets checked: if that
# link actually resolves to an atlanticcouncil.org page (a Dispatch or report - confirmed some
# "READ" links go to entirely different outlets instead, e.g. a Washington Times piece), it becomes
# its own "AC Intel (Linked Article)" row with the real fetched AC page as Full Text. A link that
# resolves anywhere else is dropped entirely - no row, per the user's own framing ("if you see...
# a read button that leads to an Atlantic Council page... then you take that").
#
# Verified against all 34 real captured editions before writing any of this: 175 real action items
# found, 97 of them "READ". Title/blurb are NOT numbered-item-shaped - verified via a real second
# example in the same inbox: a later "FUTURE INTEL" section item has the exact same title-blurb-
# action shape with no number at all. So extraction is anchored on the action paragraph itself,
# walking backward to the nearest real-text paragraph for the title (skipping over any image-only
# or bare-number paragraph in between - item 1 of every edition has an image banner paragraph
# between its title and blurb that a fixed-offset lookback would otherwise misread as the title).
_AC_INTEL_ACTION_RE = re.compile(r"^(READ|WATCH|REGISTER|LISTEN|EXPLORE|RSVP|JOIN|DISCOVER MORE|VIEW)$", re.I)
_AC_INTEL_NUMBER_ONLY_RE = re.compile(r"^\d+\.?$")


def _is_ac_intel_skippable_paragraph(compacted: str) -> bool:
    stripped = compact_text(_LINK_TAG_RE.sub("", compacted))
    return not stripped or bool(_AC_INTEL_NUMBER_ONLY_RE.match(stripped))


def _extract_ac_intel_read_items(body: str) -> list:
    """Returns [{"title", "link", "blurb"}, ...] for every "READ"-labeled item (not WATCH/
    REGISTER/etc.) - domain-checking happens later in _fetch_ac_intel_read_item, since that needs
    to actually resolve the link."""
    text = clean_text(body or "")
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    items = []
    for i, para in enumerate(paragraphs):
        no_link = compact_text(_LINK_TAG_RE.sub("", para))
        action_match = _AC_INTEL_ACTION_RE.match(no_link)
        if not action_match or action_match.group(1).upper() != "READ":
            continue
        link_match = _LINK_TAG_RE.search(para)
        if not link_match:
            continue
        blurb = compact_text(paragraphs[i - 1]) if i >= 1 else ""
        title = None
        j = i - 2
        while j >= 0:
            candidate = compact_text(paragraphs[j])
            if not _is_ac_intel_skippable_paragraph(candidate):
                title = candidate
                break
            j -= 1
        if not title or not blurb:
            continue
        items.append({"title": clean_title(title), "link": link_match.group(0).strip("<>"), "blurb": blurb})
    return items


def _fetch_ac_intel_read_item(link: str) -> dict:
    """Navigates the link for real and returns {"is_ac_domain", "full_text"} - full_text is only
    ever populated when is_ac_domain is True, since a non-AC destination is dropped entirely by
    the caller regardless of what text it might contain."""
    try:
        from playwright.sync_api import sync_playwright
        from bs4 import BeautifulSoup
        from urllib.parse import urlparse
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(link, timeout=30000, wait_until="domcontentloaded")
            page.wait_for_timeout(2000)
            final_url = page.url
            html = page.content()
            browser.close()
        host = urlparse(final_url).netloc.lower()
        if not (host == "atlanticcouncil.org" or host.endswith(".atlanticcouncil.org")):
            return {"is_ac_domain": False, "full_text": ""}
        if "performing security verification" in html.lower() or "verifies you are not a bot" in html.lower():
            return {"is_ac_domain": True, "full_text": ""}
        soup = BeautifulSoup(html, "html.parser")
        for selector in (".article-main", ".article-content"):
            container = soup.select_one(selector)
            if container is not None:
                text = compact_text(container.get_text(" ", strip=True))
                if text:
                    return {"is_ac_domain": True, "full_text": text}
        container = soup.find("article")
        text = compact_text(container.get_text(" ", strip=True)) if container else ""
        return {"is_ac_domain": True, "full_text": text}
    except Exception:
        return {"is_ac_domain": False, "full_text": ""}


# 2026-10-22, per user request: CFR's "The Daily Brief" is a 3-part structure - a lead essay (kept
# as today's single "CFR Daily Brief" row), a single featured "CFR Spotlight: <title>" item linking
# to a real cfr.org piece (fetched in full - cfr.org has no heavy bot-protection), and a "Regional
# Roundup" of short wire-service headlines grouped under region headers (Africa, Americas, Asia
# Pacific, Europe, Middle East - the only 5 seen across 32 real captured editions).
#
# Regional Roundup rows do NOT fetch the linked wire-service article (per user correction,
# 2026-10-22 - the earlier version did, which was both wrong per spec and slow: ~8-10 browser
# launches per edition just to hit mostly-blocked wire sites). Link is still kept so the real
# story is one click away, but Article Title is "<Region>; <bolded headline>" and Full Text is
# just the brief's own blurb (the line(s) right after the headline) - no network call at all for
# these rows now.
#
# Verified against all 32 real editions before writing any of this:
# - CFR Spotlight appears in ~1/3 of editions; Regional Roundup in all of them.
# - Each Regional Roundup bullet is "<Title ending in . or ! or ?> <blurb> <link>" (period-
#   separated, NOT MEC's colon-separated "Headline: blurb") - splitting naively on the first
#   sentence-ending punctuation mis-split real titles containing "U.S."/"U.K."/"U.N." (e.g. "Canada
#   Imposes New U.S." / "Tariffs. Retaliatory..." instead of the real, whole title) until an
#   explicit abbreviation guard was added and verified clean against all 171 real bullets collected.
_CFR_SPOTLIGHT_RE = re.compile(r"cfr spotlight\s*:", re.I)
_CFR_ROUNDUP_HEADING_RE = re.compile(r"regional roundup", re.I)
# "Manage Your Email Preferences" is the start of the standard footer that follows the roundup in
# every real sample checked - a reliable, already-used-elsewhere boilerplate marker (see
# strip_email_preferences_footer in email_common.py) rather than a CFR-specific guess.
_CFR_ROUNDUP_STOP_RE = re.compile(r"manage your email preferences", re.I)
_CFR_KNOWN_REGION_HEADERS = {"africa", "americas", "asia pacific", "europe", "middle east"}
# Abbreviation guard verified against all 171 real Regional Roundup bullets collected across 32
# editions (2026-10-22) - each alternative must be fixed-width for Python's re module, so common
# abbreviations are chained as separate lookbehinds rather than one alternation.
_CFR_TITLE_SPLIT_RE = re.compile(
    r"^(.{8,160}?(?<!\bU\.S)(?<!\bU\.K)(?<!\bU\.N)(?<!\bMr)(?<!\bDr)(?<!\bMs)(?<!Mrs)(?<!\bSt)"
    r"(?<!vs)(?<!etc)[.!?][\"'’”]?)\s+(?=[A-Z0-9À-ÖØ-Þ“\"])(.{15,})$",
    re.S,
)


def _strip_cfr_sections(body: str) -> str:
    """Trims the raw email body to everything BEFORE whichever of CFR Spotlight / Regional
    Roundup appears first, so the main essay row doesn't duplicate content that becomes its own
    rows below. Returns body unchanged if neither section is present."""
    matches = [m.start() for m in (_CFR_SPOTLIGHT_RE.search(body or ""), _CFR_ROUNDUP_HEADING_RE.search(body or "")) if m]
    return body[:min(matches)] if matches else body


def _extract_cfr_spotlight_item(body: str) -> dict:
    """Returns {"title", "link", "blurb"} for the single "CFR Spotlight: <title> <link>" item, or
    {} if this edition doesn't have one (most don't - see module note above). "blurb" is the
    brief's own paragraph about the feature (everything after the link up to the next section) -
    kept as a fallback for the rare case the real fetch fails, same reasoning as the Regional
    Roundup items below (cfr.org itself is rarely blocked, but "rarely" isn't "never")."""
    text = clean_text(body or "")
    heading_match = _CFR_SPOTLIGHT_RE.search(text)
    if not heading_match:
        return {}
    after = text[heading_match.end():]
    link_match = _LINK_TAG_RE.search(after)
    if not link_match:
        return {}
    title = clean_title(compact_text(after[:link_match.start()]))
    if not title:
        return {}
    rest = after[link_match.end():]
    stop_match = _CFR_ROUNDUP_HEADING_RE.search(rest)
    blurb = clean_article_text(rest[:stop_match.start()] if stop_match else rest[:1000])
    return {"title": title, "link": link_match.group(0).strip("<>"), "blurb": blurb}


def _is_cfr_region_header(line: str) -> bool:
    if line.strip().lower() in _CFR_KNOWN_REGION_HEADERS:
        return True
    # Defense in depth for a region not yet seen in the real corpus - same safety-net pattern
    # MEC's own region-header check uses.
    return len(line.split()) <= 3 and bool(geography_from_label(line))


def _extract_cfr_regional_roundup_items(body: str) -> list:
    """Returns [{"title", "link", "region", "blurb"}, ...], one per real headline - NOT one per
    region (see module note above: a real bug in MEC's own splitter used to lose every headline
    after the first in a multi-headline region, fixed 2026-10-01 - this is built with that lesson
    already applied, not repeating it). "blurb" is the brief's own 1-2 sentence description -
    kept as a fallback for when the real fetch fails (these are mostly wire-service links, often
    blocked - see _fetch_cfr_roundup_full_text), same reasoning as MEC's own fallback."""
    text = clean_text(body or "")
    heading_match = _CFR_ROUNDUP_HEADING_RE.search(text)
    if not heading_match:
        return []
    after = text[heading_match.end():]
    stop_match = _CFR_ROUNDUP_STOP_RE.search(after)
    block = after[:stop_match.start()] if stop_match else after

    lines = [line.strip() for line in block.splitlines()]
    lines = [line for line in lines if line]

    items = []
    current_region = None
    buffer = []

    def flush():
        if current_region is None or not buffer:
            return
        region_text = "\n".join(buffer)
        for chunk in _BULLET_SPLIT_RE.split(region_text):
            chunk = chunk.strip()
            link_match = _LINK_TAG_RE.search(chunk)
            if not link_match:
                continue
            link = link_match.group(0).strip("<>")
            chunk_no_link = compact_text(_LINK_TAG_RE.sub("", chunk))
            split_match = _CFR_TITLE_SPLIT_RE.match(chunk_no_link)
            if split_match:
                title, blurb = clean_title(split_match.group(1)), split_match.group(2)
            else:
                title, blurb = clean_title(chunk_no_link[:150]), ""
            if title:
                items.append({"title": title, "link": link, "region": current_region, "blurb": blurb})

    for line in lines:
        if _is_cfr_region_header(line):
            flush()
            current_region = line
            buffer = []
            continue
        if current_region is None:
            continue  # intro sentence before the first region header - not real content
        buffer.append(line)
    flush()
    return items


def _fetch_cfr_spotlight_full_text(link: str) -> str:
    # 2026-10-22: CFR Spotlight links point to cfr.org itself, which already has a working
    # generic <article> fallback for its regular website-scraped content (see FULL_TEXT_SELECTORS
    # in sources.py) - confirmed live that a plain request resolves the SafeLinks+link.cfr.org
    # tracking redirect chain fine (unlike Atlantic Council's HubSpot links above), so no browser
    # fetch is needed here.
    return sources.fetch_full_text(link, "CFR") if link else ""


def _build_base_article(clean_subject: str, body: str, received_date: str) -> dict:
    title = clean_title(clean_subject)
    text = clean_newsletter_body(body)
    text = strip_unfamiliar_sender_banner(text)
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
    }


def parse_email(subject: str, body: str, received_date: str, vendor: str = "",
                 sender_name: str = "", sender_email: str = "") -> list:
    """Uniform entry point every vendor parser module exposes (see
    parse_and_append_vendor_emails.py). Three-way branch: a real event invitation (unchanged
    behavior - see module docstring), one of this vendor's named newsletters (New Process & Scope
    only - matched by sender identity where that's the more reliable signal, else subject text -
    see NEWSLETTER_SENDER_RULES_BY_VENDOR/NEWSLETTER_NAMES_BY_VENDOR above), or neither, in which
    case the email is dropped exactly as it always has been for this module."""
    clean_subject = strip_external_prefix(clean_text(subject))

    if is_event_subject(subject):
        article = _build_base_article(clean_subject, body, received_date)
        # 2026-09-09, per user request: the real event date, not `published` above (an invite
        # sent ahead of the event falls back to "when this email arrived" - see
        # extract_event_date()'s docstring in email_common.py). Real subjects/bodies for these
        # three orgs consistently state it as "(Month Day)" in the subject or "on Weekday, Month
        # Day" in the body.
        event_date = extract_event_date(subject, body, received_date)
        article.update({
            "newsletter_name": f"{vendor} Event",
            "is_event": True,
            "event_date": event_date.isoformat() if event_date else "",
            "announced_date": (received_date or "")[:10],
            # 2026-09-11, per reviewer feedback: dedicated Event Time + a real registration link.
            "event_time": extract_event_time(subject, body),
            "registration_link": extract_registration_link(body),
        })
        return [article] if article["title"] and article["full_text"] else []

    newsletter_name = _matched_newsletter_name(vendor, clean_subject, sender_name, sender_email)
    if newsletter_name:
        # 2026-10-22, per user request: only Inflection Points ever carries the trailing
        # "Highlights from the Atlantic Council" list (see _extract_ac_highlights_items's module
        # note), and only CFR's "The Daily Brief" ever carries CFR Spotlight/Regional Roundup (see
        # _extract_cfr_spotlight_item/_extract_cfr_regional_roundup_items's module note) - every
        # other newsletter here keeps its original single-row behavior unchanged.
        highlight_items = _extract_ac_highlights_items(body) if newsletter_name == "Inflection Points" else []
        is_cfr_daily_brief = newsletter_name == "The Daily Brief" and vendor == "CFR"
        spotlight_item = _extract_cfr_spotlight_item(body) if is_cfr_daily_brief else {}
        roundup_items = _extract_cfr_regional_roundup_items(body) if is_cfr_daily_brief else []
        # 2026-10-22, per user request: AC Intel's own main row stays completely untouched (no
        # trimming - see _extract_ac_intel_read_items's module note) - only new rows get added.
        is_ac_intel = newsletter_name == "AC Intel" and vendor == "Atlantic Council"
        ac_intel_read_items = _extract_ac_intel_read_items(body) if is_ac_intel else []

        if highlight_items:
            main_body = _strip_ac_highlights_section(body)
        elif spotlight_item or roundup_items:
            main_body = _strip_cfr_sections(body)
        else:
            main_body = body
        article = _build_base_article(clean_subject, main_body, received_date)
        article.update({"newsletter_name": newsletter_name, "is_event": False})
        articles = [article] if article["title"] and article["full_text"] else []

        if highlight_items:
            with ThreadPoolExecutor(max_workers=4) as pool:
                fetched_texts = list(pool.map(_fetch_ac_highlight_full_text, highlight_items))
            for item, full_text in zip(highlight_items, fetched_texts):
                geography = infer_geography(f"{item['title']} {full_text}")
                topic = topic_from_geography(geography) or infer_topic(f"{item['title']} {full_text}")
                articles.append({
                    "title": item["title"],
                    "url": item["link"],
                    "geography": geography,
                    "geography_code": alpha2_from_geography(geography),
                    "full_text": full_text,
                    "topic": topic,
                    "published": article.get("published", ""),
                    "newsletter_name": "Inflection Points (Linked Article)",
                    "is_event": False,
                })

        if spotlight_item:
            full_text = _fetch_cfr_spotlight_full_text(spotlight_item["link"]) or spotlight_item.get("blurb", "")
            geography = infer_geography(f"{spotlight_item['title']} {full_text}")
            topic = topic_from_geography(geography) or infer_topic(f"{spotlight_item['title']} {full_text}")
            articles.append({
                "title": spotlight_item["title"],
                "url": spotlight_item["link"],
                "geography": geography,
                "geography_code": alpha2_from_geography(geography),
                "full_text": full_text,
                "topic": topic,
                "published": article.get("published", ""),
                "newsletter_name": "CFR Daily Brief (CFR Spotlight Section)",
                "is_event": False,
            })

        # 2026-10-22, per user correction: do NOT fetch the linked wire-service article for
        # Regional Roundup rows - keep the Link (so the real story is still one click away), but
        # Full Text is just the brief's own blurb (the line(s) right after the bolded headline),
        # nothing fetched from the newspaper itself. Also no longer a network call at all here, so
        # this whole section is now effectively instant instead of ~8-10 browser launches per
        # edition - directly fixes the slowness, not just worked around it.
        for item in roundup_items:
            full_text = item.get("blurb", "")
            geography = geography_from_label(item["region"]) or infer_geography(f"{item['title']} {full_text}")
            topic = topic_from_geography(geography) or infer_topic(f"{item['title']} {full_text}")
            articles.append({
                # 2026-10-22, per user request: Article Title is "<Region>; <bolded headline>",
                # not just the headline alone - the region prefix is explicit in the title text
                # itself, not only in the separate Geography column below.
                "title": f"{item['region']}; {item['title']}" if item.get("region") else item["title"],
                "url": item["link"],
                "geography": geography,
                "geography_code": alpha2_from_geography(geography),
                "full_text": full_text,
                "topic": topic,
                "published": article.get("published", ""),
                "newsletter_name": "CFR Daily Brief (Regional Roundup Section)",
                "is_event": False,
            })

        if ac_intel_read_items:
            with ThreadPoolExecutor(max_workers=4) as pool:
                fetch_results = list(pool.map(_fetch_ac_intel_read_item, (i["link"] for i in ac_intel_read_items)))
            for item, result in zip(ac_intel_read_items, fetch_results):
                if not result["is_ac_domain"]:
                    continue  # the READ link went somewhere other than atlanticcouncil.org - skip
                full_text = result["full_text"] or item.get("blurb", "")
                geography = infer_geography(f"{item['title']} {full_text}")
                topic = topic_from_geography(geography) or infer_topic(f"{item['title']} {full_text}")
                articles.append({
                    "title": item["title"],
                    "url": item["link"],
                    "geography": geography,
                    "geography_code": alpha2_from_geography(geography),
                    "full_text": full_text,
                    "topic": topic,
                    "published": article.get("published", ""),
                    "newsletter_name": "AC Intel (Linked Article)",
                    "is_event": False,
                })
        return articles

    return []
