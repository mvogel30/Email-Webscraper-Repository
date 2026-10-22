"""Parser for Signum Global Advisors emails (any @signumglobal.com sender, any subject).

Unlike the sibling geopolitics-article-scraper tool's version of this file - which only ever
sees Andrew Bishop's "Iran:"-tagged emails (filtered at export time) and so defaults geography/
topic to "Iran" when nothing else matches - this tool sees every Signum report on every subject,
so there is no safe default to fall back on: geography/topic are left blank when nothing is
inferred, rather than guessed.
"""

import re

from email_common import (
    clean_text, compact_text, strip_links, strip_external_prefix, clean_title, clean_article_text,
    clean_newsletter_body, strip_unfamiliar_sender_banner, infer_publication_date, extract_event_date,
    extract_event_time, extract_registration_link,
    geography_from_label, infer_geography, alpha2_from_geography, topic_from_geography, infer_topic,
)


# Signum runs a few distinctly-named recurring series alongside its default country/topic notes
# (e.g. "SIGNUM: Iran: ...", "SIGNUM: US: ..." - these don't have a more specific series name of
# their own beyond Signum itself) - identified 2026-08-05 from real captured subjects, not
# assumed. Checked against the ORIGINAL subject (before the generic "SIGNUM:\s*" prefix strip
# below), since e.g. "SIGNUM Zoom:" and "SIGNUM Investor Trip:" don't match that prefix pattern
# (no colon immediately after "SIGNUM") and would otherwise fall through to the title unchanged.
def detect_newsletter_name(subject: str) -> str:
    clean_subject = strip_external_prefix(clean_text(subject))
    if re.search(r"\bSIGNUM Investor Trip\b", clean_subject, re.I):
        return "SIGNUM Investor Trip"
    if re.search(r"\bSIGNUM Zoom\b", clean_subject, re.I):
        return "SIGNUM Zoom"
    # 2026-08-05, per user request (same reasoning as Zoom/Investor Trip): an in-person event
    # invitation, not research content. Location varies (e.g. "SIGNUM Paris Policy Day"), so
    # matched generically rather than pinned to one city - a real example also had a "NEW DATE:"
    # prefix from a reschedule, so this doesn't anchor to the start of the subject either.
    if re.search(r"\bSIGNUM\s+[A-Za-z]+\s+Policy Day\b", clean_subject, re.I):
        return "SIGNUM Policy Day"
    # 2026-08-07: general catch for in-person event invitations tied to a specific city/venue
    # (e.g. "SIGNUM New York: Hybrid Roundtable with..."), found after Zoom/Investor Trip/Policy
    # Day were already excluded one at a time - rather than adding a new named exception every
    # time a new event format shows up, this catches the shape they all share: "SIGNUM <word(s)>:"
    # with a colon appearing shortly after a location/descriptor word. Verified against every real
    # captured Signum subject to date: real reports always use "SIGNUM:" with the colon
    # IMMEDIATELY after (no word in between, e.g. "SIGNUM: Iran: ..."), and "Signum Weekly Digest
    # - <date>" has no colon anywhere - neither false-triggers this.
    if re.search(r"\bSIGNUM\s+[A-Za-z]+(?:\s+[A-Za-z]+)?\s*:", clean_subject, re.I):
        return "SIGNUM Event"
    if re.search(r"\bWeek ahead\b", clean_subject, re.I):
        return "Week Ahead"
    return "Signum"


def parse_signum_email(subject: str, body: str, received_date: str) -> dict:
    """Parse one Signum email into a normalized article dict:
    {title, geography, geography_code, full_text, topic, published, newsletter_name}."""
    newsletter_name = detect_newsletter_name(subject)
    subject = strip_external_prefix(clean_text(subject))
    title = clean_title(re.sub(r"^SIGNUM:\s*", "", subject, flags=re.I))
    geography = infer_geography(title)
    title_parts = [part.strip() for part in title.split(":") if part.strip()]
    if title_parts:
        geography = geography_from_label(title_parts[0]) or geography

    text = clean_signum_body(body)
    title_index = text.lower().find(title.lower())
    if title_index >= 0:
        text = text[title_index + len(title):].strip()
    text = clean_signum_article_text(text)
    geography = geography or infer_geography(f"{title} {text}")
    topic = infer_topic(title) or topic_from_geography(geography)

    published = infer_publication_date(body, received_date)

    return {
        "title": title,
        "geography": geography,
        "geography_code": alpha2_from_geography(geography),
        "full_text": text,
        "topic": topic,
        "published": published.isoformat() if published else "",
        "newsletter_name": newsletter_name,
    }


# 2026-08-05, per user request: these are invitations to attend a Signum-hosted call/trip/event,
# not reporting or research content. Originally excluded entirely; 2026-09-02, per user request,
# re-added but routed to a separate "Events & Calls" sheet instead of the main Repository Rows
# sheet (see EVENT_NEWSLETTER_NAMES check in parse_email() below) - kept out of the main research
# feed, but no longer dropped outright.
#
# 2026-09-11, per reviewer feedback: Events & Calls is meant for virtual/Zoom or hybrid events
# only, not purely in-person ones - "SIGNUM Investor Trip" (a multi-day physical trip) and
# "SIGNUM Policy Day" (a full in-person conference day) are never virtual/hybrid by their nature,
# so they're back to being fully excluded (dropped entirely, same as before 2026-09-02) rather
# than routed to Events & Calls. "SIGNUM Zoom" (always virtual) and the generic "SIGNUM Event"
# catch-all (covers hybrid ones too, e.g. "SIGNUM New York: Hybrid Roundtable...") still route to
# Events & Calls as before.
EVENT_NEWSLETTER_NAMES = {"SIGNUM Zoom", "SIGNUM Event"}
EXCLUDED_NEWSLETTER_NAMES = {"SIGNUM Investor Trip", "SIGNUM Policy Day"}


def parse_email(subject: str, body: str, received_date: str, vendor: str = "",
                 sender_name: str = "", sender_email: str = "") -> list:
    """Uniform entry point every vendor parser module exposes (see
    parse_and_append_vendor_emails.py): one email in, zero-or-more normalized article dicts out.
    `vendor`/`sender_name`/`sender_email` are unused here (this module only ever serves Signum,
    matched by domain alone) - accepted for interface consistency with the other parser modules,
    some of which do need them (hub_advisor_events_parser.py's sender-identity matching)."""
    article = parse_signum_email(subject, body, received_date)
    if article.get("newsletter_name") in EXCLUDED_NEWSLETTER_NAMES:
        return []
    article["is_event"] = article.get("newsletter_name") in EVENT_NEWSLETTER_NAMES
    if article["is_event"]:
        # 2026-09-09, per user request: Events & Calls needs the REAL event date (usually stated
        # right in the subject, e.g. "(Wednesday, September 3)"), not `published` above, which -
        # for an invite sent well before the event, the normal case - is really just "when this
        # email arrived" (see extract_event_date()'s docstring in email_common.py).
        event_date = extract_event_date(subject, body, received_date)
        article["event_date"] = event_date.isoformat() if event_date else ""
        article["announced_date"] = (received_date or "")[:10]
        # 2026-09-11, per reviewer feedback: a dedicated Event Time column and a real registration
        # link, neither of which existed before - see both functions' docstrings in
        # email_common.py for the real body shapes these were built and verified against.
        article["event_time"] = extract_event_time(subject, body)
        article["registration_link"] = extract_registration_link(body)
    return [article] if article.get("title") and article.get("full_text") else []


def clean_signum_body(value: str) -> str:
    text = clean_newsletter_body(value)
    text = strip_unfamiliar_sender_banner(text)
    text = re.sub(
        r"^(?:June|July|August|September|October|November|December|January|February|March|April|May)\s+\d{1,2},\s+\d{4}\s+",
        "",
        text,
        flags=re.I,
    )
    return text


def clean_signum_article_text(value: str) -> str:
    text = compact_text(value)
    text = re.split(
        r"\s+(?:Author:|Andrew D\. Bishop\b|Rafael Ch\b|Lew Lukens\b|Signum Global Advisors is a policy and strategy firm|This material is prepared by SIGNUM|Copyright\b|Update preferences\b|For client inquiries\b)",
        text,
        maxsplit=1,
        flags=re.I,
    )[0]
    # Analyst sign-off block, e.g. "Andrew Bishop Partner, Senior Analyst | Signum Global +44...".
    text = re.split(
        r"\s+[A-Z][A-Za-z.'’\-]+(?:\s+[A-Z][A-Za-z.'’\-]+){0,3}\s*,?\s*"
        r"(?:Partner|Senior|Director|Analyst|Founder|Co-Founder|Chief|Head|Managing Director|Vice President|Advisers?|Advisors?)"
        r"[^|]{0,100}\|\s*Signum Global",
        text,
        maxsplit=1,
        flags=re.I,
    )[0]
    text = re.split(r"\s+\|\s*Signum Global\b", text, maxsplit=1, flags=re.I)[0]
    text = re.split(r"\s+(?:Regards,|Best,|Sincerely,)\b", text, maxsplit=1, flags=re.I)[0]
    return clean_article_text(text)
