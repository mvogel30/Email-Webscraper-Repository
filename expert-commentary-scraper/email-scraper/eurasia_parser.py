"""Parser for Eurasia Group emails (any @eurasiagroup.net sender, any subject).

Adapted from the sibling geopolitics-article-scraper tool's version of this file, which keeps
only Iran-relevant sections/emails. This tool has no topic filter, so every section of the
multi-topic digest ("EG Update" - see detect_newsletter_name) is kept (not just Iran-relevant
ones), and every other Eurasia Group email is kept whole - the per-country/topic section-splitting
itself is still valuable on its own (it turns one multi-topic email into several properly-tagged
rows), it just no longer gates on Iran-relevance.

Also tags every returned article with a "newsletter_name" - which named Eurasia Group format it
came from (e.g. "EG Update", "EG Conference Call", "EG Daily Brief") - see detect_newsletter_name().
"""

import re

import email_common as shared

clean_text = shared.clean_text
compact_text = shared.compact_text
strip_links = shared.strip_links
clean_title = shared.clean_title
clean_article_text = shared.clean_article_text
clean_newsletter_body = shared.clean_newsletter_body
infer_publication_date = shared.infer_publication_date
extract_event_date = shared.extract_event_date
extract_event_time = shared.extract_event_time
extract_registration_link = shared.extract_registration_link
geography_from_label = shared.geography_from_label
infer_geography = shared.infer_geography
alpha2_from_geography = shared.alpha2_from_geography
topic_from_geography = shared.topic_from_geography
infer_topic = shared.infer_topic
COUNTRY_ALIASES = shared.COUNTRY_ALIASES


# Section-heading labels used by "EG Daily Brief" that aren't plain country names (topic/sector
# labels rather than geographies).
EURASIA_LABEL_ALIASES = {
    "MENA CRISIS": "MENA Crisis",
    "OIL": "Oil",
    "GLOBAL MACRO": "Global Macro",
    "ENERGY": "Energy",
    "RESOURCES": "Resources",
    "GEO-TECHNOLOGY": "Geo-technology",
    "AEROSPACE & DEFENSE": "Aerospace & Defense",
    "AEROSPACE AND DEFENSE": "Aerospace & Defense",
    "EG UPDATE": "EG Update",
}


def is_valid_eurasia_label(label: str) -> bool:
    parts = [compact_text(part).strip(" -:") for part in re.split(r"\s*,\s*", label) if part.strip()]
    if not parts:
        return False
    return all(part.upper() in COUNTRY_ALIASES or part.upper() in EURASIA_LABEL_ALIASES for part in parts)


def is_plausible_eurasia_label(label: str) -> bool:
    label = compact_text(label).strip()
    if not (3 <= len(label) <= 55) or len(label.split()) > 7:
        return False
    if label.lower() in {"the top line", "summary", "upcoming conference calls", "upcoming events", "eg in depth"}:
        return False
    return True


# Fallback heading matcher for a country/region label not in COUNTRY_ALIASES/EURASIA_LABEL_ALIASES yet.
GENERIC_EURASIA_HEADING = re.compile(r"^(?P<label>[A-Z][A-Za-z.&,'’\- ]{1,55}?)\s*[—–]\s*(?P<title>.{8,})$")

EURASIA_FOOTER_RE = re.compile(r"reach out to discuss|read more here", re.I)


def is_eurasia_feature_title(line: str) -> bool:
    """A standalone, title-only feature heading (no 'Label—Title', no bullets) that follows an
    article's footer - e.g. the 'Ian Bremmer on ...' President's-note items."""
    line = compact_text(line)
    if line.startswith(("*", "•")) or "—" in line or "–" in line:
        return False
    if not (12 <= len(line) <= 110) or not line[:1].isupper() or len(line.split()) < 3:
        return False
    low = line.lower()
    blocked = (
        "reach out to discuss", "read more here", "upcoming", "closing quip", "around the world",
        "explanation of political", "manage your", "copyright", "compiled by", "pdf or print",
    )
    return not any(term in low for term in blocked)


def split_title_from_body(text: str) -> tuple:
    text = compact_text(text)
    match = re.search(r"(?<=[.!?])\s+(?=[A-Z])", text)
    if not match:
        return text[:240], text
    return text[: match.start()].strip(), text[match.end():].strip()


def _is_bare_label_line(line: str) -> bool:
    """A line that is JUST one or more label terms and nothing else (e.g. "iran", "us - china",
    "ukraine") - the real "EG Update" section-heading format, discovered 2026-08-05 from an
    actual captured email: sections there are a bare label alone on its own line, with the real
    content starting on later lines - NOT "LABEL - Title" on one line, which is a different
    format this parser also still supports (both are handled; this one takes priority when a
    line matches it, since e.g. "us - china" would otherwise misparse as label="us"/title="china"
    under the single-line pattern, silently losing the fact that it's a joint US-China section)."""
    normalized = re.sub(r"\s*[—–-]\s*", ",", line)
    return is_valid_eurasia_label(normalized)


# A copyright line ("(c) 2026 Eurasia Group. All rights reserved. ...") reliably marks the start
# of legal boilerplate at the very end of a real email - found 2026-08-05 in real captured
# bodies, where the existing footer/terminator markers below didn't match this email's actual
# sign-off style at all (no "reach out to discuss", no "closing quip"), so without this the
# entire copyright notice would otherwise get appended as trailing noise onto the last section.
_COPYRIGHT_LINE_RE = re.compile(r"^\s*[©c]\s*\d{4}\s+eurasia group", re.I)


def _daily_brief_sections(body: str) -> list:
    """Split an EG Daily Brief body into {label, title, body} sections, same heading grammar
    used across every EG Daily Brief."""
    lower = body.lower()
    top_line_index = lower.find("the top line")
    section = body[top_line_index + len("the top line"):] if top_line_index >= 0 else body

    for marker in ["upcoming events", "closing quip", "around the world", "explanation of political trajectories", "pdf or print"]:
        index = section.lower().find(marker)
        if index > 0:
            section = section[:index]

    labels = sorted(set(COUNTRY_ALIASES.keys()) | set(EURASIA_LABEL_ALIASES.keys()), key=len, reverse=True)
    label_terms = "|".join(re.escape(label) for label in labels)
    heading_pattern = re.compile(
        rf"^(?P<label>(?:{label_terms})(?:\s*,\s*(?:{label_terms}))*)\s*(?:[—–-]|\?)\s*(?P<title>.+)$",
        flags=re.I,
    )

    sections = []
    current = None
    footer_seen = False
    after_bullet_marker = False
    for line in section.splitlines():
        line = compact_text(line)
        if not line:
            continue
        lower_line = line.lower()
        if lower_line in {"upcoming conference calls", "upcoming events"}:
            if current:
                sections.append(current)
                current = None
            footer_seen = False
            after_bullet_marker = False
            continue
        if lower_line in {"closing quip", "around the world"} or lower_line.startswith("explanation of political trajectories"):
            break
        if _COPYRIGHT_LINE_RE.match(line):
            break
        if line in {"*", "•"}:
            # In the real "EG Daily Brief" format, every bullet point renders as a lone "*" on
            # its own line, with the bullet's actual text on the following line (an artifact of
            # how the sending platform lays out <li> markers in plain text) - found 2026-08-05
            # from real captured bodies. Earlier this line was treated as an end-of-section
            # signal, which was wrong: it fired on the FIRST bullet of every section, truncating
            # each section down to just its heading. Skip the bare marker instead so it doesn't
            # leak into the joined body text as a stray "*" prefix, without closing anything.
            after_bullet_marker = True
            continue
        # A line immediately following a bare bullet marker is ALWAYS that bullet's own body
        # text, never a real section heading - found 2026-08-17 from a real captured email whose
        # bullet text ("Government measures—including strategic buffer stockpiling...") happened
        # to contain an em dash for its own internal emphasis, which GENERIC_EURASIA_HEADING (a
        # fallback for a real but not-yet-aliased heading) matched anyway, splitting a genuine
        # section's bullet list into a bogus new "section" mid-list and losing the real content.
        # Real headings are always their own standalone line, never placed directly after a "*".
        is_bullet_body = after_bullet_marker
        after_bullet_marker = False
        if not is_bullet_body and _is_bare_label_line(line):
            if current:
                sections.append(current)
            current = {"label": line, "title": None, "body": []}
            footer_seen = False
            continue
        label = title_text = None
        if not is_bullet_body:
            match = heading_pattern.match(line)
            label = match.group("label") if (match and is_valid_eurasia_label(match.group("label"))) else None
            title_text = match.group("title") if label is not None else None
            if label is None:
                generic = GENERIC_EURASIA_HEADING.match(line)
                if generic and is_plausible_eurasia_label(generic.group("label")):
                    label = generic.group("label")
                    title_text = generic.group("title")
        if label is not None:
            if current:
                sections.append(current)
            title = title_text
            initial_body = []
            inline_match = re.match(r"(.{10,160}?):\s+(.{40,})$", title)
            if inline_match:
                title = inline_match.group(1)
                initial_body.append(inline_match.group(2))
            current = {"label": label, "title": title, "body": initial_body}
            footer_seen = False
            continue
        if not is_bullet_body and current and footer_seen and is_eurasia_feature_title(line):
            sections.append(current)
            current = {"label": "", "title": line, "body": []}
            footer_seen = False
            continue
        if current:
            current["body"].append(line)
            if EURASIA_FOOTER_RE.search(line):
                footer_seen = True
    if current:
        sections.append(current)
    return sections


def parse_eurasia_daily_brief(subject: str, body: str, received_date: str) -> list:
    """Split an EG Daily Brief into its sections - every section is kept (no topic filter)."""
    published = infer_publication_date(body, received_date)
    articles = []
    for section_data in _daily_brief_sections(clean_text(body)):
        title_text = section_data["title"]
        body_text = " ".join(section_data["body"])
        if not compact_text(title_text or ""):
            # Bare-label section (see _is_bare_label_line): no inline title, but real body
            # content - derive a title from the body's own first sentence instead of dropping
            # the section for having an empty title.
            title_text, body_text = split_title_from_body(body_text)
        elif not compact_text(body_text):
            # The other existing case: an inline "Label: longer body text" match put everything
            # into the title with no separate body - split that back apart.
            title_text, body_text = split_title_from_body(title_text)
        full_text = clean_article_text(body_text)
        title = clean_title(title_text)
        if not (title and full_text):
            continue

        geography = geography_from_label(section_data["label"]) or infer_geography(f"{title} {full_text}")
        topic = topic_from_geography(geography) or infer_topic(f"{title} {full_text}")
        articles.append({
            "title": title,
            "geography": geography,
            "geography_code": alpha2_from_geography(geography),
            "full_text": full_text,
            "topic": topic,
            "published": published.isoformat() if published else "",
        })
    return articles


# Prefixes stripped from the subject before it becomes the row's Title, for newsletter types
# whose Newsletter Name column would otherwise just repeat what's already in the subject (e.g.
# Title "EG Global Macro: G7 inaction..." when Newsletter Name already says "EG Global Macro").
# Scoped to only these known categories, not applied to the generic "Eurasia Group" bucket other
# practice-tagged notes (China, Energy & Climate, etc.) still fall into. A separate regex from
# _EG_TAG_PREFIX_RE above (that one's match spans the whole subject, for tag-content detection;
# this one matches only the prefix itself, for removal via .sub()).
_EG_TAG_PREFIX_STRIP_RE = re.compile(r"^(?:\[EXT\]\s*)?EG\s+[^:]+:\s*", re.I)
_EG_CONFERENCE_CALL_NOTES_PREFIX_RE = re.compile(r"^(?:\[EXT\]\s*)?EG\s+CONFERENCE CALL NOTES:\s*", re.I)
# Newsletter names that get their whole raw subject stripped down before it becomes the Title -
# these are the ones this module names directly by string (see detect_newsletter_name above);
# any OTHER "EG ..." name (from _eg_practice_newsletter_name's fallback, e.g. "EG China", "EG
# Energy, Climate & Resources") is handled by the generic startswith("EG ") branch below instead,
# using the same _EG_TAG_PREFIX_STRIP_RE - not by adding a new dict entry per practice area.
_NAMED_EG_TITLE_PREFIX_RE = {
    "EG Conference Call Notes": _EG_CONFERENCE_CALL_NOTES_PREFIX_RE,
}


def parse_eurasia_other_email(subject: str, body: str, received_date: str, newsletter_name: str = "") -> list:
    """Any non-Daily-Brief Eurasia Group email (e.g. "EG Update") - kept whole as one article."""
    clean_subject = shared.strip_external_prefix(clean_text(subject))
    text = clean_newsletter_body(body)
    text = shared.strip_unfamiliar_sender_banner(text)

    # Strip the redundant "EG <tags>: " (or "EG CONFERENCE CALL NOTES: ") prefix from the Title,
    # since Newsletter Name already carries that information - covers every practice-tagged note
    # (Global Macro, Geo-technology, and every other practice area named via
    # _eg_practice_newsletter_name), not just the two originally called out by name.
    prefix_re = _NAMED_EG_TITLE_PREFIX_RE.get(newsletter_name)
    if prefix_re is None and newsletter_name.startswith("EG "):
        prefix_re = _EG_TAG_PREFIX_STRIP_RE
    if prefix_re:
        stripped = prefix_re.sub("", clean_subject).strip()
        if stripped:
            clean_subject = stripped
    title = clean_title(clean_subject)
    geography = infer_geography(f"{title} {text}")
    topic = topic_from_geography(geography) or infer_topic(f"{title} {text}")
    published = infer_publication_date(body, received_date)
    return [{
        "title": title,
        "geography": geography,
        "geography_code": alpha2_from_geography(geography),
        "full_text": clean_article_text(text),
        "topic": topic,
        "published": published.isoformat() if published else "",
    }]


# Eurasia Group runs several distinctly-named/formatted email types, identified 2026-08-05 from
# real captured samples (not assumed). Two of them are rich multi-section digests that both use
# the same "the top line"/"eg in depth"/"LABEL—Title" heading grammar and both get split into one
# row per section (see parse_eurasia_email below):
#   - "EG Update": subject-tagged, casual "dear michael," format with bare labels ("iran",
#     "ukraine") alone on their own line.
#   - "EG Daily Brief": body-tagged (the subject is just that day's lead-story teaser), with
#     "Label—Title" headings (e.g. "Saudi Arabia, Oil—...", "Bolivia—..."), plus an "eg in depth"
#     block containing further inline "Label—Title: body" sub-topics. Initially misjudged (from
#     only seeing a short trailing snippet of two real samples) as a single-topic alert kept whole
#     - reading the FULL bodies on 2026-08-05 showed it's actually the richer of the two formats,
#     each with 6-9+ distinct topic sections that each need their own row.
# Practice-tag prefix used by Eurasia's standalone notes: "EG <tag1>[, <tag2>...] [QUICKTAKE]:
# <title>" (e.g. "EG Global Macro: G7 inaction...", "EG Geo-technology, China: Zhipu's gains...").
# Checked by tag content, not the whole subject, so a title that happens to mention one of these
# words later on doesn't false-match.
_EG_TAG_PREFIX_RE = re.compile(r"^(?:\[EXT\]\s*)?EG\s+([^:]+):\s*.+$", re.I)


def _eg_practice_tags(subject: str) -> str:
    match = _EG_TAG_PREFIX_RE.match((subject or "").strip())
    return match.group(1).lower() if match else ""


_EG_QUICKTAKE_SUFFIX_RE = re.compile(r"\s*QUICKTAKE\s*$", re.I)


def _eg_practice_newsletter_name(subject: str) -> str:
    """For any other standalone "EG <tag1>[, <tag2>...] [QUICKTAKE]: <title>" note not already
    named above (e.g. "EG China: ...", "EG Energy, Climate & Resources, China QUICKTAKE: ...") -
    2026-08-18, per the user: name it by its own practice tag(s) (e.g. "EG China", "EG Energy,
    Climate & Resources, China") instead of lumping every practice area into the generic "Eurasia
    Group" catch-all. Only called as the last resort in detect_newsletter_name, after every
    multi-section/named/invite format has already had first refusal, so it can never misclassify
    an EG Update/EG Daily Brief/EG Conference Call as a generic practice note."""
    match = _EG_TAG_PREFIX_RE.match((subject or "").strip())
    if not match:
        return ""
    tags = _EG_QUICKTAKE_SUFFIX_RE.sub("", match.group(1)).strip()
    return f"EG {tags}" if tags else ""


def detect_newsletter_name(subject: str, body: str) -> str:
    lower_subject = (subject or "").lower()
    if "eg update" in lower_subject:
        return "EG Update"
    # 2026-08-19: "Research Summary for <date>, 2026" is an older subject convention for the same
    # report as "Daily Research Summary - <date>" (found while investigating why some Eurasia
    # rows stayed generic) - just missing the word "Daily". Broadened from an exact
    # "daily research summary" match to "research summary" so both conventions map to the one
    # canonical name, rather than the older wording falling through to the generic bucket.
    if "research summary" in lower_subject:
        return "Daily Research Summary"
    # Checked BEFORE the invitation check below - found 2026-08-18 (manager asked for these
    # specifically) that real post-call summaries ("EG CONFERENCE CALL NOTES: <title>") were being
    # silently dropped entirely: "eg conference call" is a substring of "eg conference call notes",
    # so they were matching the invitation check further down and getting excluded as if they were
    # an invite, when they're actually the opposite - real content written AFTER a call happened.
    if "conference call notes" in lower_subject:
        return "EG Conference Call Notes"
    # "Political Risk Data – <Month> <Year>" - the monthly Political Risk Monitor release (named
    # "Political Risk Data" in its own subject, no "EG" prefix, unlike the practice notes below).
    if "political risk data" in lower_subject:
        return "EG Political Risk Monitor"
    tags = _eg_practice_tags(subject)
    if "global macro" in tags:
        return "EG Global Macro"
    if "geo-technology" in tags or "geotechnology" in tags:
        return "EG Geo-technology"
    if "eg conference call" in lower_subject:
        return "EG Conference Call"
    if re.search(r"\bdaily brief\b", body or "", re.I):
        return "EG Daily Brief"
    # Not every conference-call invite says "EG Conference Call" in the subject (e.g. "Analysing
    # global oil & gas markets" doesn't) - found 2026-08-05 falling through to the generic
    # fallback below instead. The body's own "Upcoming Conference Call" tag is the reliable
    # signal, same idea as "Daily Brief" being body-tagged rather than subject-tagged above.
    if re.search(r"\bupcoming conference call\b", body or "", re.I):
        return "EG Conference Call"
    # Last resort, checked only after every multi-section/named/invite format above has already
    # had first refusal - so this can never intercept an EG Update/EG Daily Brief/EG Conference
    # Call and misroute it away from multi-section splitting or the invite exclusion.
    practice_name = _eg_practice_newsletter_name(subject)
    if practice_name:
        return practice_name
    # 2026-08-19, per the user: label this bucket explicitly "- No Tag" rather than leaving it
    # looking identical to the vendor's own name - makes clear (confirmed via investigation) that
    # these are genuinely untagged in the source subject, not an unclassified gap on our end.
    return "Eurasia Group - No Tag"


def parse_eurasia_email(subject: str, body: str, received_date: str) -> list:
    """Parse one Eurasia Group email into zero or more normalized article dicts:
    {title, geography, geography_code, full_text, topic, published, newsletter_name}."""
    newsletter_name = detect_newsletter_name(subject, body)
    # Both "EG Update" and "EG Daily Brief" are rich multi-section digests using the same
    # heading grammar (see detect_newsletter_name's docstring) - each section becomes its own
    # row. Every other Eurasia Group format (EG Conference Call, Daily Research Summary, generic
    # admin threads) is kept as a single whole-email row.
    is_multi_section = newsletter_name in ("EG Update", "EG Daily Brief")
    articles = (
        parse_eurasia_daily_brief(subject, body, received_date) if is_multi_section
        else parse_eurasia_other_email(subject, body, received_date, newsletter_name)
    )
    for article in articles:
        article["newsletter_name"] = newsletter_name
    return articles


# 2026-08-05, per user request (same reasoning as Signum's SIGNUM Zoom/Investor Trip exclusion):
# these are invitations to join a call, not reporting or research content. Originally excluded
# entirely; 2026-09-02, per user request, re-added but routed to a separate "Events & Calls"
# sheet instead of the main Repository Rows sheet (see is_event below) - kept out of the main
# research feed, but no longer dropped outright. Distinct from "EG Conference Call Notes" (a real
# kept newsletter - actual notes from a call, not an invitation to one).
EVENT_NEWSLETTER_NAMES = {"EG Conference Call"}


def _trim_eg_conference_call_text(text: str) -> str:
    """2026-09-23, per QA feedback: EG Conference Call invites' plain-text body carries two real
    but non-event pieces of noise, confirmed present in every one of 35 real rows checked:
      1) A duplicated LEAD-IN - the HTML-to-text conversion emits a truncated preview snippet of
         the invite's own first sentence, cut off mid-word, before the real body starts fresh at
         its own "Upcoming Conference Call" marker (the same literal marker already relied on by
         extract_event_date() above). Confirmed real in all 35 rows, never at position 0.
      2) A trailing SPEAKER-BIO/JOIN-INSTRUCTIONS block - a "Speakers" heading (confirmed to occur
         exactly once in every row, never as an incidental word inside real content) followed by
         each speaker's full bio and, often, generic "How to Join" call-platform instructions.
         Real information about the analysts, but not part of the actual call description, and
         explicitly flagged by the reviewer as noise. Cutting at the same, single "Speakers" marker
         removes both in one step, since the bios and instructions were never usefully separable.
    Both cuts are no-ops (leave text untouched) if their marker isn't found, so a body that doesn't
    match this exact shape is never damaged."""
    marker = "Upcoming Conference Call"
    idx = text.find(marker)
    if idx > 0:
        text = text[idx:]
    speakers_idx = text.find("Speakers")
    if speakers_idx > 0:
        text = text[:speakers_idx].rstrip()
    return text


def parse_email(subject: str, body: str, received_date: str, vendor: str = "",
                 sender_name: str = "", sender_email: str = "") -> list:
    """Uniform entry point every vendor parser module exposes (see
    parse_and_append_vendor_emails.py). `vendor`/`sender_name`/`sender_email` are unused here
    (this module only ever serves Eurasia Group) - accepted for interface consistency with the
    other parser modules, some of which do need them."""
    articles = parse_eurasia_email(subject, body, received_date)
    is_event = bool(articles) and articles[0].get("newsletter_name") in EVENT_NEWSLETTER_NAMES
    for article in articles:
        article["is_event"] = is_event
        if is_event:
            # 2026-09-09, per user request: EG Conference Call invites only ever state the real
            # call date in the BODY, after the literal marker "Upcoming Conference Call" (e.g.
            # "Aug 24th, 2026") - `published` above is really just "when this email arrived" for
            # these, same issue as Signum (see extract_event_date()'s docstring).
            event_date = extract_event_date(subject, body, received_date)
            article["event_date"] = event_date.isoformat() if event_date else ""
            article["announced_date"] = (received_date or "")[:10]
            # 2026-09-11, per reviewer feedback: dedicated Event Time + a real registration link
            # (both new) - see both functions' docstrings in email_common.py.
            article["event_time"] = extract_event_time(subject, body)
            article["registration_link"] = extract_registration_link(body)
            article["full_text"] = _trim_eg_conference_call_text(article.get("full_text", ""))
    return [a for a in articles if a.get("title") and a.get("full_text")]
