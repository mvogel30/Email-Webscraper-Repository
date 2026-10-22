"""Parser for Eurasia Group emails (any @eurasiagroup.net sender).

Adapted from the sibling newsletter-repository tool's backtest_newsletters.py (built by a
colleague for Outlook-based newsletter parsing) — specifically its parse_eurasia_daily_brief
section-splitting logic, trimmed to this project's needs. Generic text-cleaning, geography, and
date helpers are imported directly from signum_parser rather than duplicated.

Unlike Signum (filtered to Iran at export time via a reliable "Iran:" subject tag), Eurasia Group
doesn't tag Iran content in the subject line, and sends more than one email format ("EG Daily
Brief" plus others like "EG Update"). So export pulls every eurasiagroup.net email, and this
module decides Iran-relevance itself:
  - "EG Daily Brief" emails are split into their per-country/topic sections (same heading format
    the colleague's tool parses), and only sections that are actually about Iran are kept.
  - Any other Eurasia Group email is kept whole, as a single article, only if its subject or body
    mentions Iran.
"""

import re

import signum_parser as shared

clean_text = shared.clean_text
compact_text = shared.compact_text
strip_links = shared.strip_links
clean_title = shared.clean_title
clean_article_text = shared.clean_article_text
clean_newsletter_body = shared.clean_newsletter_body
infer_publication_date = shared.infer_publication_date
geography_from_label = shared.geography_from_label
infer_geography = shared.infer_geography
alpha2_from_geography = shared.alpha2_from_geography
topic_from_geography = shared.topic_from_geography
infer_topic = shared.infer_topic
COUNTRY_ALIASES = shared.COUNTRY_ALIASES


# Iran-relevance check: Eurasia Group doesn't subject-tag Iran content the way Andrew Bishop
# does for Signum, so this project has to decide it itself from subject+body text.
IRAN_KEYWORDS_RE = re.compile(
    r"\biran(?:ian)?\b|\btehran\b|\birgc\b|\bkhamenei\b|\bpezeshkian\b|\braisi\b"
    r"|\bquds force\b|\bhormuz\b",
    re.I,
)


def is_iran_relevant(*texts: str) -> bool:
    return any(IRAN_KEYWORDS_RE.search(text or "") for text in texts)


# Section-heading labels used by "EG Daily Brief" that aren't plain country names (topic/sector
# labels rather than geographies) — verbatim from the colleague's tool.
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
    article's footer — e.g. the 'Ian Bremmer on ...' President's-note items."""
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


def _daily_brief_sections(body: str) -> list:
    """Split an EG Daily Brief body into {label, title, body} sections, same heading grammar
    as the colleague's tool (label—title on one line, bullet body underneath)."""
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
            continue
        if lower_line in {"closing quip", "around the world"} or lower_line.startswith("explanation of political trajectories"):
            break
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
        if current and footer_seen and is_eurasia_feature_title(line):
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
    """Split an EG Daily Brief into its sections, keep only the Iran-relevant ones."""
    published = infer_publication_date(body, received_date)
    articles = []
    for section_data in _daily_brief_sections(clean_text(body)):
        title_text = section_data["title"]
        body_text = " ".join(section_data["body"])
        if not compact_text(body_text):
            title_text, body_text = split_title_from_body(title_text)
        full_text = clean_article_text(body_text)
        title = clean_title(title_text)
        if not (title and full_text):
            continue
        if not is_iran_relevant(section_data["label"], title, full_text):
            continue

        geography = geography_from_label(section_data["label"]) or infer_geography(f"{title} {full_text}") or "Iran"
        topic = topic_from_geography(geography) or infer_topic(f"{title} {full_text}") or "Iran"
        articles.append({
            "title": title,
            "geography": geography,
            "geography_code": alpha2_from_geography(geography),
            "full_text": full_text,
            "topic": topic,
            "published": published.isoformat() if published else "",
        })
    return articles


def parse_eurasia_other_email(subject: str, body: str, received_date: str) -> list:
    """Any non-Daily-Brief Eurasia Group email (e.g. "EG Update") — kept whole as one
    article, only if it's actually about Iran."""
    clean_subject = shared.strip_external_prefix(clean_text(subject))
    text = clean_newsletter_body(body)
    if not is_iran_relevant(clean_subject, text):
        return []

    title = clean_title(clean_subject)
    geography = infer_geography(f"{title} {text}") or "Iran"
    topic = topic_from_geography(geography) or infer_topic(f"{title} {text}") or "Iran"
    published = infer_publication_date(body, received_date)
    return [{
        "title": title,
        "geography": geography,
        "geography_code": alpha2_from_geography(geography),
        "full_text": clean_article_text(text),
        "topic": topic,
        "published": published.isoformat() if published else "",
    }]


def parse_eurasia_email(subject: str, body: str, received_date: str) -> list:
    """Parse one Eurasia Group email into zero or more normalized article dicts matching this
    project's schema: {title, geography, geography_code, full_text, topic, published}."""
    is_daily_brief = "daily brief" in (subject or "").lower()
    if is_daily_brief:
        return parse_eurasia_daily_brief(subject, body, received_date)
    return parse_eurasia_other_email(subject, body, received_date)
