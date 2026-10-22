"""Parser for Signum Global Advisors' Iran emails from Andrew Bishop
(andrew@signumglobal.com, subject contains "Iran:").

Adapted from the sibling newsletter-repository tool's backtest_newsletters.py (built by a
colleague for Outlook-based newsletter parsing), trimmed to just the Signum-specific logic —
this project doesn't need that tool's Eurasia/Bloomberg/Geopolitical Dispatch parsers.
"""

import datetime as dt
import re


def parse_signum_email(subject: str, body: str, received_date: str) -> dict:
    """Parse one Signum email into a normalized article dict matching this project's schema:
    {title, geography, geography_code, full_text, topic, published}."""
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
    # This tool only ever sees Andrew Bishop's Iran-subject emails (filtered at export time),
    # so default to "Iran" rather than leaving geography/topic blank when nothing else matches.
    geography = geography or infer_geography(f"{title} {text}") or "Iran"
    topic = infer_topic(title) or topic_from_geography(geography) or "Iran"

    published = infer_publication_date(body, received_date)

    return {
        "title": title,
        "geography": geography,
        "geography_code": alpha2_from_geography(geography),
        "full_text": text,
        "topic": topic,
        "published": published.isoformat() if published else "",
    }


# --------------------------------------------------------------------------- #
# Text cleaning (verbatim from the newsletter tool's helpers)
# --------------------------------------------------------------------------- #
def clean_text(value: str) -> str:
    return (
        (value or "")
        .replace("\r", "\n")
        .replace("\xa0", " ")
        .replace("͏", "")
        .replace("‌", "")
        .replace("‍", "")
        .replace("\xad", "")
    )


def compact_text(value: str) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def strip_links(value: str) -> str:
    value = re.sub(r"\s*<[a-z][a-z0-9+.\-]*://[^>]*>\s*", " ", value, flags=re.I)
    value = re.sub(r"\s*<mailto:[^>]+>\s*", " ", value, flags=re.I)
    value = re.sub(r"\b[a-z][a-z0-9+.\-]*://\S+", " ", value, flags=re.I)
    value = re.sub(r"\bmailto:\S+", " ", value, flags=re.I)
    return compact_text(value)


def strip_external_prefix(subject: str) -> str:
    return clean_text(subject).replace("[EXT] ", "", 1)


def clean_title(value: str) -> str:
    value = re.sub(r"\bRead in browser\b", " ", clean_text(value), flags=re.I)
    return re.sub(r"\s+Read more here.*$", "", compact_text(value), flags=re.I)


def clean_article_text(value: str) -> str:
    value = re.sub(r"\s+Read more here\s+<[^>]+>.*$", "", compact_text(value), flags=re.I)
    value = re.sub(r"\s+Reach out to discuss:.*$", "", value, flags=re.I)
    value = re.sub(r"\s*_{5,}\s*", " ", value)
    return strip_links(value).strip()


def clean_newsletter_body(value: str) -> str:
    value = re.sub(r"<https?://[^>]*\.(?:jpg|jpeg|png|gif)[^>]*>[^<]*?\bPhotographer:[^\n]*", " ", value, flags=re.I)
    value = compact_text(value)
    value = re.sub(r"^View in browser\s+<https?://[^>]+>\s*", "", value, flags=re.I)
    value = re.sub(r"^Read in browser\s+<https?://[^>]+>\s*", "", value, flags=re.I)
    value = re.sub(r"^(?:<https?://[^>]+>\s*)+", "", value, flags=re.I)
    value = re.sub(r"\s*<https?://[^>]+>\s*", " ", value, flags=re.I)
    value = re.sub(r"\s*<mailto:[^>]+>\s*", " ", value, flags=re.I)
    value = re.sub(r"\bView in browser\b", " ", value, flags=re.I)
    value = re.sub(r"\bRead in browser\b", " ", value, flags=re.I)
    value = re.sub(r"\bREAD IN APP\b", " ", value, flags=re.I)
    return compact_text(strip_links(value))


def clean_signum_body(value: str) -> str:
    text = clean_newsletter_body(value)
    # Outlook/Exchange's automatic "unfamiliar sender" safety banner — not real content, but
    # shows up whenever the sender isn't in the recipient's frequent-contacts list yet.
    text = re.sub(r"^You don.t often get email from .+?\.\s*Learn why this is important\s*", "", text, flags=re.I)
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


# --------------------------------------------------------------------------- #
# Date parsing
# --------------------------------------------------------------------------- #
MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7, "aug": 8,
    "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}
_MONTH_PATTERN = "|".join(sorted(MONTHS.keys(), key=len, reverse=True))


def month_index(month_name: str) -> int:
    return MONTHS[month_name.lower().rstrip(".")]


def infer_publication_date(body: str, received_date: str):
    try:
        received = dt.datetime.fromisoformat(received_date.replace("Z", "+00:00")).date()
    except (ValueError, AttributeError):
        received = None

    def near_received(candidate: dt.date) -> bool:
        return received is None or abs((candidate - received).days) <= 7

    for pattern in (
        rf"\b(\d{{1,2}})\s+({_MONTH_PATTERN})\.?\s+(\d{{4}})\b",
        rf"\b({_MONTH_PATTERN})\.?\s+(\d{{1,2}}),\s+(\d{{4}})\b",
    ):
        for match in re.finditer(pattern, body, re.I):
            groups = match.groups()
            try:
                if groups[0].isdigit():
                    candidate = dt.date(int(groups[2]), month_index(groups[1]), int(groups[0]))
                else:
                    candidate = dt.date(int(groups[2]), month_index(groups[0]), int(groups[1]))
            except ValueError:
                continue
            if near_received(candidate):
                return candidate

    return received


# --------------------------------------------------------------------------- #
# Geography / topic inference (trimmed COUNTRY_ALIASES table, same source as the web scraper)
# --------------------------------------------------------------------------- #
COUNTRY_ALIASES = {
    "US": "US", "USA": "US", "U.S.": "US", "UNITED STATES": "US",
    "UK": "UK", "U.K.": "UK", "UNITED KINGDOM": "UK", "BRITAIN": "UK",
    "CHINA": "China", "TAIWAN": "Taiwan", "IRAN": "Iran", "RUSSIA": "Russia", "UKRAINE": "Ukraine",
    "VENEZUELA": "Venezuela", "BRAZIL": "Brazil", "ISRAEL": "Israel", "JAPAN": "Japan",
    "EU": "Europe", "EUROPE": "Europe", "EUROPEAN UNION": "European Union", "MENA": "MENA",
    "SAUDI ARABIA": "Saudi Arabia", "UAE": "United Arab Emirates",
    "UNITED ARAB EMIRATES": "United Arab Emirates", "QATAR": "Qatar", "OMAN": "Oman",
    "KUWAIT": "Kuwait", "BAHRAIN": "Bahrain", "IRAQ": "Iraq", "SYRIA": "Syria",
    "LEBANON": "Lebanon", "JORDAN": "Jordan", "EGYPT": "Egypt", "TURKEY": "Turkey",
    "YEMEN": "Yemen", "THE GULF": "Gulf", "GULF": "Gulf", "RED SEA": "Red Sea",
    "PAKISTAN": "Pakistan", "AFGHANISTAN": "Afghanistan", "INDIA": "India",
    "NORTH KOREA": "North Korea", "SOUTH KOREA": "South Korea",
}


def geography_from_label(label: str) -> str:
    cleaned = re.sub(r"\([^)]*\)", " ", label).replace(".", "/")
    countries = []
    for part in re.split(r"\s*(?:/|,|;|\band\b|&|\+|-|–)\s*", cleaned, flags=re.I):
        part = compact_text(part).strip(" -:")
        if not part:
            continue
        canonical = COUNTRY_ALIASES.get(part.upper())
        if canonical and canonical not in countries:
            countries.append(canonical)
    return ", ".join(countries)


def infer_geography(text: str) -> str:
    candidates = list(COUNTRY_ALIASES.keys())
    found_positions = []
    for candidate in candidates:
        match = re.search(rf"\b{re.escape(candidate)}\b", text, re.I)
        if match:
            canonical = COUNTRY_ALIASES[candidate]
            if canonical not in [g for _, g in found_positions]:
                found_positions.append((match.start(), canonical))
    aliases = {
        "Israel": r"\bIsraeli\b|\bNetanyahu\b", "Russia": r"\bRussian\b|\bPutin\b|\bMoscow\b",
        "Ukraine": r"\bUkrainian\b|\bZelenskyy?\b|\bKyiv\b",
        "Gulf": r"\bGulf\b|\bQatar\b|\bBahrain\b|\bKuwait\b",
        "Red Sea": r"\bBab-el-Mandeb\b|\bHouthis?\b",
    }
    for geography, pattern in aliases.items():
        match = re.search(pattern, text, re.I)
        if match and geography not in [g for _, g in found_positions]:
            found_positions.append((match.start(), geography))
    found_positions.sort(key=lambda item: item[0])
    ordered = []
    for _, geography in found_positions:
        if geography not in ordered:
            ordered.append(geography)
    return ", ".join(ordered[:3])


ALPHA2_CODES = {
    "US": "US", "UK": "GB", "China": "CN", "Russia": "RU", "Ukraine": "UA", "Iran": "IR",
    "Taiwan": "TW", "Venezuela": "VE", "Brazil": "BR", "Israel": "IL", "Japan": "JP",
    "Europe": "EU", "European Union": "EU", "MENA": "MENA", "Saudi Arabia": "SA",
    "United Arab Emirates": "AE", "Qatar": "QA", "Oman": "OM", "Kuwait": "KW", "Bahrain": "BH",
    "Iraq": "IQ", "Syria": "SY", "Lebanon": "LB", "Jordan": "JO", "Egypt": "EG", "Turkey": "TR",
    "Yemen": "YE", "Gulf": "Gulf", "Red Sea": "Red Sea", "Pakistan": "PK", "Afghanistan": "AF",
    "India": "IN", "North Korea": "KP", "South Korea": "KR",
}


def alpha2_from_geography(geography: str) -> str:
    return ", ".join(ALPHA2_CODES[part.strip()] for part in geography.split(",") if part.strip() in ALPHA2_CODES)


REGIONAL_LABELS = {"Europe", "European Union", "MENA", "Gulf", "Red Sea"}
GEOGRAPHY_TOPIC_OVERRIDES = {"US": "US politics", "Russia": "Russia-Ukraine", "Ukraine": "Russia-Ukraine", "Iran": "Iran"}


def topic_from_geography(geography: str) -> str:
    primary = next((part.strip() for part in geography.split(",") if part.strip()), "")
    if not primary or primary in REGIONAL_LABELS:
        return ""
    return GEOGRAPHY_TOPIC_OVERRIDES.get(primary, f"{primary} politics")


def infer_topic(text: str) -> str:
    lower = text.lower()
    if "iran" in lower:
        return "Iran"
    if "russia" in lower or "ukraine" in lower:
        return "Russia-Ukraine"
    if "hormuz" in lower or "gulf" in lower:
        return "Gulf"
    return ""
