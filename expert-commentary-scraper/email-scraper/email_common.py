"""Shared text-cleaning, date-parsing, and geography-inference helpers used by every vendor
parser (signum_parser.py, eurasia_parser.py, and any future one). Adapted from the sibling
geopolitics-article-scraper tool's signum_parser.py, which itself adapted these from a
colleague's backtest_newsletters.py newsletter tool.

Kept as its own module (rather than living inside one vendor's parser file, as the sibling
Iran-scraper tool does) since more vendors are coming and none of this is Signum- or
Eurasia-specific.
"""

import datetime as dt
import re


# --------------------------------------------------------------------------- #
# Text cleaning
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
    # Found 2026-08-05: "Read more here" isn't always immediately followed by a link tag (e.g.
    # "Read more here and join our Iran update conference call today at 10 AM ET here.") - it's a
    # CTA phrase either way, so strip from it onward regardless of what follows, same as
    # clean_title() already does.
    value = re.sub(r"\s+Read more here\b.*$", "", compact_text(value), flags=re.I)
    value = re.sub(r"\s+Reach out to discuss:.*$", "", value, flags=re.I)
    value = re.sub(r"\s*_{5,}\s*", " ", value)
    value = strip_links(value).strip()
    # A lone trailing bullet marker left behind once the CTA after it is stripped (e.g. "...year's
    # end. *" once "Read more here..." is removed from after the "*").
    return re.sub(r"\s*\*\s*$", "", value)


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


# Outlook/Exchange's automatic "unfamiliar sender" safety banner - not real content, but shows
# up whenever the sender isn't in the recipient's frequent-contacts list yet. `.+?` (non-greedy)
# is deliberate: the sender's own email address contains a dot, so a greedy `[^.]+` breaks.
_UNFAMILIAR_SENDER_BANNER_RE = re.compile(r"^You don.t often get email from .+?\.\s*Learn why this is important\s*", re.I)


def strip_unfamiliar_sender_banner(value: str) -> str:
    return _UNFAMILIAR_SENDER_BANNER_RE.sub("", value)


# Found 2026-08-05 in a real Middle East Council email: a "August 5, 2026 | Subscribe |" nav
# tag at the very start of the body, and an address/unsubscribe block at the very end ("Would
# you like to change how you receive these emails? Please update your preferences or
# unsubscribe..."). Standard email-marketing-platform boilerplate, not specific to one vendor, so
# kept here rather than in one vendor's own parser.
_SUBSCRIBE_HEADER_RE = re.compile(r"^[A-Za-z]+\.?\s+\d{1,2},?\s+\d{4}\s*\|\s*Subscribe\s*\|\s*", re.I)
_EMAIL_PREFERENCES_FOOTER_RE = re.compile(r"\s*Would you like to change how you receive these emails\?.*$", re.I | re.S)


def strip_subscribe_header(value: str) -> str:
    return _SUBSCRIBE_HEADER_RE.sub("", value)


def strip_email_preferences_footer(value: str) -> str:
    return _EMAIL_PREFERENCES_FOOTER_RE.sub("", value)


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


_PORTAL_CHROME_RE = re.compile(r"Customize Email Preferences.*?Discover more on the Portal", re.I | re.S)


def infer_publication_date(body: str, received_date: str):
    try:
        received = dt.datetime.fromisoformat(received_date.replace("Z", "+00:00")).date()
    except (ValueError, AttributeError):
        received = None

    def near_received(candidate: dt.date) -> bool:
        return received is None or abs((candidate - received).days) <= 7

    # 2026-10-01, per QA feedback (real example: "US regulatory complacency allows financial
    # risks to build in AI boom" stored as Sept 25 when the email itself was received Sept 28):
    # Eurasia Group emails often embed an inline portal-preview widget ("Quick take"/"Practice
    # Letter" ... "Customize Email Preferences" ... <repeated title> <own dateline, e.g. "25 Sep
    # 2026 09:58 EDT"> "Discover more on the Portal") in the MIDDLE of the real content - confirmed
    # live on 189 real Eurasia Group rows. Its own dateline can differ from when THIS email was
    # actually sent/received (the portal timestamp reflects whenever the piece was first posted
    # there, not this particular distribution), so the date-scan below must not be able to match
    # inside it - stripped from a working copy first, body text passed to the caller is untouched.
    body = _PORTAL_CHROME_RE.sub(" ", body)

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


# 2026-09-09: event/call invitations (Signum Zoom/Event, Eurasia Conference Call, CSIS/Atlantic
# Council email invites) need the REAL event date, not infer_publication_date()'s notion of
# "publication date" (which - for an invite sent well ahead of the event, the normal case - just
# falls back to when the EMAIL was received, confirmed against real rows: e.g. an invite received
# Sep 3 for an event on Oct 14 was stored as Sep 3). Verified against real captured data before
# writing this:
#   - Signum: date is in the SUBJECT, in parens, with or without a weekday - "(Wednesday,
#     September 3)" / "(September 15)".
#   - CSIS/Atlantic Council (hub_advisor_events_parser): same subject shape - "(June 11)" - and
#     often ALSO restated in the body as "on Thursday, June 11 at 2:00 PM EDT".
#   - Eurasia Group Conference Call: date is ONLY in the body, after the literal marker "Upcoming
#     Conference Call", written WITH an ordinal suffix and year - "Aug 24th, 2026" - a shape
#     infer_publication_date()'s own regexes don't handle (they require a bare digit immediately
#     before the comma, not "24th,").
# None of the above give a year in the Signum/CSIS/AC case, so one is inferred from the received
# date (rolled forward a year if the resulting date would otherwise land more than a few days in
# the past - handles an invite sent in December for a January event).
#
# Two direction branches, not just "Month Day": a real Eurasia body reading "Monday, 27 July
# 09:30am US East" was initially mismatched as "July 09" (reading the TIME "09:30am" as if it
# were the day-of-month) - found by testing against real data, not assumed. Fixed two ways: (1)
# a "Day Month" branch handles this exact "27 July" order directly, and (2) a guard rejects a
# "Month Day" match whose day-digits are immediately followed by ":" or "am"/"pm" (a time, not a
# day) even where that order IS the real one elsewhere.
_WEEKDAY_PATTERN = r"(?:Mon|Tue(?:s)?|Wed(?:nes)?|Thu(?:rs)?|Fri|Sat(?:ur)?|Sun)(?:day)?"
_NOT_A_TIME = r"(?!\s*:)(?!\s*(?:am|pm)\b)"
_EVENT_DATE_RE = re.compile(
    rf"\b(?:on\s+)?(?:{_WEEKDAY_PATTERN},?\s+)?(?:"
    rf"({_MONTH_PATTERN})\.?\s+(\d{{1,2}})(?!\d){_NOT_A_TIME}(?:st|nd|rd|th)?(?:,?\s*(\d{{4}}))?"
    rf"|"
    rf"(\d{{1,2}})(?:st|nd|rd|th)?\s+({_MONTH_PATTERN})\.?(?!\d)(?:,?\s*(\d{{4}}))?"
    rf")\b",
    re.I,
)


def extract_event_date(subject: str, body: str, received_date: str):
    """Real event date for an event/call invitation - checks the SUBJECT first (covers Signum/
    CSIS/Atlantic Council's "(Month Day)" bracket convention), then the BODY (covers Eurasia
    Group, whose Conference Call invites only state the date there). Falls back to the email's
    own received date if no real date is found anywhere, same safe fallback as
    infer_publication_date()."""
    try:
        received = dt.datetime.fromisoformat((received_date or "").replace("Z", "+00:00")).date()
    except (ValueError, AttributeError):
        received = None

    assume_year = received.year if received else dt.date.today().year
    for text in (subject or "", body or ""):
        match = _EVENT_DATE_RE.search(text)
        if not match:
            continue
        month_day_groups, day_month_groups = match.groups()[:3], match.groups()[3:]
        if month_day_groups[0] is not None:
            month_str, day_str, year_str = month_day_groups
        else:
            day_str, month_str, year_str = day_month_groups
        try:
            month, day = month_index(month_str), int(day_str)
            year = int(year_str) if year_str else assume_year
            candidate = dt.date(year, month, day)
        except (ValueError, KeyError):
            continue
        if not year_str and received is not None and candidate < received - dt.timedelta(days=3):
            try:
                candidate = dt.date(year + 1, month, day)
            except ValueError:
                pass
        return candidate

    return received


# 2026-09-11, per reviewer feedback: Events & Calls needs a real "Event Time" column and a real
# registration link for email-sourced events, neither of which existed before. Built and verified
# against real bodies, not guessed:
#   - Signum: "Thursday August 6 11am ET / 4pm BST / 7pm GST / 11pm SGT" then "REGISTER <url>".
#   - Eurasia: "Register Here <url>" (no explicit time found in real bodies checked).
#   - CSIS/Atlantic Council (hub_advisor_events_parser): "Thursday, June 11, 2026" / "2:00 - 2:45
#     PM EDT" / "Online - Livestream" then "Register Here <url>".
# Timezone abbreviations are matched against an explicit whitelist, not bare letters - an earlier
# looser attempt falsely matched "21-24 Septe" out of a WEF date RANGE ("21-24 September 2026") as
# if it were a time.
_TZ_ABBREV = r"(?:ET|EDT|EST|PT|PDT|PST|CT|CDT|CST|MT|MDT|MST|GMT|BST|CET|CEST|UTC|SGT|GST|IST|AEST|AEDT|HKT|JST|KST)"
# 2026-09-23, per QA feedback: a real Atlantic Council event email's body used "7:00 p.m. ET" -
# periods in "p.m.", which the plain "am|pm" shape below didn't match, silently leaving that
# event's time blank (same underlying gap independently found in CFR's own event pages, which
# write "4:15 p.m. (EDT)" - a parenthesized zone too). _AMPM accepts an optional period after each
# letter; _TZ lets the abbreviation be optionally wrapped in parens.
_AMPM = r"(?:[ap]\.?m\.?)"
# 2026-10-22, real gap found (MEC's own event invite): "6:00 pm - 8:00 pm (GMT+3)" was being cut
# to "...(GMT" - the abbreviation-only pattern had no room for a numeric UTC offset suffix, which
# GMT/UTC-style zones commonly carry (unlike the US ET/PT-style abbreviations, which never do).
_TZ = rf"\(?{_TZ_ABBREV}(?:[+-]\d{{1,2}}(?::?\d{{2}})?)?\)?"
_TIME_SEP = r"(?:-|–|to)"
_EVENT_TIME_RE = re.compile(
    rf"\d{{1,2}}(?::\d{{2}})?\s*{_AMPM}?\s*{_TIME_SEP}\s*\d{{1,2}}(?::\d{{2}})?\s*{_AMPM}?\s*{_TZ}"
    rf"(?:\s*[/|]\s*\d{{1,2}}(?::\d{{2}})?\s*{_AMPM}?\s*{_TIME_SEP}?\s*\d{{0,2}}:?\d{{0,2}}\s*{_AMPM}?\s*{_TZ})*"
    rf"|\d{{1,2}}(?::\d{{2}})?\s*{_AMPM}\s*{_TZ}(?:\s*/\s*\d{{1,2}}(?::\d{{2}})?\s*{_AMPM}\s*{_TZ})*"
    # 2026-09-23, per QA feedback (Signum): a real single-city, in-person roundtable invite states
    # only "Wednesday, October 14 12:00 PM" - no timezone at all (unlike Signum's multi-timezone
    # webinar format the branch above was built for). Safe as its own fallback, unlike the bare
    # "N-M TZ" range above: colon-separated minutes + a mandatory explicit am/pm is a specific
    # enough shape that it doesn't need a timezone to be a plausible real time.
    rf"|\d{{1,2}}:\d{{2}}\s*{_AMPM}",
    re.I,
)


def extract_event_time(subject: str, body: str) -> str:
    """Real event time (with timezone) for an event/call invitation - checks the SUBJECT first,
    then the BODY. Returns "" (not None) when nothing is found, since this goes straight into a
    text cell, unlike extract_event_date()'s date return."""
    for text in (subject or "", body or ""):
        match = _EVENT_TIME_RE.search(text)
        if match:
            value = re.sub(r"\s+", " ", match.group(0)).strip()
            # 2026-09-23, per QA feedback: a real garbled cell ("3-0pt") traced back to this regex's
            # bare-range branch matching an ordinary "N-M" number pair immediately followed by a
            # real timezone abbreviation (here "PT"), with nothing else marking it as a time - the
            # same false-positive shape the module already fought once for WEF date ranges (see the
            # comment above _TZ_ABBREV). A genuine event time always has minutes (":00") or an
            # explicit am/pm somewhere (with or without periods); reject the match if it has
            # neither, rather than widening the regex itself and risking new false positives.
            if not re.search(r":|[ap]\.?m\.?", value, re.I):
                continue
            return value
    return ""


# Matches "Register Here <url>" / "REGISTER <url>" (Signum, no "Here") / "RSVP <url>" - confirmed
# 2026-09-11 that a Eurasia REMINDER email can use "RSVP" as its button text where the original
# invite for the SAME webinar used "Register Here" - both real, both need catching. The URL
# itself is always a tracking-wrapped link (Microsoft Safelinks, or the vendor's own click-
# tracking domain) unique per email send - real, clickable, but NOT stable across an invite's own
# "REMINDER" re-send, so this is deliberately kept OUT of dedup (see
# parse_and_append_vendor_emails.py's emails_to_articles() - it's stored as "registration_link",
# a separate field from "url", which is what dedup actually keys on).
# 2026-09-23, per QA feedback: a real Atlantic Council event email ("Meet the introducers...")
# came through with its Link blank - traced to its button text being "Register now" (also seen:
# "Register today"), which the "here"-only variant above doesn't cover.
# 2026-09-23, per QA feedback (Signum): a real HYBRID event's two buttons - "RSVP TO JOIN
# IN-PERSON <url>" / "RSVP TO JOIN ONLINE <url>" - also came through blank, since neither matched
# "rsvp" immediately followed by the URL bracket. `(?:[^<\n]{0,25})?` allows a short run of extra
# button words between the keyword and the bracket without opening this up to matching across an
# unrelated later "<...>" elsewhere in the body.
_REGISTRATION_LINK_RE = re.compile(r"(?:register(?:\s+(?:here|now|today))?|rsvp)(?:[^<\n]{0,25})?\s*<([^>\s]+)>", re.I)


def extract_registration_link(body: str) -> str:
    match = _REGISTRATION_LINK_RE.search(body or "")
    return match.group(1) if match else ""


# --------------------------------------------------------------------------- #
# Geography / topic inference
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
    "BOLIVIA": "Bolivia", "ZAMBIA": "Zambia",
    # Added 2026-09-23 for the Middle East Council "Daily Brief" country-section splitter (see
    # generic_vendor_parser.py) - found in real captured emails, missing from the original list
    # (which was built for Eurasia Group's own section labels). A systematic scan of every real
    # country-header-shaped line across all 37 real captured MEC emails caught Algeria/Morocco/
    # Tunisia missing entirely - their whole sections were silently dropped, not just mistagged,
    # since a line with no recognized alias never opens a country section at all.
    "TÜRKIYE": "Turkey", "TURKIYE": "Turkey", "SUDAN": "Sudan", "LIBYA": "Libya",
    "PALESTINE": "Palestine", "ALGERIA": "Algeria", "MOROCCO": "Morocco", "TUNISIA": "Tunisia",
    "UNITED ARAB EMIRATE": "United Arab Emirates",
    # 2026-09-23, per QA feedback ("a lot of empty geography cells"): this list was built
    # reactively, one gap at a time (originally just Eurasia Group's own section labels, later a
    # handful more for MEC) - never a real attempt at global coverage. A production check of the
    # single biggest remaining blank-geography bucket (Eurasia Group, 78 rows) found real, common
    # countries missing entirely - South Africa (ANC/local elections coverage), Argentina (Milei/
    # PASO coverage), Germany (Merz/CDU coverage) - none of which are obscure. Expanded broadly
    # here to the countries/regions this tool's own sources (Eurasia Group's global political-risk
    # coverage especially, plus Signum/MEC) realistically discuss, rather than patching one more
    # name every time a reviewer happens to notice another gap.
    "GERMANY": "Germany", "FRANCE": "France", "ITALY": "Italy", "SPAIN": "Spain",
    "PORTUGAL": "Portugal", "NETHERLANDS": "Netherlands", "BELGIUM": "Belgium", "POLAND": "Poland",
    "SWEDEN": "Sweden", "NORWAY": "Norway", "DENMARK": "Denmark", "FINLAND": "Finland",
    "SWITZERLAND": "Switzerland", "AUSTRIA": "Austria", "GREECE": "Greece", "IRELAND": "Ireland",
    "HUNGARY": "Hungary", "ROMANIA": "Romania", "CZECH REPUBLIC": "Czech Republic",
    "SLOVAKIA": "Slovakia", "BULGARIA": "Bulgaria", "CROATIA": "Croatia", "SERBIA": "Serbia",
    "GEORGIA": "Georgia", "ARMENIA": "Armenia", "AZERBAIJAN": "Azerbaijan", "BELARUS": "Belarus",
    "MOLDOVA": "Moldova", "KAZAKHSTAN": "Kazakhstan", "UZBEKISTAN": "Uzbekistan",
    "MONGOLIA": "Mongolia",
    "SOUTH AFRICA": "South Africa", "NIGERIA": "Nigeria", "KENYA": "Kenya",
    "ETHIOPIA": "Ethiopia", "GHANA": "Ghana", "SOMALIA": "Somalia", "SOUTH SUDAN": "South Sudan",
    "DEMOCRATIC REPUBLIC OF CONGO": "DRC", "DRC": "DRC", "ANGOLA": "Angola",
    "MOZAMBIQUE": "Mozambique", "ZIMBABWE": "Zimbabwe", "BOTSWANA": "Botswana",
    "NAMIBIA": "Namibia", "RWANDA": "Rwanda", "UGANDA": "Uganda", "TANZANIA": "Tanzania",
    "SENEGAL": "Senegal", "MALI": "Mali", "NIGER": "Niger", "CHAD": "Chad",
    "CAMEROON": "Cameroon", "IVORY COAST": "Ivory Coast", "CÔTE D'IVOIRE": "Ivory Coast",
    "ARGENTINA": "Argentina", "MEXICO": "Mexico", "CHILE": "Chile", "COLOMBIA": "Colombia",
    "PERU": "Peru", "CUBA": "Cuba", "PANAMA": "Panama", "ECUADOR": "Ecuador",
    "PARAGUAY": "Paraguay", "URUGUAY": "Uruguay", "GUATEMALA": "Guatemala",
    "HONDURAS": "Honduras", "EL SALVADOR": "El Salvador", "NICARAGUA": "Nicaragua",
    "DOMINICAN REPUBLIC": "Dominican Republic", "HAITI": "Haiti", "JAMAICA": "Jamaica",
    "CANADA": "Canada",
    "INDONESIA": "Indonesia", "MALAYSIA": "Malaysia", "THAILAND": "Thailand",
    "VIETNAM": "Vietnam", "PHILIPPINES": "Philippines", "SINGAPORE": "Singapore",
    "MYANMAR": "Myanmar", "CAMBODIA": "Cambodia", "LAOS": "Laos", "BANGLADESH": "Bangladesh",
    "SRI LANKA": "Sri Lanka", "NEPAL": "Nepal", "AUSTRALIA": "Australia",
    "NEW ZEALAND": "New Zealand",
    "ASEAN": "ASEAN", "SUB-SAHARAN AFRICA": "Sub-Saharan Africa",
    "LATIN AMERICA": "Latin America", "SAHEL": "Sahel", "BALKANS": "Balkans",
    "CENTRAL ASIA": "Central Asia", "HORN OF AFRICA": "Horn of Africa",
}


def geography_from_label(label: str) -> str:
    cleaned = re.sub(r"\([^)]*\)", " ", label).replace(".", "/")
    countries = []
    for part in re.split(r"\s*(?:/|,|;|\band\b|&|\+|-|–)\s*", cleaned, flags=re.I):
        part = compact_text(part).strip(" -:")
        if not part:
            continue
        canonical = COUNTRY_ALIASES.get(part.upper())
        if not canonical:
            # 2026-10-01, per QA feedback ("only captures the first country when multiple are
            # mentioned"): a split segment can carry trailing descriptive words after the real
            # country name - e.g. "Russia-Ukraine energy truce: Moscow's winning strategy" splits
            # on "-" into "Russia" and "Ukraine energy truce", and the exact-match lookup above
            # only ever matches the bare first one. Confirmed live on several real Signum titles
            # that lost their second country this exact way. Falls back to a word-boundary search
            # for any known country name within the segment, same technique infer_geography()
            # already uses on free text.
            for candidate, canon in COUNTRY_ALIASES.items():
                if re.search(rf"\b{re.escape(candidate)}\b", part, re.I):
                    canonical = canon
                    break
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
        # 2026-09-24, per user feedback: "What Comes After Bab al-Mandab?" (a real MEC Daily
        # Brief title) went untagged - the strait's name has several real transliteration
        # variants ("Bab-el-Mandeb", "Bab al-Mandab", "Bab el-Mandeb", ...), and the old pattern
        # only matched the one exact spelling. Widened to cover the "el"/"al" and "eb"/"ab"
        # variation with or without hyphens.
        "Red Sea": r"\bBab[- ]?(?:el|al)[- ]?Mand(?:eb|ab)\b|\bHouthis?\b",
        # 2026-09-23: found live that "South Africa"/"Argentina"/"Germany" etc. as bare country
        # names almost never appear in real Eurasia Group coverage - it's near-always an
        # adjective form ("South African Reserve Bank"), a leader's name ("Milei", "Merz"), or a
        # capital ("Johannesburg", "Buenos Aires", "Berlin") instead, none of which the literal
        # COUNTRY_ALIASES entries above can match (a trailing `\b` after "Africa" correctly does
        # NOT match "African" - same word, different boundary). Same pattern as the Israel/
        # Russia/Ukraine entries already here, extended to the countries the QA-feedback sample
        # actually surfaced as missing.
        "South Africa": r"\bSouth African\b|\bANC\b|\bJohannesburg\b|\bPretoria\b|\bCape Town\b",
        "Argentina": r"\bArgentine\b|\bMilei\b|\bBuenos Aires\b",
        "Germany": r"\bGerman\b|\bMerz\b|\bBerlin\b|\bBundestag\b|\bCDU\b",
        "France": r"\bFrench\b|\bMacron\b|\bParis\b|\bLe Pen\b",
        "Mexico": r"\bMexican\b|\bSheinbaum\b",
        "Brazil": r"\bBrazilian\b|\bLula\b|\bBolsonaro\b",
        "China": r"\bChinese\b|\bXi Jinping\b|\bBeijing\b",
        "UK": r"\bBritish\b|\bWestminster\b|\bDowning Street\b",
        "Iran": r"\bIranian\b|\bTehran\b",
        "Saudi Arabia": r"\bSaudi\b|\bRiyadh\b",
        "Turkey": r"\bTurkish\b|\bErdo[gğ]an\b|\bAnkara\b",
        "Egypt": r"\bEgyptian\b|\bEl Sisi\b|\bCairo\b",
        "India": r"\bIndian\b|\bModi\b|\bNew Delhi\b",
        "Pakistan": r"\bPakistani\b|\bIslamabad\b",
        "Nigeria": r"\bNigerian\b|\bAbuja\b|\bLagos\b",
        "Venezuela": r"\bVenezuelan\b|\bMaduro\b|\bCaracas\b",
        "Malaysia": r"\bMalaysian\b|\bAnwar Ibrahim\b|\bPakatan Harapan\b",
        # 2026-09-23, per user follow-up ("what is the issue with the empty ones"): checked every
        # remaining blank row directly - most are genuinely thematic/administrative with no real
        # country to assign, but several DO name a current head of state not yet aliased.
        "Senegal": r"\bSenegalese\b|\bDiomaye Faye\b|\bFaye\b",
        "Indonesia": r"\bIndonesian\b|\bPrabowo\b|\bJakarta\b",
        "Colombia": r"\bColombian\b|\bGustavo Petro\b|\bBogot[aá]\b",
        "Ecuador": r"\bEcuadorian\b|\bDaniel Noboa\b|\bNoboa\b",
        "Peru": r"\bPeruvian\b|\bFujimori\b|\bLima\b",
        "Bangladesh": r"\bBangladeshi\b|\bTarique Rahman\b|\bDhaka\b",
        "Sri Lanka": r"\bSri Lankan\b|\bDissanayake\b|\bColombo\b",
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
    "Sudan": "SD", "Libya": "LY", "Palestine": "PS",
    "Algeria": "DZ", "Morocco": "MA", "Tunisia": "TN",
    "Germany": "DE", "France": "FR", "Italy": "IT", "Spain": "ES", "Portugal": "PT",
    "Netherlands": "NL", "Belgium": "BE", "Poland": "PL", "Sweden": "SE", "Norway": "NO",
    "Denmark": "DK", "Finland": "FI", "Switzerland": "CH", "Austria": "AT", "Greece": "GR",
    "Ireland": "IE", "Hungary": "HU", "Romania": "RO", "Czech Republic": "CZ",
    "Slovakia": "SK", "Bulgaria": "BG", "Croatia": "HR", "Serbia": "RS", "Georgia": "GE",
    "Armenia": "AM", "Azerbaijan": "AZ", "Belarus": "BY", "Moldova": "MD",
    "Kazakhstan": "KZ", "Uzbekistan": "UZ", "Mongolia": "MN",
    "South Africa": "ZA", "Nigeria": "NG", "Kenya": "KE", "Ethiopia": "ET", "Ghana": "GH",
    "Somalia": "SO", "South Sudan": "SS", "DRC": "CD", "Angola": "AO", "Mozambique": "MZ",
    "Zimbabwe": "ZW", "Botswana": "BW", "Namibia": "NA", "Rwanda": "RW", "Uganda": "UG",
    "Tanzania": "TZ", "Senegal": "SN", "Mali": "ML", "Niger": "NE", "Chad": "TD",
    "Cameroon": "CM", "Ivory Coast": "CI",
    "Argentina": "AR", "Mexico": "MX", "Chile": "CL", "Colombia": "CO", "Peru": "PE",
    "Cuba": "CU", "Panama": "PA", "Ecuador": "EC", "Paraguay": "PY", "Uruguay": "UY",
    "Guatemala": "GT", "Honduras": "HN", "El Salvador": "SV", "Nicaragua": "NI",
    "Dominican Republic": "DO", "Haiti": "HT", "Jamaica": "JM", "Canada": "CA",
    "Indonesia": "ID", "Malaysia": "MY", "Thailand": "TH", "Vietnam": "VN",
    "Philippines": "PH", "Singapore": "SG", "Myanmar": "MM", "Cambodia": "KH",
    "Laos": "LA", "Bangladesh": "BD", "Sri Lanka": "LK", "Nepal": "NP",
    "Australia": "AU", "New Zealand": "NZ",
}


def alpha2_from_geography(geography: str) -> str:
    return ", ".join(ALPHA2_CODES[part.strip()] for part in geography.split(",") if part.strip() in ALPHA2_CODES)


REGIONAL_LABELS = {
    "Europe", "European Union", "MENA", "Gulf", "Red Sea", "ASEAN", "Sub-Saharan Africa",
    "Latin America", "Sahel", "Balkans", "Central Asia", "Horn of Africa",
}
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
