import re
import json
import sys
import unicodedata
from copy import deepcopy
from pathlib import Path
from urllib.parse import urlparse, unquote

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

try:
    import pycountry
except Exception:
    pycountry = None

try:
    from email_validator import validate_email as _validate_email
except Exception:
    _validate_email = None


# =============================================================================
# CONFIG
# =============================================================================

INPUT_FILE = (
    r"output/Maheshwari/"
    r"Maheshwari_CANCER_University_DATA_08-09-2026/"
    r"Maheshwari_CANCER_University_DATA_08-09-2026.xlsx"
)

OUTPUT_DIR = (
    r"cleaned_output/Maheshwari/"
    r"Maheshwari_CANCER_University_DATA_08-09-2026"
)

OUTPUT_FILENAME = "Maheshwari_CANCER_University_DATA_08-09-2026_CLEANED.xlsx"
SHEET_NAME = "Scraped Data"

EMAIL_COLUMN = "email"
NAME_COLUMN = "name"

# Core policy
EMAIL_IS_MANDATORY = True
NAME_FROM_EMAIL_IF_BAD_OR_EMPTY = True
REPLACE_WRONG_NAME_WHEN_EMAIL_IS_STRONG = True

# Generic/shared mailboxes are valid emails, but they must not manufacture a person name.
KEEP_SHARED_EMAILS = True
BLANK_PERSON_NAME_FOR_SHARED_EMAIL = True

# Deduplication
DEDUPE_PERSONAL_EMAILS = True
PRESERVE_SHARED_EMAIL_PER_PERSON = True

# Alternate-email protection
MAX_ALTERNATE_EMAILS = 3
CLEAR_SUSPICIOUS_ALTERNATES = True

# Audit sheets
WRITE_CLEANING_AUDIT = True
WRITE_REJECTED_ROWS = True


# =============================================================================
# EXPECTED COLUMNS
# =============================================================================

EXPECTED_COLUMNS = [
    "name",
    "email",
    "country",
    "alternate_emails",
    "email_type",
    "email_conflict",
    "employee_name",
    "employee_email",
    "page_type",
    "journal_name",
    "editorial_role",
    "academic_title",
    "academic_rank",
    "specialty",
    "affiliation",
    "university",
    "faculty",
    "school",
    "department",
    "institute",
    "division",
    "city",
    "address",
    "phone",
    "orcid",
    "google_scholar",
    "scopus_author_id",
    "researcher_id",
    "pubmed",
    "profile_url",
    "personal_homepage",
    "source_url",
    "country_source",
    "confidence",
    "extraction_method",
    "scrape_status",
]


# =============================================================================
# EMAIL RULES
# =============================================================================

EMAIL_RE = re.compile(
    r"^[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?"
    r"(?:\.[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?)+$",
    re.I,
)

# Common extraction-artifact words that are almost never intended public suffixes
# in scraped institutional email data.
ARTIFACT_TLDS = {
    "turn", "human", "results", "result", "page", "pages", "html", "text",
    "example", "invalid", "test", "none", "null", "undefined", "here",
    "click", "below", "find", "going", "however", "profiles", "foroughi",
    "bathina", "priors", "engineering", "visit", "with", "from", "this",
    "that", "where", "when", "you", "my", "his", "her", "she", "we", "for",
}

FAKE_EXACT_EMAILS = {
    "or@all.turn",
    "study@midgestation.human",
    "also@delivery.results",
    "analyst@musc.his",
    "person@email.cz",
    "bluesky@bsky.social",
}

GENERIC_LOCAL_ROOTS = {
    "info", "contact", "contacts", "office", "admin", "administrator",
    "webmaster", "support", "help", "reception", "secretary", "faculty",
    "staff", "team", "department", "dept", "editor", "editorial",
    "enquiries", "inquiries", "communications", "communication", "media",
    "press", "admissions", "admission", "student", "students", "website",
    "research", "researchteam", "clinic", "lab", "laboratory", "hr",
    "jobs", "careers", "leadership", "postbac", "dlag", "all",
}

GENERIC_LOCAL_SUBSTRINGS = {
    "webmaster", "contact", "support", "admin", "faculty", "staff",
    "department", "editorial", "secretary", "reception", "admission",
    "studentoffice", "media", "communications",
}

# Academic/public-sector suffixes. Used only to repair obvious appended-word artifacts:
#   abi.herrmann@admin.cam.ac.uk.cross -> abi.herrmann@admin.cam.ac.uk
ACADEMIC_SUFFIXES = [
    ".ac.uk", ".edu.au", ".edu", ".ac.in", ".edu.sg", ".ac.jp", ".ac.kr",
    ".edu.cn", ".ac.cn", ".edu.hk", ".ac.nz", ".edu.my", ".ac.za",
    ".gov.uk", ".gov.au", ".gov", ".org", ".com",
]

TLD_COUNTRY = {
    "au": "Australia", "at": "Austria", "bd": "Bangladesh", "be": "Belgium",
    "br": "Brazil", "ca": "Canada", "ch": "Switzerland", "cn": "China",
    "cz": "Czech Republic", "de": "Germany", "dk": "Denmark", "eg": "Egypt",
    "es": "Spain", "et": "Ethiopia", "fi": "Finland", "fr": "France",
    "gr": "Greece", "hk": "Hong Kong", "id": "Indonesia", "ie": "Ireland",
    "il": "Israel", "in": "India", "ir": "Iran", "it": "Italy", "jp": "Japan",
    "ke": "Kenya", "kr": "South Korea", "lb": "Lebanon", "lk": "Sri Lanka",
    "lv": "Latvia", "mt": "Malta", "my": "Malaysia", "ng": "Nigeria",
    "nl": "Netherlands", "no": "Norway", "np": "Nepal", "nz": "New Zealand",
    "ph": "Philippines", "pk": "Pakistan", "pl": "Poland", "pt": "Portugal",
    "sa": "Saudi Arabia", "se": "Sweden", "sg": "Singapore", "th": "Thailand",
    "tn": "Tunisia", "tr": "Turkey", "tw": "Taiwan", "ug": "Uganda",
    "uk": "United Kingdom", "us": "United States", "za": "South Africa",
}


# =============================================================================
# NAME RULES
# =============================================================================

BAD_NAME_EXACT = {
    "", "telephone", "phone", "email", "e-mail", "contact", "contact us",
    "faculty", "staff", "team", "research team", "leadership",
    "administrative leadership", "view profile", "view full profile",
    "profile", "read more", "learn more", "home", "website", "location",
    "follow us", "follow us on bluesky", "bluesky", "researchers",
    "professional staff", "academic staff", "academic visitors",
    "administrative staff", "admissions information", "about us",
    "about the group", "group leaders", "our people", "our team",
    "student support", "student wellbeing", "current students",
}

BAD_NAME_PHRASES = {
    "skip to", "click here", "learn more", "read more", "view profile",
    "view full profile", "contact us", "find a doctor", "find a researcher",
    "for more information", "for all media inquiries", "employment opportunities",
    "javascript isn't enabled", "javascript is not enabled", "submit applications",
    "learn more about this group", "track attendance", "by visiting the",
    "or visiting the", "or dr.", "or dr ", "and dr.", "and dr ",
    "federally funded website", "application information", "seminar:",
}

ADDRESS_WORDS = {
    "road", "rd", "street", "st", "avenue", "ave", "boulevard", "blvd",
    "lane", "ln", "drive", "dr", "highway", "hwy", "box", "suite", "unit",
    "building", "floor", "room", "campus", "postcode", "zip",
}

ROLE_WORDS = {
    "professor", "associate professor", "assistant professor", "lecturer",
    "senior lecturer", "reader", "research fellow", "senior research fellow",
    "research associate", "senior research associate", "postdoctoral researcher",
    "postdoctoral fellow", "academic visitor", "honorary lecturer",
    "honorary research fellow", "director", "manager", "coordinator",
    "administrator", "analyst", "technician", "scientist", "researcher",
    "physician", "consultant",
}

PREFIX_RE = re.compile(
    r"^\s*(?:(?:professor|prof|dr|doctor|doc|mudr|mr|mrs|miss|ms)\.?\s+)+",
    re.I,
)

CREDENTIALS = (
    r"(?:MD|M\.D\.|PhD|Ph\.D\.|DO|D\.O\.|PsyD|EdD|JD|MPH|MHA|MPA|MBA|"
    r"MSW|MEd|MSc|MS|MA|BSc|BA|MBBS|MBChB|BMedSci|ChB|MRCP|MRCPCH|"
    r"FRCPCH|CSc|ScD|DDS|DMD|RN|PharmD|FACOG|FACS|FAPA|FRCP|FRCS)"
)

ROLE_SUFFIX_RE = re.compile(
    r"\b(?:"
    r"Emeritus Professor(?:\s+of\b.*)?|Associate Professor(?:\s+(?:in|of)\b.*)?|"
    r"Assistant Professor(?:\s+(?:in|of)\b.*)?|Professor(?:\s+(?:in|of)\b.*)?|"
    r"Reader(?:\s+in\b.*)?|Senior Lecturer(?:\s+in\b.*)?|Lecturer(?:\s+in\b.*)?|"
    r"Honorary Senior Research Fellow|Honorary Research Fellow|Honorary Senior Lecturer|"
    r"Honorary Lecturer|Senior Research Fellow|Research Fellow|Advanced Research Fellow|"
    r"Senior Research Associate|Research Associate|Postdoctoral Researcher|"
    r"Postdoctoral Research Fellow|Postdoc|PhD Student|Doctoral Researcher|"
    r"Senior Transport Analyst|Transport Analyst|Teaching Fellow|Senior Teaching Fellow|"
    r"Academic Visitor|Director\b.*|Manager\b.*|Coordinator\b.*|Administrator\b.*|"
    r"Technician\b.*|Scientist\b.*"
    r")$",
    re.I,
)


# =============================================================================
# BASIC HELPERS
# =============================================================================

def clean_text(value):
    if value is None:
        return ""
    text = str(value)
    text = (
        text.replace("\u00a0", " ")
        .replace("\u200b", "")
        .replace("\ufeff", "")
        .replace("\u00ad", "")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )
    return re.sub(r"\s+", " ", text).strip()


def ascii_norm(value):
    return (
        unicodedata.normalize("NFKD", clean_text(value))
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )


def safe_float(value, default=0.0):
    try:
        return float(value)
    except Exception:
        return default


def append_method(existing, marker):
    existing = clean_text(existing)
    parts = [x for x in existing.split("+") if x]
    if marker not in parts:
        parts.append(marker)
    return "+".join(parts)


def host_from_url(url):
    try:
        return urlparse(clean_text(url)).netloc.lower().replace("www.", "")
    except Exception:
        return ""


# =============================================================================
# EMAIL NORMALIZATION / REPAIR / VALIDATION
# =============================================================================

def decode_obfuscated_email(value):
    value = clean_text(value)
    if not value:
        return ""

    value = unquote(value)
    value = re.sub(r"(?i)^mailto:", "", value).strip()
    value = re.sub(r"(?i)\s*\[\s*at\s*\]\s*", "@", value)
    value = re.sub(r"(?i)\s*\(\s*at\s*\)\s*", "@", value)
    value = re.sub(r"(?i)\s+at\s+", "@", value)
    value = re.sub(r"(?i)\s*\[\s*dot\s*\]\s*", ".", value)
    value = re.sub(r"(?i)\s*\(\s*dot\s*\)\s*", ".", value)
    value = re.sub(r"(?i)\s+dot\s+", ".", value)
    value = re.sub(r"\s*@\s*", "@", value)
    value = re.sub(r"\s*\.\s*", ".", value)
    return value.strip()


def normalize_email(value):
    value = decode_obfuscated_email(value).lower()
    if not value:
        return ""

    # Pull one email-looking token from contaminated text.
    m = re.search(
        r"[a-z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-z0-9.-]+\.[a-z]{2,32}",
        value,
        re.I,
    )
    if not m:
        return ""

    email = m.group(0).strip(" <>[](){};,:'\"./")
    return email


def repair_appended_domain_artifact(email, source_url="", profile_url=""):
    """
    Repair only obvious cases where a normal institutional domain is followed
    by one extra scraped word.

    Example:
        abi.herrmann@admin.cam.ac.uk.cross
        -> abi.herrmann@admin.cam.ac.uk
    """
    email = normalize_email(email)
    if "@" not in email:
        return email, False

    local, domain = email.rsplit("@", 1)
    domain_low = domain.lower()

    # Strong academic-suffix repair.
    for suffix in ACADEMIC_SUFFIXES:
        pos = domain_low.find(suffix + ".")
        if pos >= 0:
            candidate_domain = domain_low[: pos + len(suffix)]
            trailing = domain_low[pos + len(suffix) + 1 :]
            if trailing and re.fullmatch(r"[a-z]{2,24}", trailing):
                candidate = f"{local}@{candidate_domain}"
                return candidate, True

    # Source/profile host can prove an institutional domain embedded in artifact.
    known_hosts = [host_from_url(source_url), host_from_url(profile_url)]
    known_hosts = [h for h in known_hosts if h and "." in h]

    for host in known_hosts:
        parts = host.split(".")
        # Try progressively shorter institutional suffixes.
        for i in range(len(parts)):
            suffix = ".".join(parts[i:])
            if len(suffix.split(".")) < 2:
                continue
            if domain_low == suffix:
                return email, False
            marker = suffix + "."
            if marker in domain_low and domain_low.startswith(domain_low.split(marker)[0]):
                before, after = domain_low.split(marker, 1)
                candidate_domain = (before + suffix).strip(".")
                if after and re.fullmatch(r"[a-z]{2,24}", after):
                    return f"{local}@{candidate_domain}", True

    return email, False


def is_valid_email(email):
    email = normalize_email(email)
    if not email or email in FAKE_EXACT_EMAILS:
        return False

    if not EMAIL_RE.fullmatch(email):
        return False

    if email.count("@") != 1 or ".." in email:
        return False

    local, domain = email.rsplit("@", 1)
    if not local or not domain or "." not in domain:
        return False

    labels = domain.split(".")
    if any(not label or label.startswith("-") or label.endswith("-") for label in labels):
        return False

    tld = labels[-1].lower()
    if tld in ARTIFACT_TLDS:
        return False

    # Reject common sentence-fragment domains.
    if domain in {
        "all.turn", "midgestation.human", "delivery.results",
        "risk.support",
    }:
        return False

    if _validate_email is not None:
        try:
            _validate_email(email, check_deliverability=False)
        except Exception:
            return False

    return True


def normalize_and_validate_email(value, source_url="", profile_url=""):
    email = normalize_email(value)
    if not email:
        return "", False, "invalid_email"

    repaired, changed = repair_appended_domain_artifact(
        email,
        source_url=source_url,
        profile_url=profile_url,
    )

    if is_valid_email(repaired):
        return repaired, changed, ""

    if repaired != email and is_valid_email(email):
        return email, False, ""

    return "", changed, "invalid_email"


def is_shared_email(email):
    email = normalize_email(email)
    if "@" not in email:
        return False

    local = email.split("@", 1)[0].lower()
    root = re.split(r"[._+\-]", local)[0]

    if local in GENERIC_LOCAL_ROOTS or root in GENERIC_LOCAL_ROOTS:
        return True

    return any(token in local for token in GENERIC_LOCAL_SUBSTRINGS)


# =============================================================================
# EMAIL -> NAME
# =============================================================================

def email_local_parts(email):
    email = normalize_email(email)
    if "@" not in email:
        return []

    local = email.split("@", 1)[0].split("+", 1)[0].lower()
    raw = [x for x in re.split(r"[._\-]+", local) if x]

    parts = []
    for token in raw:
        token = re.sub(r"\d+$", "", token)
        token = re.sub(r"[^a-zà-öø-ÿā-ž]", "", token, flags=re.I)
        if token:
            parts.append(token)
    return parts


def pretty_token(token):
    if len(token) == 1:
        return token.upper() + "."
    return token[:1].upper() + token[1:].lower()


def strong_name_from_email(email):
    """
    Strong derivation only when email has explicit token boundaries.

    john.smith@x.edu    -> John Smith
    j.smith@x.edu       -> J. Smith
    maria.apud-bell@x   -> Maria Apud Bell

    compact usernames such as aapatel@uw.edu are NOT split blindly.
    """
    if is_shared_email(email):
        return ""

    parts = email_local_parts(email)
    if len(parts) < 2:
        return ""

    meaningful = [x for x in parts if len(x) >= 2]
    if len(meaningful) < 1:
        return ""

    return " ".join(pretty_token(x) for x in parts)


def weak_name_from_email(email):
    """
    Last-resort name because the user requested a name for bad/empty rows.

    For compact usernames we can only format the mailbox, not infer a full legal
    name. Example: aapatel -> Aapatel. The provenance is marked as weak.
    """
    if is_shared_email(email):
        return ""

    strong = strong_name_from_email(email)
    if strong:
        return strong

    local = normalize_email(email).split("@", 1)[0].split("+", 1)[0]
    local = re.sub(r"\d+$", "", local)
    local = re.sub(r"[^A-Za-zÀ-ÖØ-öø-ÿĀ-ž]", "", local)

    if len(local) < 3:
        return ""

    return pretty_token(local)


# =============================================================================
# PROFILE URL -> NAME FALLBACK
# =============================================================================

def name_from_profile_url(profile_url):
    """
    Use only person-looking URL leafs.

    /people/john-smith -> John Smith
    /profile/123/john-smith -> John Smith
    researcherprofiles.../Daniel.Runco -> Daniel Runco
    """
    url = clean_text(profile_url)
    if not url:
        return ""

    try:
        p = urlparse(url)
        segments = [unquote(x) for x in p.path.split("/") if x]
    except Exception:
        return ""

    if not segments:
        return ""

    candidates = []
    for seg in reversed(segments[-3:]):
        seg = re.sub(r"\.(?:html?|php|aspx?)$", "", seg, flags=re.I)
        if not seg or seg.isdigit():
            continue
        if len(seg) > 80:
            continue
        if re.search(r"[-._]", seg):
            parts = [x for x in re.split(r"[-._]+", seg) if x]
            if 2 <= len(parts) <= 6 and all(re.search(r"[A-Za-z]", x) for x in parts):
                candidates.append(" ".join(pretty_token(x) for x in parts))

    for candidate in candidates:
        if not is_bad_name(candidate):
            return candidate
    return ""


# =============================================================================
# NAME CLEANING / VALIDATION
# =============================================================================

def clean_initials(name):
    return re.sub(r"(?<=[A-Za-z])\.(?=[A-Za-z])", ". ", name)


def remove_credentials(name):
    name = clean_text(name)
    for _ in range(8):
        old = name
        name = re.sub(
            rf"(?:[,\s]+{CREDENTIALS}\.?)\s*$",
            "",
            name,
            flags=re.I,
        )
        name = re.sub(r"\s*\((?:Hons|Honours)\)\s*$", "", name, flags=re.I)
        name = name.strip(" ,;-")
        if name == old:
            break
    return clean_text(name)


def convert_surname_first(name):
    if "," not in name:
        return name

    left, right = name.split(",", 1)
    left, right = clean_text(left), clean_text(right)

    if not left or not right:
        return left or right

    # Do not rearrange sentence-like values.
    if len(left.split()) > 3 or len(right.split()) > 5:
        return name

    right = PREFIX_RE.sub("", right)
    right = remove_credentials(right)
    right = clean_initials(right)

    if not right:
        return left

    return f"{right} {left}"


def strip_role_suffix(name):
    name = clean_text(name)
    if not name:
        return ""

    # Remove role text only if at least two person-like tokens remain.
    m = ROLE_SUFFIX_RE.search(name)
    if m:
        head = name[:m.start()].strip(" ,;-")
        if len(tokenise_name(head)) >= 2:
            return head
    return name


def clean_person_name(name):
    name = clean_text(name)
    if not name:
        return ""

    name = re.sub(r"^\s*profile\s+of\s+", "", name, flags=re.I)
    name = PREFIX_RE.sub("", name)
    name = convert_surname_first(name)
    name = PREFIX_RE.sub("", name)
    name = remove_credentials(name)
    name = clean_initials(name)
    name = strip_role_suffix(name)
    name = re.sub(r"\s+", " ", name)
    return name.strip(" ,;-|()")


def tokenise_name(name):
    value = ascii_norm(name)
    tokens = re.findall(r"[a-z]+", value)

    ignore = {
        "dr", "doctor", "prof", "professor", "mr", "mrs", "miss", "ms",
        "md", "phd", "do", "jd", "associate", "assistant", "senior",
        "honorary", "research", "fellow", "lecturer", "reader",
    }

    return [x for x in tokens if len(x) >= 2 and x not in ignore]


def looks_like_address(name):
    n = clean_text(name)
    low = ascii_norm(n)

    if re.match(r"^\d{2,}\b", n):
        return True

    tokens = set(re.findall(r"[a-z]+", low))
    address_hits = len(tokens & ADDRESS_WORDS)
    if address_hits >= 2:
        return True

    if re.search(r"\b(?:po box|p\.o\. box|postcode|zip)\b", low):
        return True

    return False


def looks_like_sentence_or_label(name):
    n = clean_text(name)
    low = ascii_norm(n)

    if not n:
        return True

    if low.strip(" :;,.|-") in BAD_NAME_EXACT:
        return True

    if any(p in low for p in BAD_NAME_PHRASES):
        return True

    # Parenthesis/punctuation fragments from nearby prose.
    if n.startswith((")", "(", ",", ";", ":", "-", "–", "—")):
        return True

    if n.endswith(("(", ")", ":", ";")):
        return True

    # Full sentences / page titles.
    words = n.split()
    if len(words) > 9:
        return True

    if len(words) >= 5 and re.search(
        r"\b(?:the|and|or|by|for|from|with|visiting|application|information|"
        r"website|seminar|progress|research|contacting|attendees)\b",
        low,
    ):
        return True

    # Dates/page titles.
    if re.search(r"\b(?:19|20)\d{2}\b", n) and len(words) > 3:
        return True

    if looks_like_address(n):
        return True

    return False


def is_bad_name(name):
    name = clean_person_name(name)
    if not name:
        return True

    if "@" in name or "://" in name:
        return True

    if len(name) > 100:
        return True

    if re.fullmatch(r"[\W\d_]+", name):
        return True

    if looks_like_sentence_or_label(name):
        return True

    tokens = tokenise_name(name)
    if not tokens:
        return True

    # Organization/section labels.
    low = ascii_norm(name)
    org_words = {
        "university", "hospital", "department", "faculty", "school", "college",
        "institute", "center", "centre", "clinic", "laboratory", "office",
        "committee", "team", "group", "program", "programme",
    }
    if sum(1 for word in org_words if word in low) >= 1 and len(tokens) <= 6:
        return True

    return False


# =============================================================================
# NAME/EMAIL CONSISTENCY
# =============================================================================

def email_name_tokens(email):
    strong = strong_name_from_email(email)
    return tokenise_name(strong) if strong else []


def name_matches_email(name, email):
    nt = tokenise_name(name)
    et = email_name_tokens(email)

    if not nt or not et:
        return None  # not enough evidence

    overlap = set(nt) & set(et)
    if overlap:
        return True

    # Initial + surname: J Smith vs John Smith style handling.
    n_last = nt[-1] if nt else ""
    e_last = et[-1] if et else ""
    if n_last and e_last and (n_last == e_last or n_last in e_last or e_last in n_last):
        return True

    return False


def profile_name_supports_email_name(profile_name, email_name):
    pt = set(tokenise_name(profile_name))
    et = set(tokenise_name(email_name))
    if not pt or not et:
        return False
    return bool(pt & et)


def choose_best_name(original_name, email, profile_url=""):
    """
    Returns: (name, source_marker, changed, reason)

    Priority when original name is wrong/empty:
      1. strong email-derived name
      2. person-like profile URL name
      3. weak email local-part formatting
      4. blank for shared/role mailbox
    """
    original_clean = clean_person_name(original_name)
    bad = is_bad_name(original_clean)

    strong_email_name = strong_name_from_email(email)
    profile_name = name_from_profile_url(profile_url)
    weak_email_name = weak_name_from_email(email)

    # Shared mailboxes must not create fake person identities.
    if is_shared_email(email):
        if original_clean and not bad:
            return original_clean, "source_name", original_clean != clean_text(original_name), "shared_email_keep_valid_source_name"
        if BLANK_PERSON_NAME_FOR_SHARED_EMAIL:
            return "", "shared_email_no_person_name", bool(clean_text(original_name)), "shared_email_bad_or_missing_name"
        return weak_email_name, "email_local_weak", True, "shared_email_user_requested_fallback"

    # Empty/bad source name.
    if bad:
        if strong_email_name:
            return strong_email_name, "email_strong", True, "bad_or_empty_name_replaced_from_email"

        if profile_name:
            return profile_name, "profile_url", True, "bad_or_empty_name_replaced_from_profile_url"

        if weak_email_name:
            return weak_email_name, "email_local_weak", True, "bad_or_empty_name_replaced_from_compact_email"

        return "", "blank", bool(clean_text(original_name)), "bad_name_no_safe_replacement"

    # Existing source name is plausible. Replace only on strong contradiction.
    if REPLACE_WRONG_NAME_WHEN_EMAIL_IS_STRONG and strong_email_name:
        match = name_matches_email(original_clean, email)

        if match is False:
            # Profile URL can act as a second vote.
            if profile_name and profile_name_supports_email_name(profile_name, strong_email_name):
                return strong_email_name, "email_strong+profile_support", True, "name_email_mismatch_profile_supports_email"

            # firstname.lastname-like email is sufficiently strong by itself.
            et = email_name_tokens(email)
            if len(et) >= 2 and all(len(x) >= 2 for x in et):
                return strong_email_name, "email_strong", True, "name_email_strong_mismatch"

    return original_clean, "source_name", original_clean != clean_text(original_name), "kept_source_name"


# =============================================================================
# COUNTRY CLEANING
# =============================================================================

COUNTRY_ALIASES = {
    "usa": "United States",
    "u.s.a.": "United States",
    "united states of america": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "england": "United Kingdom",
    "scotland": "United Kingdom",
    "wales": "United Kingdom",
    "south korea": "South Korea",
    "republic of korea": "South Korea",
    "uae": "United Arab Emirates",
    "turkiye": "Turkey",
    "türkiye": "Turkey",
}


def build_country_lookup():
    lookup = {k.lower(): v for k, v in COUNTRY_ALIASES.items()}
    if pycountry:
        for c in pycountry.countries:
            lookup[c.name.lower()] = c.name
            if hasattr(c, "official_name"):
                lookup[c.official_name.lower()] = c.name
            if hasattr(c, "common_name"):
                lookup[c.common_name.lower()] = c.name
            lookup[c.alpha_2.lower()] = c.name
            lookup[c.alpha_3.lower()] = c.name
    return lookup


COUNTRY_LOOKUP = build_country_lookup()


def normalize_country(value):
    v = clean_text(value).strip(" ,.;:-")
    if not v:
        return ""
    return COUNTRY_LOOKUP.get(v.lower(), COUNTRY_ALIASES.get(v.lower(), v if len(v) <= 80 else ""))


def country_from_email(email):
    email = normalize_email(email)
    if "@" not in email:
        return ""
    domain = email.split("@", 1)[1]
    tld = domain.rsplit(".", 1)[-1].lower()
    return TLD_COUNTRY.get(tld, "")


def country_from_source_url(source_url):
    host = host_from_url(source_url)
    if not host:
        return ""
    tld = host.rsplit(".", 1)[-1].lower()
    return TLD_COUNTRY.get(tld, "")


def clean_country(existing_country, email, source_url):
    existing = normalize_country(existing_country)
    if existing:
        return existing, "existing_clean"

    c = country_from_email(email)
    if c:
        return c, "email_tld_cleaning"

    c = country_from_source_url(source_url)
    if c:
        return c, "source_tld_cleaning"

    return "", ""


# =============================================================================
# ALTERNATE EMAILS
# =============================================================================

def split_emails(value):
    value = clean_text(value)
    if not value:
        return []

    # Extract email tokens rather than blindly splitting on whitespace.
    return re.findall(
        r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,32}",
        value,
    )


def clean_alternate_emails(value, primary_email, source_url="", profile_url="", email_conflict="no"):
    primary = normalize_email(primary_email)
    output = []
    seen = set()

    for raw in split_emails(value):
        email, _, reason = normalize_and_validate_email(
            raw,
            source_url=source_url,
            profile_url=profile_url,
        )
        if reason or not email or email == primary or email in seen:
            continue
        seen.add(email)
        output.append(email)

    if not output:
        return ""

    # Broad-page contamination protection.
    if CLEAR_SUSPICIOUS_ALTERNATES:
        if len(output) > MAX_ALTERNATE_EMAILS:
            return ""

        # If source did not mark an actual conflict, several alternates are suspicious.
        if str(email_conflict or "").lower() != "yes" and len(output) > 1:
            return ""

    return "; ".join(output)


# =============================================================================
# ROW QUALITY / MERGING
# =============================================================================

def row_quality(row):
    score = 0

    if clean_text(row.get("name")):
        score += 30
    if clean_text(row.get("profile_url")):
        score += 30
    if clean_text(row.get("affiliation")):
        score += 10
    if clean_text(row.get("academic_title")):
        score += 10
    if clean_text(row.get("country")):
        score += 5
    if row.get("email_type") == "personal":
        score += 10

    method = clean_text(row.get("extraction_method")).lower()
    if "profile" in method:
        score += 20
    if "mailto" in method:
        score += 8
    if "page_text" in method:
        score -= 10
    if "fallback" in method:
        score -= 5

    score += min(15, int(safe_float(row.get("confidence"), 0) / 10))
    return score


MERGE_FIELDS = [
    "country", "academic_title", "academic_rank", "specialty", "affiliation",
    "university", "faculty", "school", "department", "institute", "division",
    "city", "address", "phone", "orcid", "google_scholar", "scopus_author_id",
    "researcher_id", "pubmed", "profile_url", "personal_homepage", "country_source",
]


def merge_duplicate_rows(a, b):
    """Merge duplicate observations without changing the selected primary identity."""
    best, other = (a, b) if row_quality(a) >= row_quality(b) else (b, a)
    result = deepcopy(best)

    for field in MERGE_FIELDS:
        if not clean_text(result.get(field)) and clean_text(other.get(field)):
            result[field] = other.get(field)

    result["confidence"] = max(
        safe_float(a.get("confidence"), 0),
        safe_float(b.get("confidence"), 0),
    )

    methods = []
    for m in (clean_text(a.get("extraction_method")), clean_text(b.get("extraction_method"))):
        for part in m.split("+"):
            if part and part not in methods:
                methods.append(part)
    result["extraction_method"] = "+".join(methods)

    return result


# =============================================================================
# CLEAN ONE ROW
# =============================================================================

def clean_row(row):
    audit = {
        "excel_row": row.get("_excel_row"),
        "old_name": clean_text(row.get("name")),
        "old_email": clean_text(row.get("email")),
        "new_name": "",
        "new_email": "",
        "name_action": "",
        "email_action": "",
        "reason": "",
    }

    source_url = clean_text(row.get("source_url"))
    profile_url = clean_text(row.get("profile_url"))

    # Email mandatory.
    email, repaired, email_reason = normalize_and_validate_email(
        row.get("email"),
        source_url=source_url,
        profile_url=profile_url,
    )

    if email_reason or not email:
        audit["reason"] = email_reason or "invalid_email"
        return None, audit

    row["email"] = email
    audit["new_email"] = email
    audit["email_action"] = "repaired_domain_artifact" if repaired else "normalized"

    # Name selection.
    new_name, name_source, changed, name_reason = choose_best_name(
        row.get("name"),
        email,
        profile_url=profile_url,
    )
    row["name"] = new_name
    audit["new_name"] = new_name
    audit["name_action"] = name_source
    audit["reason"] = name_reason

    if changed:
        if name_source.startswith("email_strong"):
            row["extraction_method"] = append_method(
                row.get("extraction_method"),
                "cleaned_name_from_email_strong",
            )
        elif name_source == "email_local_weak":
            row["extraction_method"] = append_method(
                row.get("extraction_method"),
                "cleaned_name_from_email_weak",
            )
        elif name_source == "profile_url":
            row["extraction_method"] = append_method(
                row.get("extraction_method"),
                "cleaned_name_from_profile_url",
            )
        elif name_source == "shared_email_no_person_name":
            row["extraction_method"] = append_method(
                row.get("extraction_method"),
                "cleaned_shared_email_name_removed",
            )
        else:
            row["extraction_method"] = append_method(
                row.get("extraction_method"),
                "cleaned_name",
            )

    if repaired:
        row["extraction_method"] = append_method(
            row.get("extraction_method"),
            "repaired_email_domain_artifact",
        )

    # Email type.
    row["email_type"] = "shared/role" if is_shared_email(email) else "personal"

    # We only preserve explicit validated alternates.
    row["alternate_emails"] = clean_alternate_emails(
        row.get("alternate_emails"),
        email,
        source_url=source_url,
        profile_url=profile_url,
        email_conflict=row.get("email_conflict"),
    )
    row["email_conflict"] = "yes" if row["alternate_emails"] else "no"

    # Country.
    country, country_source = clean_country(
        row.get("country"),
        email,
        source_url,
    )
    row["country"] = country
    if country_source and (
        not clean_text(row.get("country_source"))
        or country_source != "existing_clean"
    ):
        row["country_source"] = country_source

    # General cleanup.
    for field in [
        "journal_name", "editorial_role", "academic_title", "academic_rank",
        "specialty", "affiliation", "university", "faculty", "school",
        "department", "institute", "division", "city", "address",
    ]:
        row[field] = clean_text(row.get(field))

    row["phone"] = clean_text(row.get("phone"))[:150]

    for field in [
        "orcid", "google_scholar", "scopus_author_id", "researcher_id",
        "pubmed", "profile_url", "personal_homepage", "source_url",
    ]:
        v = clean_text(row.get(field))
        row[field] = "" if v.lower() in {"none", "null", "n/a", "na", "-"} else v

    # Confidence caps for weak identity derivation.
    conf = safe_float(row.get("confidence"), 0)
    if name_source == "email_local_weak":
        row["confidence"] = min(conf or 70, 70)
    elif name_source == "shared_email_no_person_name":
        row["confidence"] = min(conf or 65, 65)

    row["scrape_status"] = "accepted"
    return row, audit


# =============================================================================
# READ / WRITE
# =============================================================================

def read_excel(path):
    wb = load_workbook(path)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]

    headers = [clean_text(cell.value) for cell in ws[1]]
    if EMAIL_COLUMN not in headers:
        raise RuntimeError(
            f"Required column '{EMAIL_COLUMN}' not found. Found columns: {headers}"
        )

    rows = []
    for excel_row, values in enumerate(
        ws.iter_rows(min_row=2, values_only=True),
        start=2,
    ):
        row = {}
        for i, header in enumerate(headers):
            if header:
                row[header] = values[i] if i < len(values) else None
        row["_excel_row"] = excel_row
        rows.append(row)

    return wb, ws, headers, rows


def dedupe_rows(rows):
    """
    Personal email:
        one best merged row per email.

    Shared/role email:
        preserve distinct person/profile mappings; do not collapse every person
        into one generic mailbox row.
    """
    merged = {}
    duplicate_count = 0

    for row in rows:
        email = normalize_email(row.get("email"))
        name_key = " ".join(tokenise_name(row.get("name")))
        profile = clean_text(row.get("profile_url")).lower()

        if is_shared_email(email) and PRESERVE_SHARED_EMAIL_PER_PERSON:
            if name_key:
                key = ("shared_person", email, name_key)
            elif profile:
                key = ("shared_profile", email, profile)
            else:
                key = ("shared_email", email)
        else:
            key = ("personal_email", email)

        if key not in merged:
            merged[key] = row
        else:
            duplicate_count += 1
            merged[key] = merge_duplicate_rows(merged[key], row)

    return list(merged.values()), duplicate_count


def style_sheet(ws, headers):
    fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    font = Font(bold=True, color="FFFFFF")

    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "name": 32, "email": 42, "country": 20, "alternate_emails": 48,
        "email_type": 16, "email_conflict": 16, "employee_name": 20,
        "employee_email": 36, "page_type": 24, "journal_name": 38,
        "editorial_role": 32, "academic_title": 42, "academic_rank": 28,
        "specialty": 40, "affiliation": 65, "university": 45, "faculty": 42,
        "school": 42, "department": 45, "institute": 45, "division": 40,
        "city": 22, "address": 60, "phone": 25, "orcid": 35,
        "google_scholar": 50, "scopus_author_id": 30, "researcher_id": 30,
        "pubmed": 50, "profile_url": 70, "personal_homepage": 70,
        "source_url": 75, "country_source": 28, "confidence": 15,
        "extraction_method": 75, "scrape_status": 18,
    }

    for col_num, header in enumerate(headers, start=1):
        letter = get_column_letter(col_num)
        ws.column_dimensions[letter].width = widths.get(header, 24)


def write_audit_sheet(wb, audits):
    if "Cleaning Audit" in wb.sheetnames:
        del wb["Cleaning Audit"]

    ws = wb.create_sheet("Cleaning Audit")
    headers = [
        "excel_row", "old_name", "old_email", "new_name", "new_email",
        "name_action", "email_action", "reason",
    ]
    ws.append(headers)

    for item in audits:
        ws.append([item.get(h, "") for h in headers])

    style_sheet(ws, headers)


def write_rejected_sheet(wb, rejected):
    if "Rejected Rows" in wb.sheetnames:
        del wb["Rejected Rows"]

    ws = wb.create_sheet("Rejected Rows")
    headers = ["excel_row", "name", "email", "reason", "source_url", "profile_url"]
    ws.append(headers)

    for item in rejected:
        ws.append([item.get(h, "") for h in headers])

    style_sheet(ws, headers)


# =============================================================================
# MAIN
# =============================================================================

def clean_workbook(input_file, output_dir, output_filename):
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path.resolve()}")

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / output_filename

    wb, ws, headers, raw_rows = read_excel(input_path)

    print("=" * 100)
    print("ROBUST UNIVERSITY / SCIENTIFIC EMAIL CLEANER V2")
    print("=" * 100)
    print(f"Input : {input_path}")
    print(f"Rows  : {len(raw_rows)}")
    print()

    cleaned = []
    rejected = []
    audits = []

    repaired_emails = 0
    strong_email_names = 0
    weak_email_names = 0
    profile_names = 0
    bad_names_removed = 0
    shared_emails = 0

    for row in raw_rows:
        result, audit = clean_row(row.copy())
        audits.append(audit)

        if result is None:
            rejected.append({
                "excel_row": row.get("_excel_row"),
                "name": clean_text(row.get("name")),
                "email": clean_text(row.get("email")),
                "reason": audit.get("reason", "invalid_email"),
                "source_url": clean_text(row.get("source_url")),
                "profile_url": clean_text(row.get("profile_url")),
            })
            continue

        if audit.get("email_action") == "repaired_domain_artifact":
            repaired_emails += 1

        action = audit.get("name_action")
        if action and action.startswith("email_strong"):
            strong_email_names += 1
        elif action == "email_local_weak":
            weak_email_names += 1
        elif action == "profile_url":
            profile_names += 1
        elif action == "shared_email_no_person_name":
            bad_names_removed += 1

        if result.get("email_type") == "shared/role":
            shared_emails += 1

        cleaned.append(result)

    cleaned, duplicate_count = dedupe_rows(cleaned)

    cleaned.sort(
        key=lambda r: (
            clean_text(r.get("name")).casefold(),
            clean_text(r.get("email")).casefold(),
        )
    )

    # Replace old data.
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for row in cleaned:
        ws.append([
            "" if row.get(header) is None else row.get(header, "")
            for header in headers
        ])

    style_sheet(ws, headers)

    if WRITE_CLEANING_AUDIT:
        write_audit_sheet(wb, audits)

    if WRITE_REJECTED_ROWS:
        write_rejected_sheet(wb, rejected)

    wb.save(output_path)

    unique_emails = len({normalize_email(r.get("email")) for r in cleaned})
    blank_names = sum(1 for r in cleaned if not clean_text(r.get("name")))

    print("=" * 100)
    print("CLEANING SUMMARY")
    print("=" * 100)
    print(f"Original rows                    : {len(raw_rows)}")
    print(f"Invalid/junk emails removed      : {len(rejected)}")
    print(f"Email-domain artifacts repaired  : {repaired_emails}")
    print(f"Duplicate observations merged    : {duplicate_count}")
    print(f"Final rows                       : {len(cleaned)}")
    print(f"Unique emails                    : {unique_emails}")
    print(f"Strong names taken from email    : {strong_email_names}")
    print(f"Weak compact-email names used    : {weak_email_names}")
    print(f"Names recovered from profile URL : {profile_names}")
    print(f"Bad shared-email names removed   : {bad_names_removed}")
    print(f"Shared/role email rows            : {shared_emails}")
    print(f"Blank names remaining            : {blank_names}")
    print()
    print(f"Saved: {output_path.resolve()}")
    print("=" * 100)

    return output_path


if __name__ == "__main__":
    try:
        clean_workbook(
            input_file=INPUT_FILE,
            output_dir=OUTPUT_DIR,
            output_filename=OUTPUT_FILENAME,
        )
    except Exception as exc:
        print()
        print("=" * 100)
        print("ERROR")
        print("=" * 100)
        print(str(exc))
        print()
        raise