from __future__ import annotations

import argparse
import concurrent.futures as cf
import logging
import random
import re
import threading
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

INDEX_URL = "https://www.gfmer.ch/International_activities_En/GFMER_members_list.htm"
DEFAULT_OUTPUT = "output/Sobha/gfmer_data/gfmer_members.xlsx"
DEFAULT_CHECKPOINT = "output/Sobha/gfmer_data/gfmer_members_checkpoint.csv"
TIMEOUT = 45
MAX_WORKERS = 8
REQUEST_DELAY_SECONDS = 0.15

EMAIL_RE = re.compile(r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}", re.I)
YEAR_RE = re.compile(r"/SRH-Course-(\d{4})/", re.I)
GENERIC_EMAIL_PREFIXES = {"info","contact","office","admin","support","webmaster","secretariat"}

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("gfmer")
_thread_local = threading.local()


def clean(value: Any) -> str:
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ").replace("\u200b", "").replace("\ufeff", "")
    return re.sub(r"\s+", " ", value).strip()


def normalize_email(value: Any) -> str:
    text = clean(value)
    if not text:
        return ""
    if text.lower().startswith("mailto:"):
        text = text[7:]
    text = text.split("?", 1)[0]
    m = EMAIL_RE.search(text)
    return m.group(0).lower() if m else ""


def is_person_email(email: str) -> bool:
    email = normalize_email(email)
    if not email:
        return False
    return email.split("@", 1)[0].lower() not in GENERIC_EMAIL_PREFIXES


def get_session() -> requests.Session:
    session = getattr(_thread_local, "session", None)
    if session is not None:
        return session
    session = requests.Session()
    retry = Retry(total=5, connect=5, read=5, status=5, backoff_factor=0.7,
                  status_forcelist=(429, 500, 502, 503, 504),
                  allowed_methods=frozenset(["GET"]), raise_on_status=False)
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/154.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9",
    })
    _thread_local.session = session
    return session


def get_html(url: str) -> str:
    response = get_session().get(url, timeout=TIMEOUT)
    if response.status_code >= 400:
        raise RuntimeError(f"HTTP {response.status_code}: {url}")
    if not response.encoding or response.encoding.lower() == "iso-8859-1":
        response.encoding = response.apparent_encoding or "utf-8"
    time.sleep(REQUEST_DELAY_SECONDS)
    return response.text


@dataclass
class IndexMember:
    member_text: str
    city: str
    country: str
    profile_url: str


@dataclass
class MemberRecord:
    index_member_text: str = ""
    index_city: str = ""
    index_country: str = ""
    name: str = ""
    credentials: str = ""
    title: str = ""
    organization: str = ""
    department: str = ""
    affiliation: str = ""
    city: str = ""
    country: str = ""
    email: str = ""
    phone: str = ""
    website: str = ""
    image_url: str = ""
    course_year: str = ""
    profile_url: str = ""
    source_url: str = ""
    biography: str = ""
    interests: str = ""
    education: str = ""
    additional_links: str = ""
    status: str = "OK"
    error: str = ""


def parse_member_index(html: str) -> list[IndexMember]:
    """
    Parse the old GFMER member table robustly.

    Important:
    Do NOT rely on `tbody tr`. The page uses legacy HTML and different parsers
    may normalize/move <tr> elements. We therefore inspect all rows below
    table#example and require at least:
        member/profile link | city | country
    """
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    table = soup.select_one(
        "table#example"
    )

    if not table:
        # Diagnostic fallback in case the ID is altered by malformed markup.
        candidates = soup.find_all("table")

        log.error(
            "table#example not found. Total tables in page: %s",
            len(candidates),
        )

        raise RuntimeError(
            "Could not find table#example on the GFMER member directory page."
        )

    all_rows = table.find_all("tr")

    log.info(
        "Raw table rows found: %s",
        len(all_rows),
    )

    members: list[IndexMember] = []
    seen_urls: set[str] = set()

    for row_no, row in enumerate(
        all_rows,
        start=1,
    ):
        cells = row.find_all(
            "td",
            recursive=False,
        )

        # Some malformed HTML causes td nesting. Use non-recursive first,
        # then fall back to regular find_all.
        if len(cells) < 3:
            cells = row.find_all(
                "td"
            )

        if len(cells) < 3:
            continue

        link = cells[0].find(
            "a",
            href=True,
        )

        if not link:
            continue

        href = clean(
            link.get("href")
        )

        if not href:
            continue

        profile_url = urljoin(
            INDEX_URL,
            href,
        )

        member_text = clean(
            link.get_text(
                " ",
                strip=True,
            )
        )

        city = clean(
            cells[1].get_text(
                " ",
                strip=True,
            )
        )

        country = clean(
            cells[2].get_text(
                " ",
                strip=True,
            )
        )

        # Keep only useful profile-like links from this directory.
        if not member_text:
            continue

        if profile_url in seen_urls:
            continue

        seen_urls.add(
            profile_url
        )

        members.append(
            IndexMember(
                member_text=member_text,
                city=city,
                country=country,
                profile_url=profile_url,
            )
        )

    log.info(
        "Valid member/profile rows parsed: %s",
        len(members),
    )

    if not members:
        # Write a diagnostic HTML copy so you can inspect exactly what your
        # machine received from GFMER.
        diagnostic_path = Path(
            "gfmer_index_debug.html"
        )

        try:
            diagnostic_path.write_text(
                html,
                encoding="utf-8",
                errors="ignore",
            )

            log.error(
                "No member rows parsed. Saved received HTML to: %s",
                diagnostic_path.resolve(),
            )
        except Exception:
            pass

        raise RuntimeError(
            "GFMER directory page was downloaded, but zero member rows were "
            "parsed. Check gfmer_index_debug.html; the site may have returned "
            "a different/interstitial page."
        )

    return members


def get_main_content(soup: BeautifulSoup) -> Tag:
    for selector in ["#content", "#Content", "#maincontent", "#main-content", "main", "body"]:
        node = soup.select_one(selector)
        if node:
            return node
    return soup


def all_emails(content: Tag) -> list[str]:
    found = []
    for a in content.select('a[href^="mailto:"]'):
        email = normalize_email(a.get("href") or a.get_text(" ", strip=True))
        if email and email not in found:
            found.append(email)
    for email in EMAIL_RE.findall(content.get_text(" ", strip=True)):
        email = normalize_email(email)
        if email and email not in found:
            found.append(email)
    return found


def extract_primary_email(content: Tag) -> str:
    emails = all_emails(content)
    for email in emails:
        if is_person_email(email):
            return email
    return emails[0] if emails else ""


def likely_name_from_index(member_text: str) -> str:
    return clean(member_text.split(" - ", 1)[0] if " - " in member_text else member_text)


def likely_affiliation_from_index(member_text: str) -> str:
    return clean(member_text.split(" - ", 1)[1] if " - " in member_text else "")


def get_title_paragraphs(content: Tag) -> list[str]:
    values = []
    for selector in ["p.TitreBleu", "p.titrebleu", ".TitreBleu", ".Titre2", "h1", "h2"]:
        for node in content.select(selector):
            text = clean(node.get_text(" ", strip=True))
            if text and text not in values:
                values.append(text)
    return values


def paragraph_with_email(content: Tag) -> Tag | None:
    for a in content.select('a[href^="mailto:"]'):
        current = a
        while current:
            if isinstance(current, Tag) and current.name in {"p", "div", "td", "li"}:
                text = clean(current.get_text(" ", strip=True))
                if EMAIL_RE.search(text):
                    return current
            current = current.parent
    return None


def lines_from_element(element: Tag | None) -> list[str]:
    if not element:
        return []
    soup = BeautifulSoup(str(element), "lxml")
    root = soup.body or soup
    for br in root.find_all("br"):
        br.replace_with("\n")
    return [clean(line) for line in root.get_text("\n", strip=True).splitlines() if clean(line)]


def split_name_credentials(text: str) -> tuple[str, str]:
    parts = [clean(x) for x in clean(text).split(",")]
    if len(parts) <= 1:
        return clean(text), ""
    return parts[0], ", ".join(p for p in parts[1:] if p)


def infer_title_org_from_lines(lines: list[str], name: str, email: str) -> tuple[str, str]:
    useful = []
    for line in lines:
        if EMAIL_RE.search(line):
            continue
        if name and clean(line).casefold() == clean(name).casefold():
            continue
        useful.append(line)
    if not useful:
        return "", ""
    first = useful[0]
    if "," in first:
        left, right = first.split(",", 1)
        return clean(left), clean(right)
    return first, ""


def infer_department(text: str) -> str:
    for pattern in [r"(Department of [^.;|]+)", r"(School of [^.;|]+)", r"(Faculty of [^.;|]+)", r"(College of [^.;|]+)", r"(Institute of [^.;|]+)"]:
        m = re.search(pattern, text, re.I)
        if m:
            return clean(m.group(1))
    return ""


def extract_phone(content: Tag) -> str:
    text = content.get_text(" ", strip=True)
    m = re.search(r"(?:Phone|Tel(?:ephone)?|Mobile)\s*:?\s*(\+?\d[\d\s()./\-]{6,}\d)", text, re.I)
    return clean(m.group(1)) if m else ""


def extract_image_url(content: Tag, profile_url: str) -> str:
    images = content.select("img[src]")
    for img in images:
        if clean(img.get("alt") or img.get("title")) and clean(img.get("src")):
            return urljoin(profile_url, clean(img.get("src")))
    for img in images:
        if clean(img.get("src")):
            return urljoin(profile_url, clean(img.get("src")))
    return ""


def extract_website(content: Tag, profile_url: str) -> str:
    profile_host = urlparse(profile_url).netloc.lower()
    for a in content.select("a[href]"):
        href = clean(a.get("href"))
        if not href or href.startswith(("mailto:", "#", "javascript:")):
            continue
        absolute = urljoin(profile_url, href)
        parsed = urlparse(absolute)
        if parsed.scheme not in {"http", "https"}:
            continue
        host = parsed.netloc.lower()
        if host and host != profile_host and "gfmer.ch" not in host:
            return absolute
    return ""


def extract_links(content: Tag, profile_url: str) -> str:
    links, seen = [], set()
    for a in content.select("a[href]"):
        href = clean(a.get("href"))
        if not href or href.startswith(("mailto:", "#", "javascript:")):
            continue
        absolute = urljoin(profile_url, href)
        if absolute == profile_url or absolute in seen:
            continue
        seen.add(absolute)
        label = clean(a.get_text(" ", strip=True))
        links.append(f"{label} -> {absolute}" if label else absolute)
    return " | ".join(links[:30])


def extract_biography(content: Tag) -> str:
    paragraphs = []
    for p in content.find_all("p"):
        text = clean(p.get_text(" ", strip=True))
        if not text:
            continue
        classes = {str(c).casefold() for c in p.get("class", [])}
        if {"titrebleu", "titre2"} & classes:
            continue
        if EMAIL_RE.search(text) and len(text) < 300:
            continue
        if len(text) >= 80:
            paragraphs.append(text)
    return "\n\n".join(paragraphs[:20])


def extract_interests_and_education(biography: str) -> tuple[str, str]:
    interests = education = ""
    if biography:
        m = re.search(r"(?:interests?|areas? of interest)\s+(?:include|are|:)\s*([^.\n]{10,300})", biography, re.I)
        if m:
            interests = clean(m.group(1))
        m = re.search(r"(?:obtained|holds?|earned|graduated|degree|masters?|ph\.?d\.?)[^.]{0,350}", biography, re.I)
        if m:
            education = clean(m.group(0))
    return interests, education


def extract_course_year(profile_url: str) -> str:
    m = YEAR_RE.search(profile_url)
    return m.group(1) if m else ""


def parse_profile(index_member: IndexMember, html: str) -> MemberRecord:
    soup = BeautifulSoup(html, "lxml")
    content = get_main_content(soup)
    index_name = likely_name_from_index(index_member.member_text)
    index_affiliation = likely_affiliation_from_index(index_member.member_text)
    title_paragraphs = get_title_paragraphs(content)

    name = ""
    for value in title_paragraphs:
        if index_name and (index_name.casefold() in value.casefold() or value.casefold() in index_name.casefold()):
            name = value
            break
    if not name and title_paragraphs:
        name = title_paragraphs[0]
    if not name:
        name = index_name

    email = extract_primary_email(content)
    contact_lines = lines_from_element(paragraph_with_email(content))
    contact_name = ""
    for line in contact_lines:
        if not EMAIL_RE.search(line):
            contact_name = line
            break

    credentials = ""
    if contact_name:
        parsed_name, credentials = split_name_credentials(contact_name)
        if parsed_name:
            name = parsed_name

    title, organization = infer_title_org_from_lines(contact_lines[1:] if len(contact_lines) > 1 else contact_lines, name, email)
    affiliation = title_paragraphs[1] if len(title_paragraphs) >= 2 else index_affiliation
    if not organization and affiliation:
        organization = affiliation

    biography = extract_biography(content)
    interests, education = extract_interests_and_education(biography)
    department = infer_department(clean(" ".join([affiliation, biography])))

    return MemberRecord(
        index_member_text=index_member.member_text,
        index_city=index_member.city,
        index_country=index_member.country,
        name=name,
        credentials=credentials,
        title=title,
        organization=organization,
        department=department,
        affiliation=affiliation,
        city=index_member.city,
        country=index_member.country,
        email=email,
        phone=extract_phone(content),
        website=extract_website(content, index_member.profile_url),
        image_url=extract_image_url(content, index_member.profile_url),
        course_year=extract_course_year(index_member.profile_url),
        profile_url=index_member.profile_url,
        source_url=index_member.profile_url,
        biography=biography,
        interests=interests,
        education=education,
        additional_links=extract_links(content, index_member.profile_url),
        status="OK",
        error="",
    )


def scrape_one(index_member: IndexMember) -> MemberRecord:
    try:
        return parse_profile(index_member, get_html(index_member.profile_url))
    except Exception as exc:
        return MemberRecord(
            index_member_text=index_member.member_text,
            index_city=index_member.city,
            index_country=index_member.country,
            name=likely_name_from_index(index_member.member_text),
            affiliation=likely_affiliation_from_index(index_member.member_text),
            city=index_member.city,
            country=index_member.country,
            profile_url=index_member.profile_url,
            source_url=index_member.profile_url,
            course_year=extract_course_year(index_member.profile_url),
            status="ERROR",
            error=str(exc)[:2000],
        )


def save_checkpoint(records: list[MemberRecord], checkpoint_path: Path) -> None:
    checkpoint_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    columns = list(
        MemberRecord.__dataclass_fields__.keys()
    )

    df = pd.DataFrame(
        [asdict(r) for r in records],
        columns=columns,
    )

    df.to_csv(
        checkpoint_path,
        index=False,
        encoding="utf-8-sig",
    )


def load_checkpoint(checkpoint_path: Path) -> dict[str, MemberRecord]:
    if not checkpoint_path.exists():
        return {}
    try:
        df = pd.read_csv(checkpoint_path, dtype=str, keep_default_na=False)
        result = {}
        allowed = set(MemberRecord.__dataclass_fields__)
        for row in df.to_dict(orient="records"):
            rec = MemberRecord(**{k: row.get(k, "") for k in allowed})
            if rec.profile_url:
                result[rec.profile_url] = rec
        return result
    except Exception as exc:
        log.warning("Could not load checkpoint: %s", exc)
        return {}


def write_sheet(writer: pd.ExcelWriter, df: pd.DataFrame, sheet_name: str) -> None:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.freeze_panes = "A2"
    if ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    for column_cells in ws.columns:
        letter = column_cells[0].column_letter
        max_length = 0
        for cell in column_cells[:500]:
            max_length = max(max_length, len(clean(cell.value)))
        ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 70)


def export_excel(records: list[MemberRecord], output_path: Path) -> None:
    df = pd.DataFrame([asdict(r) for r in records])
    if not df.empty:
        df = df.sort_values(by=["country", "name"], na_position="last").reset_index(drop=True)
    errors_df = df[df["status"] == "ERROR"].copy() if not df.empty else pd.DataFrame()
    email_df = df[df["email"].fillna("").astype(str).str.strip() != ""].copy() if not df.empty else pd.DataFrame()
    unique_email_df = email_df.drop_duplicates(subset=["email"], keep="first").copy() if not email_df.empty else pd.DataFrame()
    summary_df = pd.DataFrame([
        {"metric": "Total directory members", "value": len(df)},
        {"metric": "Profiles with email", "value": len(email_df)},
        {"metric": "Unique emails", "value": len(unique_email_df)},
        {"metric": "Profile errors", "value": len(errors_df)},
        {"metric": "Countries", "value": df["country"].nunique() if not df.empty else 0},
    ])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        write_sheet(writer, df, "All Members")
        write_sheet(writer, unique_email_df, "Unique Emails")
        write_sheet(writer, errors_df, "Errors")
        write_sheet(writer, summary_df, "Summary")
    log.info("Excel saved: %s", output_path.resolve())


def main():
    parser = argparse.ArgumentParser(description="Scrape the complete GFMER member directory and visit each public member/profile page.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT)
    parser.add_argument("--checkpoint", default=DEFAULT_CHECKPOINT)
    parser.add_argument("--workers", type=int, default=MAX_WORKERS)
    parser.add_argument("--limit", type=int, default=0, help="Optional test limit; 0 means all members.")
    parser.add_argument("--fresh", action="store_true", help="Ignore existing checkpoint and scrape every profile again.")
    args = parser.parse_args()

    output_path = Path(args.output)
    checkpoint_path = Path(args.checkpoint)

    # Create both directories before any checkpoint/export operation.
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    checkpoint_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    log.info("Output directory     : %s", output_path.parent.resolve())
    log.info("Checkpoint directory : %s", checkpoint_path.parent.resolve())

    log.info("Downloading member directory...")
    members = parse_member_index(get_html(INDEX_URL))
    if args.limit > 0:
        members = members[:args.limit]
    log.info("Directory members found: %s", len(members))

    if not members:
        raise RuntimeError(
            "No GFMER members were discovered. Aborting instead of writing "
            "an empty checkpoint/Excel."
        )

    completed = {} if args.fresh else load_checkpoint(checkpoint_path)
    if completed:
        log.info("Checkpoint loaded: %s profiles", len(completed))

    pending = [m for m in members if m.profile_url not in completed or completed[m.profile_url].status == "ERROR"]
    log.info("Profiles remaining: %s", len(pending))

    if pending:
        workers = max(1, min(args.workers, 20))
        with cf.ThreadPoolExecutor(max_workers=workers) as executor:
            future_to_member = {executor.submit(scrape_one, m): m for m in pending}
            processed_now = 0
            for future in cf.as_completed(future_to_member):
                member = future_to_member[future]
                record = future.result()
                completed[member.profile_url] = record
                processed_now += 1
                log.info("[%s/%s] %s | %s | %s", processed_now, len(pending), record.name or member.member_text, record.email or "NO EMAIL", record.status)
                if processed_now % 25 == 0:
                    save_checkpoint([completed[m.profile_url] for m in members if m.profile_url in completed], checkpoint_path)

    final_records = [
        completed.get(m.profile_url, MemberRecord(
            index_member_text=m.member_text,
            index_city=m.city,
            index_country=m.country,
            name=likely_name_from_index(m.member_text),
            affiliation=likely_affiliation_from_index(m.member_text),
            city=m.city,
            country=m.country,
            profile_url=m.profile_url,
            source_url=m.profile_url,
            status="ERROR",
            error="Profile was not processed.",
        ))
        for m in members
    ]

    save_checkpoint(final_records, checkpoint_path)
    export_excel(final_records, output_path)

    total = len(final_records)
    with_email = sum(bool(r.email) for r in final_records)
    unique_emails = len({r.email for r in final_records if r.email})
    errors = sum(r.status == "ERROR" for r in final_records)

    print()
    print("=" * 76)
    print("GFMER MEMBER SCRAPE COMPLETE")
    print("=" * 76)
    print(f"Members          : {total}")
    print(f"With email       : {with_email}")
    print(f"Unique emails    : {unique_emails}")
    print(f"Errors           : {errors}")
    print(f"Checkpoint       : {checkpoint_path.resolve()}")
    print(f"Excel            : {output_path.resolve()}")
    print("=" * 76)


if __name__ == "__main__":
    main()
