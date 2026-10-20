from __future__ import annotations

import argparse
import concurrent.futures as cf
import io
import logging
import re
import threading
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None

try:
    import pycountry
except Exception:
    pycountry = None


DEFAULT_INPUT = "longdom_input_urls.xlsx"
DEFAULT_OUTPUT = "output/longdom/longdom_authors_emails.xlsx"
DEFAULT_CHECKPOINT = "output/longdom/longdom_authors_checkpoint.csv"

TIMEOUT = 60
MAX_WORKERS = 6
REQUEST_DELAY_SECONDS = 0.10

LONGDOM_HOSTS = {
    "longdom.org",
    "www.longdom.org",
}

EMAIL_RE = re.compile(
    r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
    re.I,
)

YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")

GENERIC_EMAILS = {
    "info@longdom.org",
    "editor@longdom.org",
    "contact@longdom.org",
    "support@longdom.org",
}

COUNTRY_ALIASES = {
    "usa": "United States",
    "u.s.a.": "United States",
    "u.s.a": "United States",
    "united states of america": "United States",
    "us": "United States",
    "uk": "United Kingdom",
    "u.k.": "United Kingdom",
    "uae": "United Arab Emirates",
    "u.a.e.": "United Arab Emirates",
    "south korea": "South Korea",
    "republic of korea": "South Korea",
    "north korea": "North Korea",
    "iran": "Iran",
    "russia": "Russia",
    "vietnam": "Vietnam",
    "viet nam": "Vietnam",
    "czech republic": "Czech Republic",
    "czechia": "Czech Republic",
    "taiwan": "Taiwan",
    "hong kong": "Hong Kong",
    "palestine": "Palestine",
}

_thread_local = threading.local()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

log = logging.getLogger("longdom_html_pdf")


@dataclass
class ListingLink:
    url: str
    year_hint: str = ""


@dataclass
class AuthorRecord:
    author: str = ""
    email: str = ""
    affiliation: str = ""
    country: str = ""
    year: str = ""

    # internal only; removed from final Excel
    listing_url: str = ""
    proceeding_url: str = ""
    pdf_url: str = ""
    status: str = "OK"
    error: str = ""


def clean(value: Any) -> str:
    if value is None:
        return ""

    value = str(value)
    value = (
        value.replace("\xa0", " ")
        .replace("\u200b", "")
        .replace("\ufeff", "")
    )

    return re.sub(r"\s+", " ", value).strip()


def normalize_email(value: Any) -> str:
    value = clean(value)

    if not value:
        return ""

    match = EMAIL_RE.search(value)

    if not match:
        return ""

    email = match.group(0).lower().strip(".,;:()[]{}<>")

    if email in GENERIC_EMAILS:
        return ""

    if email.endswith("@longdom.org"):
        return ""

    return email


def get_session() -> requests.Session:
    session = getattr(_thread_local, "session", None)

    if session is not None:
        return session

    session = requests.Session()

    retry = Retry(
        total=5,
        connect=5,
        read=5,
        status=5,
        backoff_factor=0.8,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET"]),
        raise_on_status=False,
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=20,
        pool_maxsize=20,
    )

    session.mount("https://", adapter)
    session.mount("http://", adapter)

    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/155.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9",
        }
    )

    _thread_local.session = session

    return session


def get_response(url: str, referer: str = "") -> requests.Response:
    session = get_session()

    headers = {}

    if referer:
        headers["Referer"] = referer

    response = session.get(
        url,
        headers=headers,
        timeout=TIMEOUT,
    )

    if response.status_code >= 400:
        raise RuntimeError(
            f"HTTP {response.status_code}: {url}"
        )

    time.sleep(REQUEST_DELAY_SECONDS)

    return response


def get_html(url: str, referer: str = "") -> str:
    response = get_response(url, referer=referer)

    if not response.encoding:
        response.encoding = response.apparent_encoding or "utf-8"

    return response.text


# ============================================================
# INPUT EXCEL
# ============================================================

def year_from_text(*values: str) -> str:
    for value in values:
        match = YEAR_RE.search(clean(value))

        if match:
            return match.group(1)

    return ""


def read_input_urls(input_path: Path) -> list[ListingLink]:
    if not input_path.exists():
        raise FileNotFoundError(
            f"Input Excel not found: {input_path.resolve()}"
        )

    df = pd.read_excel(
        input_path,
        dtype=str,
    ).fillna("")

    if df.empty:
        raise RuntimeError("Input Excel is empty.")

    normalized_columns = {
        clean(column).casefold(): column
        for column in df.columns
    }

    url_column = None

    for candidate in [
        "url",
        "link",
        "source_url",
        "listing_url",
        "conference_url",
    ]:
        if candidate in normalized_columns:
            url_column = normalized_columns[candidate]
            break

    if url_column is None:
        for column in df.columns:
            values = [clean(x) for x in df[column].tolist()]

            if any(
                "longdom.org" in value.lower()
                for value in values
            ):
                url_column = column
                break

    if url_column is None:
        raise RuntimeError(
            "Could not find URL column. Recommended header: url"
        )

    year_column = None

    if "year" in normalized_columns:
        year_column = normalized_columns["year"]

    rows = []
    seen = set()

    for _, row in df.iterrows():
        url = clean(row.get(url_column, ""))

        if not url.startswith(("http://", "https://")):
            continue

        parsed = urlparse(url)

        if parsed.netloc.lower() not in LONGDOM_HOSTS:
            continue

        if url in seen:
            continue

        seen.add(url)

        hint = ""

        if year_column:
            hint = year_from_text(row.get(year_column, ""))

        if not hint:
            hint = year_from_text(url)

        rows.append(
            ListingLink(
                url=url,
                year_hint=hint,
            )
        )

    log.info(
        "Input Excel rows=%s | valid unique Longdom URLs=%s",
        len(df),
        len(rows),
    )

    return rows


# ============================================================
# LISTING -> PROCEEDINGS
# ============================================================

def is_proceeding_url(url: str) -> bool:
    parsed = urlparse(url)

    return (
        parsed.netloc.lower() in LONGDOM_HOSTS
        and "/proceedings/" in parsed.path.lower()
        and parsed.path.lower().endswith(".html")
    )


def discover_proceedings(
    listing: ListingLink,
) -> list[str]:
    html = get_html(listing.url)

    soup = BeautifulSoup(html, "lxml")

    found = []
    seen = set()

    for anchor in soup.select("a[href]"):
        href = urljoin(
            listing.url,
            clean(anchor.get("href")),
        )

        if not is_proceeding_url(href):
            continue

        if href in seen:
            continue

        seen.add(href)
        found.append(href)

    log.info(
        "DISCOVER | %s | proceedings=%s",
        listing.url,
        len(found),
    )

    return found


# ============================================================
# HTML EXTRACTION
# ============================================================

def extract_pdf_url(
    soup: BeautifulSoup,
    proceeding_url: str,
) -> str:
    # Prefer anchor whose visible text is PDF.
    for anchor in soup.select("a[href]"):
        text = clean(
            anchor.get_text(" ", strip=True)
        ).casefold()

        href = clean(anchor.get("href"))

        if not href:
            continue

        absolute = urljoin(
            proceeding_url,
            href,
        )

        if text == "pdf" or absolute.lower().endswith(".pdf"):
            return absolute

    return ""


def extract_author_affiliation_from_html(
    soup: BeautifulSoup,
) -> tuple[str, str]:
    """
    Longdom proceeding HTML commonly has:
        title
        conference
        date/location
        AUTHOR
        AFFILIATION
        Keynote:/Scientific Tracks...
        Abstract:
    """
    abstract_heading = None

    for node in soup.find_all(
        ["h4", "h5", "h6", "strong"]
    ):
        text = clean(
            node.get_text(" ", strip=True)
        ).casefold()

        if text.startswith("abstract"):
            abstract_heading = node
            break

    if abstract_heading:
        previous = []
        node = abstract_heading.previous_sibling

        while node and len(previous) < 18:
            if isinstance(node, Tag):
                value = clean(
                    node.get_text(" ", strip=True)
                )

                if value:
                    previous.append(value)

            node = node.previous_sibling

        previous.reverse()

        candidates = []

        for value in previous:
            lower = value.casefold()

            if not value:
                continue

            if lower in {"html", "pdf"}:
                continue

            if lower.startswith(
                (
                    "keynote:",
                    "scientific tracks abstracts:",
                    "posters:",
                    "accepted abstracts:",
                )
            ):
                continue

            if "conference" in lower and len(value) < 200:
                continue

            if re.search(
                r"[A-Za-z]+\s+\d{1,2}(?:-\d{1,2})?,?\s+\d{4}\s*\|",
                value,
            ):
                continue

            candidates.append(value)

        # The last two meaningful blocks before category/Abstract are usually
        # author and affiliation.
        if len(candidates) >= 2:
            return (
                clean(candidates[-2]),
                clean(candidates[-1]),
            )

    # Fallback based on the body line sequence around Abstract.
    lines = [
        clean(line)
        for line in soup.get_text("\n", strip=True).splitlines()
        if clean(line)
    ]

    abstract_idx = None

    for i, line in enumerate(lines):
        if line.casefold().startswith("abstract"):
            abstract_idx = i
            break

    if abstract_idx is not None:
        before = lines[max(0, abstract_idx - 12):abstract_idx]

        filtered = []

        for value in before:
            lower = value.casefold()

            if lower in {"pdf", "html"}:
                continue

            if lower.startswith(
                (
                    "keynote:",
                    "scientific tracks abstracts:",
                    "posters:",
                    "accepted abstracts:",
                )
            ):
                continue

            if "conference" in lower and len(value) < 200:
                continue

            if " | " in value and YEAR_RE.search(value):
                continue

            filtered.append(value)

        if len(filtered) >= 2:
            return clean(filtered[-2]), clean(filtered[-1])

    return "", ""


def extract_email_from_html(
    soup: BeautifulSoup,
) -> str:
    for anchor in soup.select('a[href^="mailto:"]'):
        email = normalize_email(
            anchor.get("href")
        )

        if email:
            return email

    for email in EMAIL_RE.findall(
        soup.get_text(" ", strip=True)
    ):
        email = normalize_email(email)

        if email:
            return email

    return ""


# ============================================================
# PDF DOWNLOAD + EXTRACTION
# ============================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> str:
    if not pdf_bytes:
        return ""

    if fitz is not None:
        try:
            document = fitz.open(
                stream=pdf_bytes,
                filetype="pdf",
            )

            pages = []

            for page in document:
                pages.append(
                    page.get_text("text")
                )

            document.close()

            text = "\n".join(pages)

            if clean(text):
                return text

        except Exception as exc:
            log.debug(
                "PyMuPDF extraction failed: %s",
                exc,
            )

    if PdfReader is not None:
        try:
            reader = PdfReader(
                io.BytesIO(pdf_bytes)
            )

            return "\n".join(
                page.extract_text() or ""
                for page in reader.pages
            )

        except Exception as exc:
            log.debug(
                "pypdf extraction failed: %s",
                exc,
            )

    return ""


def download_pdf(
    pdf_url: str,
    proceeding_url: str,
) -> bytes:
    response = get_response(
        pdf_url,
        referer=proceeding_url,
    )

    content_type = response.headers.get(
        "Content-Type",
        "",
    ).lower()

    if (
        "pdf" not in content_type
        and not response.content.startswith(b"%PDF")
    ):
        raise RuntimeError(
            f"PDF link returned non-PDF content: {pdf_url}"
        )

    return response.content


def extract_email_from_pdf_text(
    pdf_text: str,
) -> str:
    if not pdf_text:
        return ""

    # Longdom PDFs frequently format correspondence as:
    #   e: person@example.com
    # Give that highest priority.
    patterns = [
        r"\be\s*:\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})",
        r"\bemail\s*:\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})",
        r"\be-mail\s*:\s*([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})",
        EMAIL_RE.pattern,
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            pdf_text,
            re.I,
        )

        if not match:
            continue

        value = (
            match.group(1)
            if match.lastindex
            else match.group(0)
        )

        email = normalize_email(value)

        if email:
            return email

    return ""


def extract_author_affiliation_from_pdf(
    pdf_text: str,
    author_html: str,
    affiliation_html: str,
) -> tuple[str, str]:
    """
    PDF footer commonly contains:
        e: email@example.com
        Author Name
        University Name, USA
    """
    author = clean(author_html)
    affiliation = clean(affiliation_html)

    if not pdf_text:
        return author, affiliation

    lines = [
        clean(line)
        for line in pdf_text.splitlines()
        if clean(line)
    ]

    email_line_index = None

    for i, line in enumerate(lines):
        if EMAIL_RE.search(line):
            email_line_index = i
            break

    if email_line_index is not None:
        after = lines[
            email_line_index + 1:
            email_line_index + 5
        ]

        if not author and after:
            candidate = after[0]

            if len(candidate) < 150:
                author = candidate

        if not affiliation and len(after) >= 2:
            candidate = after[1]

            if len(candidate) < 300:
                affiliation = candidate

    return author, affiliation


# ============================================================
# COUNTRY
# ============================================================

def build_country_names() -> list[tuple[str, str]]:
    values: dict[str, str] = {}

    for alias, canonical in COUNTRY_ALIASES.items():
        values[alias.casefold()] = canonical

    if pycountry is not None:
        for item in pycountry.countries:
            name = clean(getattr(item, "name", ""))

            if name:
                values[name.casefold()] = name

            official = clean(
                getattr(item, "official_name", "")
            )

            if official:
                values[official.casefold()] = name or official

            common = clean(
                getattr(item, "common_name", "")
            )

            if common:
                values[common.casefold()] = name or common

    # longest first prevents "Korea"-like partial matches if aliases overlap
    return sorted(
        values.items(),
        key=lambda item: len(item[0]),
        reverse=True,
    )


COUNTRY_NAMES = build_country_names()


def extract_country(
    affiliation: str,
    pdf_text: str = "",
) -> str:
    # Prefer affiliation, because conference location must not be mistaken for
    # the author's country.
    for source in [affiliation, pdf_text]:
        text = clean(source)

        if not text:
            continue

        lower = text.casefold()

        # First check final comma-separated segment.
        final_segment = clean(
            text.rsplit(",", 1)[-1]
        ).casefold()

        for needle, canonical in COUNTRY_NAMES:
            if final_segment == needle:
                return canonical

        # Then search as a whole-word-ish phrase.
        for needle, canonical in COUNTRY_NAMES:
            if re.search(
                r"(?<![A-Za-z])"
                + re.escape(needle)
                + r"(?![A-Za-z])",
                lower,
                re.I,
            ):
                return canonical

    return ""


# ============================================================
# PROCEEDING SCRAPE
# ============================================================

def scrape_proceeding(
    listing: ListingLink,
    proceeding_url: str,
) -> AuthorRecord:
    try:
        html = get_html(
            proceeding_url,
            referer=listing.url,
        )

        soup = BeautifulSoup(
            html,
            "lxml",
        )

        author, affiliation = (
            extract_author_affiliation_from_html(
                soup
            )
        )

        email = extract_email_from_html(
            soup
        )

        pdf_url = extract_pdf_url(
            soup,
            proceeding_url,
        )

        pdf_text = ""

        if pdf_url:
            try:
                pdf_bytes = download_pdf(
                    pdf_url,
                    proceeding_url,
                )

                pdf_text = extract_pdf_text(
                    pdf_bytes
                )

                if not email:
                    email = extract_email_from_pdf_text(
                        pdf_text
                    )

                author, affiliation = (
                    extract_author_affiliation_from_pdf(
                        pdf_text,
                        author,
                        affiliation,
                    )
                )

            except Exception as pdf_exc:
                log.warning(
                    "PDF failed | %s | %s",
                    pdf_url,
                    pdf_exc,
                )

        year = year_from_text(
            listing.year_hint,
            proceeding_url,
            pdf_text,
            soup.get_text(" ", strip=True),
        )

        country = extract_country(
            affiliation,
            pdf_text,
        )

        return AuthorRecord(
            author=clean(author),
            email=clean(email),
            affiliation=clean(affiliation),
            country=clean(country),
            year=clean(year),
            listing_url=listing.url,
            proceeding_url=proceeding_url,
            pdf_url=pdf_url,
            status="OK",
            error="",
        )

    except Exception as exc:
        return AuthorRecord(
            year=listing.year_hint,
            listing_url=listing.url,
            proceeding_url=proceeding_url,
            status="ERROR",
            error=str(exc)[:2000],
        )


# ============================================================
# CHECKPOINT
# ============================================================

def save_checkpoint(
    records: list[AuthorRecord],
    path: Path,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    pd.DataFrame(
        [asdict(x) for x in records]
    ).to_csv(
        path,
        index=False,
        encoding="utf-8-sig",
    )


def load_checkpoint(
    path: Path,
) -> dict[str, AuthorRecord]:
    if not path.exists():
        return {}

    try:
        df = pd.read_csv(
            path,
            dtype=str,
            keep_default_na=False,
        )

        fields = set(
            AuthorRecord.__dataclass_fields__.keys()
        )

        output = {}

        for row in df.to_dict(orient="records"):
            payload = {
                field: row.get(field, "")
                for field in fields
            }

            record = AuthorRecord(**payload)

            if record.proceeding_url:
                output[record.proceeding_url] = record

        return output

    except Exception as exc:
        log.warning(
            "Checkpoint load failed: %s",
            exc,
        )

        return {}


# ============================================================
# FINAL EXCEL - ONLY 5 COLUMNS
# ============================================================

def export_excel(
    records: list[AuthorRecord],
    output_path: Path,
) -> None:
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    rows = [
        {
            "author": clean(record.author),
            "email": clean(record.email),
            "affiliation": clean(record.affiliation),
            "country": clean(record.country),
            "year": clean(record.year),
        }
        for record in records
        if record.status == "OK"
    ]

    df = pd.DataFrame(
        rows,
        columns=[
            "author",
            "email",
            "affiliation",
            "country",
            "year",
        ],
    )

    if not df.empty:
        # Remove exact duplicate proceeding-derived people.
        df = df.drop_duplicates(
            subset=[
                "author",
                "email",
                "affiliation",
                "year",
            ],
            keep="first",
        )

        df = df.sort_values(
            by=[
                "year",
                "author",
            ],
            ascending=[
                False,
                True,
            ],
            na_position="last",
        ).reset_index(drop=True)

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
    ) as writer:
        df.to_excel(
            writer,
            sheet_name="Authors",
            index=False,
        )

        ws = writer.sheets["Authors"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = {
            "A": 35,
            "B": 40,
            "C": 70,
            "D": 25,
            "E": 12,
        }

        for column, width in widths.items():
            ws.column_dimensions[column].width = width

    log.info(
        "Excel saved: %s",
        output_path.resolve(),
    )


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Scrape Longdom conference proceedings using both HTML and PDF. "
            "Final Excel columns: author, email, affiliation, country, year."
        )
    )

    parser.add_argument(
        "--input",
        default=DEFAULT_INPUT,
    )

    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
    )

    parser.add_argument(
        "--checkpoint",
        default=DEFAULT_CHECKPOINT,
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=MAX_WORKERS,
    )

    parser.add_argument(
        "--limit-listings",
        type=int,
        default=0,
    )

    parser.add_argument(
        "--limit-proceedings",
        type=int,
        default=0,
    )

    parser.add_argument(
        "--fresh",
        action="store_true",
    )

    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)
    checkpoint_path = Path(args.checkpoint)

    listings = read_input_urls(
        input_path
    )

    if args.limit_listings > 0:
        listings = listings[:args.limit_listings]

    jobs: list[tuple[ListingLink, str]] = []
    seen = set()

    for index, listing in enumerate(
        listings,
        start=1,
    ):
        log.info(
            "LISTING [%s/%s] %s",
            index,
            len(listings),
            listing.url,
        )

        try:
            urls = discover_proceedings(
                listing
            )

            for url in urls:
                if url in seen:
                    continue

                seen.add(url)
                jobs.append((listing, url))

        except Exception as exc:
            log.error(
                "Listing failed | %s | %s",
                listing.url,
                exc,
            )

    if args.limit_proceedings > 0:
        jobs = jobs[:args.limit_proceedings]

    log.info(
        "Unique proceedings discovered: %s",
        len(jobs),
    )

    if not jobs:
        raise RuntimeError(
            "No proceeding pages were discovered."
        )

    completed = (
        {}
        if args.fresh
        else load_checkpoint(checkpoint_path)
    )

    pending = [
        job
        for job in jobs
        if (
            job[1] not in completed
            or completed[job[1]].status == "ERROR"
        )
    ]

    log.info(
        "Proceedings remaining: %s",
        len(pending),
    )

    workers = max(
        1,
        min(args.workers, 12),
    )

    if pending:
        with cf.ThreadPoolExecutor(
            max_workers=workers
        ) as executor:
            future_map = {
                executor.submit(
                    scrape_proceeding,
                    listing,
                    url,
                ): (listing, url)
                for listing, url in pending
            }

            processed = 0

            for future in cf.as_completed(future_map):
                listing, url = future_map[future]

                record = future.result()
                completed[url] = record

                processed += 1

                log.info(
                    "[%s/%s] %s | %s | %s | %s",
                    processed,
                    len(pending),
                    record.author or "NO AUTHOR",
                    record.email or "NO EMAIL",
                    record.country or "NO COUNTRY",
                    record.status,
                )

                if processed % 20 == 0:
                    ordered = [
                        completed[url]
                        for _, url in jobs
                        if url in completed
                    ]

                    save_checkpoint(
                        ordered,
                        checkpoint_path,
                    )

    final_records = [
        completed.get(
            url,
            AuthorRecord(
                year=listing.year_hint,
                listing_url=listing.url,
                proceeding_url=url,
                status="ERROR",
                error="Not processed",
            ),
        )
        for listing, url in jobs
    ]

    save_checkpoint(
        final_records,
        checkpoint_path,
    )

    export_excel(
        final_records,
        output_path,
    )

    ok_records = [
        r
        for r in final_records
        if r.status == "OK"
    ]

    with_email = sum(
        bool(r.email)
        for r in ok_records
    )

    print()
    print("=" * 72)
    print("LONGDOM HTML + PDF SCRAPE COMPLETE")
    print("=" * 72)
    print(f"Listings        : {len(listings)}")
    print(f"Proceedings     : {len(final_records)}")
    print(f"Successful      : {len(ok_records)}")
    print(f"With email      : {with_email}")
    print(f"Without email   : {len(ok_records) - with_email}")
    print(f"Output          : {output_path.resolve()}")
    print("=" * 72)


if __name__ == "__main__":
    main()
