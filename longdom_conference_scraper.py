from __future__ import annotations

import argparse
import concurrent.futures as cf
import json
import logging
import re
import threading
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


DEFAULT_INPUT = "longdom_input.xlsx"
DEFAULT_OUTPUT = "output/longdom/longdom_conference_abstracts.xlsx"
DEFAULT_CHECKPOINT = "output/longdom/longdom_checkpoint.csv"

TIMEOUT = 60
MAX_WORKERS = 8
REQUEST_DELAY_SECONDS = 0.10

LONGDOM_HOSTS = {
    "longdom.org",
    "www.longdom.org",
}

GENERIC_EMAILS = {
    "info@longdom.org",
}

EMAIL_RE = re.compile(
    r"[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}",
    re.I,
)

URL_RE = re.compile(
    r'https?://(?:www\.)?longdom\.org/[^\s<>"\')]+',
    re.I,
)

YEAR_RE = re.compile(
    r"\b(20\d{2}|19\d{2})\b"
)

DATE_LOCATION_RE = re.compile(
    r"([A-Za-z]+\s+\d{1,2}(?:-\d{1,2})?,?\s+\d{4})\s*\|\s*(.+)"
)

_thread_local = threading.local()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

log = logging.getLogger("longdom")


# ============================================================
# Models
# ============================================================

@dataclass
class ListingLink:
    url: str
    presentation_type: str = ""
    year_hint: str = ""


@dataclass
class ProceedingRecord:
    presentation_type: str = ""
    year: str = ""

    conference_name: str = ""
    conference_date: str = ""
    conference_location: str = ""

    title: str = ""
    author: str = ""
    affiliation: str = ""
    email: str = ""

    abstract: str = ""
    biography: str = ""

    listing_url: str = ""
    proceeding_url: str = ""

    status: str = "OK"
    error: str = ""


# ============================================================
# Generic helpers
# ============================================================

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


def normalize_url(url: str) -> str:
    return clean(url).rstrip(".,);]>")


def make_session() -> requests.Session:
    session = getattr(
        _thread_local,
        "session",
        None,
    )

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

    session.mount(
        "https://",
        adapter,
    )
    session.mount(
        "http://",
        adapter,
    )

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


def get_html(url: str) -> str:
    session = make_session()

    response = session.get(
        url,
        timeout=TIMEOUT,
    )

    if response.status_code >= 400:
        raise RuntimeError(
            f"HTTP {response.status_code}: {url}"
        )

    if not response.encoding:
        response.encoding = (
            response.apparent_encoding
            or "utf-8"
        )

    time.sleep(
        REQUEST_DELAY_SECONDS
    )

    return response.text


# ============================================================
# Input URL extraction
# ============================================================

def presentation_type_from_url(
    url: str,
) -> str:
    lower = url.lower()

    if "/keynote/" in lower:
        return "Keynote"

    if "/scientific-tracks-abstracts/" in lower:
        return "Scientific Tracks Abstracts"

    if "/posters-accepted-abstracts/" in lower:
        return "Posters & Accepted Abstracts"

    if "/accepted-abstracts/" in lower:
        return "Accepted Abstracts"

    if "/posters/" in lower:
        return "Posters"

    if "/archivejfra/" in lower:
        if "keynote" in lower:
            return "Keynote"

        if "tracks" in lower:
            return "Scientific Tracks Abstracts"

    return "Conference Proceedings"


def year_from_text(
    *values: str,
) -> str:
    for value in values:
        match = YEAR_RE.search(
            clean(value)
        )

        if match:
            return match.group(1)

    return ""


def extract_listing_urls(
    input_path: Path,
) -> list[ListingLink]:
    """
    Read Longdom listing URLs from Excel.

    Supported URL columns:
      url
      URL
      link
      Link
      source_url
      listing_url

    If none of those column names exist, the script scans all columns and
    picks the first column containing Longdom URLs.

    Optional year columns:
      year
      Year
    """
    if not input_path.exists():
        raise FileNotFoundError(
            f"Input Excel not found: {input_path.resolve()}"
        )

    suffix = input_path.suffix.lower()

    if suffix not in {".xlsx", ".xls", ".xlsm"}:
        raise ValueError(
            "Input file must be an Excel file: .xlsx, .xls, or .xlsm"
        )

    df = pd.read_excel(
        input_path,
        dtype=str,
    ).fillna("")

    if df.empty:
        raise RuntimeError(
            "Input Excel is empty."
        )

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

    # If no expected header is present, detect the first column that actually
    # contains Longdom URLs.
    if url_column is None:
        for column in df.columns:
            values = [
                clean(value)
                for value in df[column].tolist()
            ]

            if any(
                value.lower().startswith(("http://", "https://"))
                and "longdom.org" in value.lower()
                for value in values
            ):
                url_column = column
                break

    if url_column is None:
        raise RuntimeError(
            "Could not find a URL column in the Excel file. "
            "Recommended header: url"
        )

    year_column = None

    for candidate in [
        "year",
        "years",
    ]:
        if candidate in normalized_columns:
            year_column = normalized_columns[candidate]
            break

    listings: list[ListingLink] = []
    seen = set()

    for row_no, row in df.iterrows():
        raw_url = clean(
            row.get(
                url_column,
                "",
            )
        )

        if not raw_url:
            continue

        # Excel cells can occasionally contain surrounding text. Extract the
        # first Longdom URL if needed.
        if not raw_url.lower().startswith(("http://", "https://")):
            match = URL_RE.search(
                raw_url
            )

            if match:
                raw_url = match.group(0)

        url = normalize_url(
            raw_url
        )

        try:
            parsed = urlparse(
                url
            )
        except Exception:
            continue

        if parsed.netloc.lower() not in LONGDOM_HOSTS:
            continue

        if "/conference-abstracts/" not in parsed.path.lower():
            continue

        if url in seen:
            continue

        seen.add(
            url
        )

        year_hint = ""

        if year_column is not None:
            year_hint = year_from_text(
                clean(
                    row.get(
                        year_column,
                        "",
                    )
                )
            )

        if not year_hint:
            year_hint = year_from_text(
                url
            )

        listings.append(
            ListingLink(
                url=url,
                presentation_type=presentation_type_from_url(
                    url
                ),
                year_hint=year_hint,
            )
        )

    log.info(
        "Excel rows read: %s | valid unique Longdom listing URLs: %s",
        len(df),
        len(listings),
    )

    if not listings:
        raise RuntimeError(
            f"No valid Longdom conference URLs were found in column "
            f"{url_column!r}."
        )

    return listings


# ============================================================
# Listing page parsing
# ============================================================

def looks_like_proceeding_url(
    href: str,
) -> bool:
    try:
        parsed = urlparse(href)

        return (
            parsed.netloc.lower()
            in LONGDOM_HOSTS
            and "/proceedings/"
            in parsed.path.lower()
            and parsed.path.lower().endswith(
                ".html"
            )
        )
    except Exception:
        return False


def parse_listing_page(
    listing: ListingLink,
    html: str,
) -> list[str]:
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    links = []
    seen = set()

    for anchor in soup.select(
        "a[href]"
    ):
        href = normalize_url(
            urljoin(
                listing.url,
                anchor.get("href", ""),
            )
        )

        if not looks_like_proceeding_url(
            href
        ):
            continue

        if href in seen:
            continue

        seen.add(href)
        links.append(href)

    log.info(
        "LISTING | %s | proceeding links=%s",
        listing.url,
        len(links),
    )

    return links


# ============================================================
# Detail page extraction
# ============================================================

def main_content(
    soup: BeautifulSoup,
) -> Tag:
    for selector in [
        "main",
        "#main-content",
        ".main-content",
        ".content",
        "#content",
        "body",
    ]:
        node = soup.select_one(
            selector
        )

        if node:
            return node

    return soup


def extract_person_email(
    content: Tag,
) -> str:
    candidates = []

    for anchor in content.select(
        'a[href^="mailto:"]'
    ):
        href = clean(
            anchor.get("href")
        )

        match = EMAIL_RE.search(
            href
        )

        if match:
            candidates.append(
                match.group(0).lower()
            )

    for match in EMAIL_RE.findall(
        content.get_text(
            " ",
            strip=True,
        )
    ):
        candidates.append(
            match.lower()
        )

    for email in candidates:
        if email in GENERIC_EMAILS:
            continue

        if email.endswith(
            "@longdom.org"
        ):
            continue

        return email

    return ""


def text_after_heading(
    soup: BeautifulSoup,
    heading_texts: tuple[str, ...],
    *,
    stop_headings: tuple[str, ...] = (),
) -> str:
    """
    Collect text after a heading like 'Abstract :' until the next heading.
    """
    for heading in soup.find_all(
        ["h2", "h3", "h4", "h5", "h6", "strong"]
    ):
        label = clean(
            heading.get_text(
                " ",
                strip=True,
            )
        ).casefold()

        if not any(
            target.casefold() in label
            for target in heading_texts
        ):
            continue

        parts = []

        node = heading.find_next_sibling()

        while node:
            if isinstance(
                node,
                Tag,
            ) and node.name in {
                "h2",
                "h3",
                "h4",
                "h5",
                "h6",
            }:
                next_label = clean(
                    node.get_text(
                        " ",
                        strip=True,
                    )
                ).casefold()

                if (
                    not stop_headings
                    or any(
                        stop.casefold()
                        in next_label
                        for stop
                        in stop_headings
                    )
                    or True
                ):
                    break

            if isinstance(
                node,
                Tag,
            ):
                text = clean(
                    node.get_text(
                        " ",
                        strip=True,
                    )
                )

                if text:
                    parts.append(
                        text
                    )

            node = node.find_next_sibling()

        return clean(
            " ".join(
                parts
            )
        )

    return ""


def heading_candidates(
    soup: BeautifulSoup,
) -> list[str]:
    output = []

    for node in soup.find_all(
        ["h1", "h2", "h3", "h4"]
    ):
        text = clean(
            node.get_text(
                " ",
                strip=True,
            )
        )

        if text and text not in output:
            output.append(text)

    return output


def parse_date_location(
    text: str,
) -> tuple[str, str]:
    text = clean(text)

    match = DATE_LOCATION_RE.search(
        text
    )

    if not match:
        return "", ""

    return (
        clean(
            match.group(1)
        ),
        clean(
            match.group(2)
        ),
    )


def extract_conference_metadata(
    soup: BeautifulSoup,
    listing: ListingLink,
) -> tuple[
    str,
    str,
    str,
    str,
]:
    """
    Returns:
      conference_name, date, location, year
    """
    headings = heading_candidates(
        soup
    )

    conference_name = ""
    conference_date = ""
    conference_location = ""

    # On Longdom proceedings pages, conference name is commonly directly
    # under the abstract title as h4/h3 text and date|location follows it.
    for node in soup.find_all(
        ["h3", "h4"]
    ):
        text = clean(
            node.get_text(
                " ",
                strip=True,
            )
        )

        lower = text.casefold()

        if (
            "conference"
            in lower
            or "congress"
            in lower
            or "summit"
            in lower
        ):
            conference_name = text

            sibling = node.find_next_sibling()

            for _ in range(5):
                if sibling is None:
                    break

                if isinstance(
                    sibling,
                    Tag,
                ):
                    sibling_text = clean(
                        sibling.get_text(
                            " ",
                            strip=True,
                        )
                    )

                    date, location = parse_date_location(
                        sibling_text
                    )

                    if date or location:
                        conference_date = date
                        conference_location = location
                        break

                sibling = sibling.find_next_sibling()

            if conference_name:
                break

    # Fallback: scan visible text for a Date | Location line.
    if not conference_date:
        body_text = soup.get_text(
            "\n",
            strip=True,
        )

        for line in body_text.splitlines():
            date, location = parse_date_location(
                line
            )

            if date:
                conference_date = date
                conference_location = location
                break

    year = year_from_text(
        conference_date,
        listing.year_hint,
        listing.url,
    )

    return (
        conference_name,
        conference_date,
        conference_location,
        year,
    )


def extract_title(
    soup: BeautifulSoup,
) -> str:
    # Longdom proceeding title is typically the first main h2.
    candidates = []

    for selector in [
        "h2",
        "h1",
    ]:
        for node in soup.select(
            selector
        ):
            text = clean(
                node.get_text(
                    " ",
                    strip=True,
                )
            )

            lower = text.casefold()

            if not text:
                continue

            if "international journal" in lower:
                continue

            if (
                "conference"
                in lower
                and len(text) < 180
            ):
                continue

            candidates.append(
                text
            )

    return (
        candidates[0]
        if candidates
        else ""
    )


def is_noise_line(
    text: str,
) -> bool:
    lower = clean(text).casefold()

    noise = [
        "open access",
        "abstract",
        "biography",
        "content links",
        "contact longdom",
        "international journal",
        "instructions for authors",
        "journal home",
        "submit manuscript",
    ]

    return any(
        token in lower
        for token in noise
    )


def extract_author_affiliation(
    soup: BeautifulSoup,
    title: str,
    conference_name: str,
) -> tuple[str, str]:
    """
    Longdom structure observed:
        <h2>Abstract title</h2>
        <h4>Conference name</h4>
        date/location
        Author
        Affiliation
        ...
        Abstract :
    We collect text blocks between the date/location and Abstract heading.
    """
    abstract_heading = None

    for node in soup.find_all(
        ["h4", "h5", "h6", "strong"]
    ):
        text = clean(
            node.get_text(
                " ",
                strip=True,
            )
        ).casefold()

        if text.startswith(
            "abstract"
        ):
            abstract_heading = node
            break

    if abstract_heading:
        previous = []

        node = abstract_heading.previous_sibling

        while node and len(
            previous
        ) < 12:
            if isinstance(
                node,
                Tag,
            ):
                text = clean(
                    node.get_text(
                        " ",
                        strip=True,
                    )
                )

                if (
                    text
                    and text != title
                    and text != conference_name
                    and not is_noise_line(
                        text
                    )
                ):
                    previous.append(
                        text
                    )

            node = node.previous_sibling

        previous.reverse()

        # Remove date/location, journal/category text, HTML/PDF labels.
        candidates = []

        for text in previous:
            lower = text.casefold()

            if parse_date_location(
                text
            )[0]:
                continue

            if lower in {
                "html",
                "pdf",
            }:
                continue

            if (
                "scientific tracks abstracts:"
                in lower
                or "keynote:"
                in lower
                or "posters"
                in lower
            ):
                continue

            candidates.append(
                text
            )

        if candidates:
            author = candidates[-2] if len(
                candidates
            ) >= 2 else candidates[0]

            affiliation = candidates[-1] if len(
                candidates
            ) >= 2 else ""

            # If last value is short person-name-like text, switch.
            if (
                affiliation
                and len(
                    affiliation.split()
                )
                <= 7
                and "," not in affiliation
                and len(
                    author.split()
                )
                > 7
            ):
                author, affiliation = (
                    affiliation,
                    author,
                )

            return (
                clean(author),
                clean(affiliation),
            )

    return "", ""


def parse_proceeding(
    listing: ListingLink,
    proceeding_url: str,
    html: str,
) -> ProceedingRecord:
    soup = BeautifulSoup(
        html,
        "lxml",
    )

    content = main_content(
        soup
    )

    title = extract_title(
        soup
    )

    (
        conference_name,
        conference_date,
        conference_location,
        year,
    ) = extract_conference_metadata(
        soup,
        listing,
    )

    author, affiliation = (
        extract_author_affiliation(
            soup,
            title,
            conference_name,
        )
    )

    abstract = text_after_heading(
        soup,
        (
            "Abstract",
        ),
    )

    biography = text_after_heading(
        soup,
        (
            "Biography",
            "Bio",
        ),
    )

    # If generic heading collector did not work because paragraphs are nested,
    # use text boundaries as a fallback.
    if not abstract:
        full_text = content.get_text(
            "\n",
            strip=True,
        )

        match = re.search(
            r"\bAbstract\s*:?\s*(.+?)(?=\bBiography\s*:|\bContent Links\b|\Z)",
            full_text,
            re.I | re.S,
        )

        if match:
            abstract = clean(
                match.group(1)
            )

    if not biography:
        full_text = content.get_text(
            "\n",
            strip=True,
        )

        match = re.search(
            r"\bBiography\s*:?\s*(.+?)(?=\bContent Links\b|\bContact Longdom\b|\Z)",
            full_text,
            re.I | re.S,
        )

        if match:
            biography = clean(
                match.group(1)
            )

    return ProceedingRecord(
        presentation_type=listing.presentation_type,
        year=year,
        conference_name=conference_name,
        conference_date=conference_date,
        conference_location=conference_location,
        title=title,
        author=author,
        affiliation=affiliation,
        email=extract_person_email(
            content
        ),
        abstract=abstract,
        biography=biography,
        listing_url=listing.url,
        proceeding_url=proceeding_url,
        status="OK",
        error="",
    )


# ============================================================
# Worker / checkpoint
# ============================================================

def scrape_one(
    listing: ListingLink,
    proceeding_url: str,
) -> ProceedingRecord:
    try:
        html = get_html(
            proceeding_url
        )

        return parse_proceeding(
            listing,
            proceeding_url,
            html,
        )

    except Exception as exc:
        return ProceedingRecord(
            presentation_type=listing.presentation_type,
            year=listing.year_hint,
            listing_url=listing.url,
            proceeding_url=proceeding_url,
            status="ERROR",
            error=str(exc)[:2000],
        )


def checkpoint_columns() -> list[str]:
    return list(
        ProceedingRecord.__dataclass_fields__.keys()
    )


def save_checkpoint(
    records: list[ProceedingRecord],
    path: Path,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    pd.DataFrame(
        [
            asdict(record)
            for record in records
        ],
        columns=checkpoint_columns(),
    ).to_csv(
        path,
        index=False,
        encoding="utf-8-sig",
    )


def load_checkpoint(
    path: Path,
) -> dict[str, ProceedingRecord]:
    if not path.exists():
        return {}

    try:
        df = pd.read_csv(
            path,
            dtype=str,
            keep_default_na=False,
        )

        output = {}

        fields = set(
            checkpoint_columns()
        )

        for row in df.to_dict(
            orient="records"
        ):
            payload = {
                key: row.get(
                    key,
                    "",
                )
                for key in fields
            }

            record = ProceedingRecord(
                **payload
            )

            if record.proceeding_url:
                output[
                    record.proceeding_url
                ] = record

        return output

    except Exception as exc:
        log.warning(
            "Could not load checkpoint: %s",
            exc,
        )

        return {}


# ============================================================
# Excel export
# ============================================================

def write_sheet(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    name: str,
) -> None:
    df.to_excel(
        writer,
        sheet_name=name,
        index=False,
    )

    ws = writer.sheets[
        name
    ]

    ws.freeze_panes = "A2"

    if ws.max_row >= 1:
        ws.auto_filter.ref = (
            ws.dimensions
        )

    for column_cells in ws.columns:
        letter = column_cells[
            0
        ].column_letter

        max_length = 0

        for cell in column_cells[
            :500
        ]:
            max_length = max(
                max_length,
                len(
                    clean(
                        cell.value
                    )
                ),
            )

        ws.column_dimensions[
            letter
        ].width = min(
            max(
                max_length + 2,
                12,
            ),
            70,
        )


def export_excel(
    records: list[ProceedingRecord],
    listings: list[ListingLink],
    output: Path,
) -> None:
    output.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    df = pd.DataFrame(
        [
            asdict(record)
            for record in records
        ],
        columns=checkpoint_columns(),
    )

    if not df.empty:
        df = df.sort_values(
            by=[
                "year",
                "presentation_type",
                "author",
                "title",
            ],
            ascending=[
                False,
                True,
                True,
                True,
            ],
            na_position="last",
        ).reset_index(
            drop=True
        )

    errors_df = (
        df[
            df["status"] == "ERROR"
        ].copy()
        if not df.empty
        else pd.DataFrame(
            columns=checkpoint_columns()
        )
    )

    email_df = (
        df[
            df["email"]
            .fillna("")
            .astype(str)
            .str.strip()
            != ""
        ].copy()
        if not df.empty
        else pd.DataFrame(
            columns=checkpoint_columns()
        )
    )

    listings_df = pd.DataFrame(
        [
            {
                "listing_url": item.url,
                "presentation_type": item.presentation_type,
                "year_hint": item.year_hint,
            }
            for item in listings
        ]
    )

    summary_df = pd.DataFrame(
        [
            {
                "metric": "Input listing URLs",
                "value": len(
                    listings
                ),
            },
            {
                "metric": "Proceedings records",
                "value": len(
                    records
                ),
            },
            {
                "metric": "Records with published person email",
                "value": len(
                    email_df
                ),
            },
            {
                "metric": "Errors",
                "value": len(
                    errors_df
                ),
            },
        ]
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        write_sheet(
            writer,
            df,
            "Proceedings",
        )

        write_sheet(
            writer,
            email_df,
            "With Emails",
        )

        write_sheet(
            writer,
            listings_df,
            "Input Listings",
        )

        write_sheet(
            writer,
            errors_df,
            "Errors",
        )

        write_sheet(
            writer,
            summary_df,
            "Summary",
        )

    log.info(
        "Excel saved: %s",
        output.resolve(),
    )


# ============================================================
# Main
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Scrape Longdom conference listing URLs from an Excel file, "
            "follow each /proceedings/ abstract page, and export structured data to Excel."
        )
    )

    parser.add_argument(
        "--input",
        default=DEFAULT_INPUT,
        help="Excel file containing Longdom conference listing URLs. Recommended column: url",
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
        help="Process only first N listing URLs for testing; 0 = all.",
    )

    parser.add_argument(
        "--limit-proceedings",
        type=int,
        default=0,
        help="Process only first N unique proceedings for testing; 0 = all.",
    )

    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Ignore checkpoint and re-scrape proceedings.",
    )

    args = parser.parse_args()

    input_path = Path(
        args.input
    )

    output_path = Path(
        args.output
    )

    checkpoint_path = Path(
        args.checkpoint
    )

    if not input_path.exists():
        raise FileNotFoundError(
            f"Input file not found: {input_path.resolve()}"
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    checkpoint_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    listings = extract_listing_urls(
        input_path
    )

    if args.limit_listings > 0:
        listings = listings[
            : args.limit_listings
        ]

    if not listings:
        raise RuntimeError(
            "No Longdom /conference-abstracts/ URLs were found in the input Excel file."
        )

    log.info(
        "Unique conference listing URLs: %s",
        len(listings),
    )

    # Discover all proceeding URLs.
    proceeding_jobs: list[
        tuple[
            ListingLink,
            str,
        ]
    ] = []

    seen_proceedings = set()

    for index, listing in enumerate(
        listings,
        start=1,
    ):
        log.info(
            "DISCOVER [%s/%s] %s",
            index,
            len(
                listings
            ),
            listing.url,
        )

        try:
            html = get_html(
                listing.url
            )

            urls = parse_listing_page(
                listing,
                html,
            )

            for url in urls:
                if url in seen_proceedings:
                    continue

                seen_proceedings.add(
                    url
                )

                proceeding_jobs.append(
                    (
                        listing,
                        url,
                    )
                )

        except Exception as exc:
            log.error(
                "Listing failed | %s | %s",
                listing.url,
                exc,
            )

    if args.limit_proceedings > 0:
        proceeding_jobs = proceeding_jobs[
            : args.limit_proceedings
        ]

    log.info(
        "Unique proceedings discovered: %s",
        len(proceeding_jobs),
    )

    if not proceeding_jobs:
        raise RuntimeError(
            "No /proceedings/ abstract links were discovered from the supplied listing URLs."
        )

    completed = (
        {}
        if args.fresh
        else load_checkpoint(
            checkpoint_path
        )
    )

    pending = [
        job
        for job in proceeding_jobs
        if (
            job[1] not in completed
            or completed[
                job[1]
            ].status
            == "ERROR"
        )
    ]

    log.info(
        "Proceedings remaining: %s",
        len(pending),
    )

    workers = max(
        1,
        min(
            int(
                args.workers
            ),
            20,
        ),
    )

    if pending:
        with cf.ThreadPoolExecutor(
            max_workers=workers
        ) as executor:
            futures = {
                executor.submit(
                    scrape_one,
                    listing,
                    proceeding_url,
                ): (
                    listing,
                    proceeding_url,
                )
                for (
                    listing,
                    proceeding_url,
                ) in pending
            }

            processed = 0

            for future in cf.as_completed(
                futures
            ):
                (
                    listing,
                    proceeding_url,
                ) = futures[
                    future
                ]

                record = future.result()

                completed[
                    proceeding_url
                ] = record

                processed += 1

                log.info(
                    "[%s/%s] %s | %s | %s",
                    processed,
                    len(
                        pending
                    ),
                    record.author
                    or "NO AUTHOR",
                    record.title[
                        :90
                    ]
                    or proceeding_url,
                    record.status,
                )

                if (
                    processed % 25
                    == 0
                ):
                    ordered = [
                        completed[
                            url
                        ]
                        for (
                            _listing,
                            url,
                        ) in proceeding_jobs
                        if url in completed
                    ]

                    save_checkpoint(
                        ordered,
                        checkpoint_path,
                    )

    final_records = [
        completed.get(
            proceeding_url,
            ProceedingRecord(
                presentation_type=listing.presentation_type,
                year=listing.year_hint,
                listing_url=listing.url,
                proceeding_url=proceeding_url,
                status="ERROR",
                error="Proceeding was not processed.",
            ),
        )
        for (
            listing,
            proceeding_url,
        ) in proceeding_jobs
    ]

    save_checkpoint(
        final_records,
        checkpoint_path,
    )

    export_excel(
        final_records,
        listings,
        output_path,
    )

    with_email = sum(
        bool(
            record.email
        )
        for record in final_records
    )

    errors = sum(
        record.status
        == "ERROR"
        for record in final_records
    )

    print()
    print("=" * 78)
    print(
        "LONGDOM CONFERENCE SCRAPE COMPLETE"
    )
    print("=" * 78)
    print(
        f"Input listings       : {len(listings)}"
    )
    print(
        f"Proceedings          : {len(final_records)}"
    )
    print(
        f"With person email    : {with_email}"
    )
    print(
        f"Errors               : {errors}"
    )
    print(
        f"Checkpoint           : {checkpoint_path.resolve()}"
    )
    print(
        f"Excel                : {output_path.resolve()}"
    )
    print("=" * 78)


if __name__ == "__main__":
    main()
