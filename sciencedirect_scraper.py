import os
import re
import json
import time
import random
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd
from bs4 import BeautifulSoup

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ============================================================
# CONFIGURATION
# ============================================================

OUTPUT_FOLDER = "output_sciencedirect"

OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "sciencedirect_authors.xlsx"
)

CHECKPOINT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "checkpoint.json"
)

ERROR_FILE = os.path.join(
    OUTPUT_FOLDER,
    "errors.json"
)


HEADLESS = False

PAGE_TIMEOUT = 60000

ARTICLE_DELAY_MIN = 1.5
ARTICLE_DELAY_MAX = 3.0

ISSUE_DELAY_MIN = 2.0
ISSUE_DELAY_MAX = 4.0


# ============================================================
# JOURNALS + YEAR RANGES
# ============================================================

JOURNALS = [
    {
        "name": "Journal of PeriAnesthesia Nursing",
        "url": "https://www.sciencedirect.com/journal/journal-of-perianesthesia-nursing/issues",
        "start_year": 2025,
        "end_year": 2026,
    },
    {
        "name": "Journal of Clinical Anesthesia",
        "url": "https://www.sciencedirect.com/journal/journal-of-clinical-anesthesia/issues",
        "start_year": 2025,
        "end_year": 2027,
    },
    {
        "name": "Journal of Cardiothoracic and Vascular Anesthesia",
        "url": "https://www.sciencedirect.com/journal/journal-of-cardiothoracic-and-vascular-anesthesia/issues",
        "start_year": 2025,
        "end_year": 2026,
    },
    {
        "name": "Journal of Anesthesia and Translational Medicine",
        "url": "https://www.sciencedirect.com/journal/journal-of-anesthesia-and-translational-medicine/issues",
        "start_year": 2025,
        "end_year": 2026,
    },
    {
        "name": "International Journal of Obstetric Anesthesia",
        "url": "https://www.sciencedirect.com/journal/international-journal-of-obstetric-anesthesia/issues",
        "start_year": 2025,
        "end_year": 2026,
    },
]


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""

    value = str(value)

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


def sleep_random(minimum, maximum):
    time.sleep(
        random.uniform(
            minimum,
            maximum
        )
    )


def safe_year(text):
    if not text:
        return None

    match = re.search(
        r"\b(20\d{2})\b",
        text
    )

    if match:
        return int(
            match.group(1)
        )

    return None


def save_json(path, data):
    with open(
        path,
        "w",
        encoding="utf-8"
    ) as f:
        json.dump(
            data,
            f,
            indent=2,
            ensure_ascii=False
        )


def load_json(path):
    if not os.path.exists(path):
        return {}

    try:
        with open(
            path,
            "r",
            encoding="utf-8"
        ) as f:
            return json.load(f)

    except Exception:
        return {}


# ============================================================
# COOKIE / CONSENT
# ============================================================

def accept_cookies(page):
    selectors = [
        "button:has-text('Accept all cookies')",
        "button:has-text('Accept All Cookies')",
        "button:has-text('Accept cookies')",
        "button[id*='accept']",
        "#onetrust-accept-btn-handler",
    ]

    for selector in selectors:

        try:
            locator = page.locator(
                selector
            )

            if locator.count() > 0:
                locator.first.click(
                    timeout=3000
                )

                time.sleep(1)
                return

        except Exception:
            pass


# ============================================================
# SCROLL PAGE
# ============================================================

def auto_scroll(page):

    try:
        previous_height = 0

        for _ in range(20):

            current_height = page.evaluate(
                "document.body.scrollHeight"
            )

            if current_height == previous_height:
                break

            page.evaluate(
                "window.scrollTo(0, document.body.scrollHeight)"
            )

            time.sleep(1)

            previous_height = current_height

    except Exception:
        pass


# ============================================================
# LOAD PAGE
# ============================================================

def load_page(page, url):

    print(
        f"Opening: {url}"
    )

    for attempt in range(1, 4):

        try:

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT
            )

            accept_cookies(page)

            time.sleep(2)

            return True

        except PlaywrightTimeoutError:

            print(
                f"Timeout attempt {attempt}: {url}"
            )

        except Exception as exc:

            print(
                f"Load error attempt {attempt}: "
                f"{exc}"
            )

        time.sleep(
            attempt * 3
        )

    return False


# ============================================================
# EXTRACT ISSUE LINKS
# ============================================================

def extract_issue_links(
    page,
    journal
):

    url = journal["url"]

    start_year = journal[
        "start_year"
    ]

    end_year = journal[
        "end_year"
    ]

    if not load_page(
        page,
        url
    ):
        return []

    auto_scroll(page)

    html = page.content()

    soup = BeautifulSoup(
        html,
        "lxml"
    )

    issue_links = {}

    for anchor in soup.find_all(
        "a",
        href=True
    ):

        href = anchor.get(
            "href",
            ""
        )

        text = clean_text(
            anchor.get_text(
                " ",
                strip=True
            )
        )

        # Typical ScienceDirect issue links contain:
        # /vol/...
        # /issue/...
        # ?pii=...
        if (
            "/journal/" not in href
            and "/issue/" not in href
            and "/vol/" not in href
        ):
            continue

        full_url = urljoin(
            "https://www.sciencedirect.com",
            href
        )

        combined_text = (
            text + " " + href
        )

        year = safe_year(
            combined_text
        )

        # If we can detect year, apply year filter immediately.
        if year:

            if not (
                start_year
                <= year
                <= end_year
            ):
                continue

        if (
            "issues" in full_url
            and full_url.rstrip("/")
            == url.rstrip("/")
        ):
            continue

        issue_links[
            full_url
        ] = {
            "url": full_url,
            "text": text,
            "possible_year": year,
        }

    print(
        f"Issue candidates found: "
        f"{len(issue_links)}"
    )

    return list(
        issue_links.values()
    )


# ============================================================
# ISSUE METADATA
# ============================================================

def extract_issue_metadata(
    soup,
    issue_url
):

    page_text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    )

    year = safe_year(
        page_text
    )

    volume = ""
    issue = ""
    issue_date = ""

    # --------------------------------------------------------
    # VOLUME
    # --------------------------------------------------------

    patterns = [
        r"\bVolume\s+([0-9A-Za-z\-]+)",
        r"\bVol\.?\s+([0-9A-Za-z\-]+)",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            page_text,
            re.I
        )

        if match:
            volume = clean_text(
                match.group(1)
            )
            break

    # --------------------------------------------------------
    # ISSUE
    # --------------------------------------------------------

    patterns = [
        r"\bIssue\s+([0-9A-Za-z\-]+)",
        r"\bNumber\s+([0-9A-Za-z\-]+)",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            page_text,
            re.I
        )

        if match:
            issue = clean_text(
                match.group(1)
            )
            break

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    month_pattern = (
        r"\b("
        r"January|February|March|April|May|June|"
        r"July|August|September|October|November|December"
        r")\s+(20\d{2})\b"
    )

    match = re.search(
        month_pattern,
        page_text,
        re.I
    )

    if match:

        issue_date = (
            f"{match.group(1)} "
            f"{match.group(2)}"
        )

        year = int(
            match.group(2)
        )

    return {
        "year": year,
        "volume": volume,
        "issue": issue,
        "issue_date": issue_date,
        "issue_url": issue_url,
    }


# ============================================================
# EXTRACT ARTICLE LINKS FROM ISSUE
# ============================================================

def extract_article_links(
    page,
    issue_url
):

    if not load_page(
        page,
        issue_url
    ):
        return [], {}

    auto_scroll(page)

    soup = BeautifulSoup(
        page.content(),
        "lxml"
    )

    issue_metadata = (
        extract_issue_metadata(
            soup,
            issue_url
        )
    )

    article_links = {}

    for anchor in soup.find_all(
        "a",
        href=True
    ):

        href = anchor[
            "href"
        ]

        if (
            "/science/article/" not in href
        ):
            continue

        full_url = urljoin(
            "https://www.sciencedirect.com",
            href
        )

        # Remove URL fragments
        full_url = full_url.split(
            "#"
        )[0]

        title = clean_text(
            anchor.get_text(
                " ",
                strip=True
            )
        )

        if full_url not in article_links:

            article_links[
                full_url
            ] = {
                "url": full_url,
                "title": title,
            }

    return (
        list(
            article_links.values()
        ),
        issue_metadata
    )


# ============================================================
# META TAG HELPER
# ============================================================

def get_meta(
    soup,
    name=None,
    property_name=None
):

    tag = None

    if name:

        tag = soup.find(
            "meta",
            attrs={
                "name": name
            }
        )

    if (
        not tag
        and property_name
    ):

        tag = soup.find(
            "meta",
            attrs={
                "property":
                    property_name
            }
        )

    if tag:

        return clean_text(
            tag.get(
                "content",
                ""
            )
        )

    return ""


# ============================================================
# AUTHOR EXTRACTION
# ============================================================

def extract_authors(soup):

    authors = []

    # --------------------------------------------------------
    # citation_author meta tags
    # --------------------------------------------------------

    meta_authors = soup.find_all(
        "meta",
        attrs={
            "name":
                "citation_author"
        }
    )

    for meta in meta_authors:

        name = clean_text(
            meta.get(
                "content",
                ""
            )
        )

        if (
            name
            and name not in authors
        ):
            authors.append(
                name
            )

    # --------------------------------------------------------
    # DC.creator fallback
    # --------------------------------------------------------

    if not authors:

        creator_tags = soup.find_all(
            "meta",
            attrs={
                "name":
                    re.compile(
                        r"DC\.creator",
                        re.I
                    )
            }
        )

        for tag in creator_tags:

            name = clean_text(
                tag.get(
                    "content",
                    ""
                )
            )

            if (
                name
                and name not in authors
            ):
                authors.append(
                    name
                )

    return authors


# ============================================================
# EMAIL EXTRACTION
# ============================================================

def extract_emails(soup):

    emails = []

    # --------------------------------------------------------
    # mailto links
    # --------------------------------------------------------

    for anchor in soup.find_all(
        "a",
        href=True
    ):

        href = anchor.get(
            "href",
            ""
        )

        if href.lower().startswith(
            "mailto:"
        ):

            email = href.split(
                ":",
                1
            )[1]

            email = email.split(
                "?"
            )[0]

            email = clean_text(
                email
            ).lower()

            if (
                email
                and email not in emails
            ):
                emails.append(
                    email
                )

    # --------------------------------------------------------
    # Regex fallback
    # --------------------------------------------------------

    page_text = soup.get_text(
        " ",
        strip=True
    )

    matches = re.findall(
        r"\b[A-Za-z0-9._%+\-]+"
        r"@[A-Za-z0-9.\-]+"
        r"\.[A-Za-z]{2,}\b",
        page_text
    )

    for email in matches:

        email = email.lower()

        if email not in emails:
            emails.append(
                email
            )

    return emails


# ============================================================
# DOI
# ============================================================

def extract_doi(soup):

    doi = get_meta(
        soup,
        name="citation_doi"
    )

    if doi:
        return doi

    text = clean_text(
        soup.get_text(
            " ",
            strip=True
        )
    )

    match = re.search(
        r"\b10\.\d{4,9}/"
        r"[-._;()/:A-Z0-9]+",
        text,
        re.I
    )

    if match:
        return match.group(
            0
        ).rstrip(
            ".,;)"
        )

    return ""


# ============================================================
# EXTRACT ARTICLE
# ============================================================

def extract_article(
    page,
    journal_name,
    article_url,
    issue_metadata
):

    if not load_page(
        page,
        article_url
    ):
        return []

    soup = BeautifulSoup(
        page.content(),
        "lxml"
    )

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    title = get_meta(
        soup,
        name="citation_title"
    )

    if not title:

        h1 = soup.find(
            "h1"
        )

        if h1:
            title = clean_text(
                h1.get_text(
                    " ",
                    strip=True
                )
            )

    # --------------------------------------------------------
    # JOURNAL
    # --------------------------------------------------------

    journal = get_meta(
        soup,
        name="citation_journal_title"
    )

    if not journal:
        journal = journal_name

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    publication_date = (
        get_meta(
            soup,
            name="citation_publication_date"
        )
        or get_meta(
            soup,
            name="citation_date"
        )
    )

    article_year = safe_year(
        publication_date
    )

    # Prefer issue year where possible.
    year = (
        issue_metadata.get(
            "year"
        )
        or article_year
    )

    # --------------------------------------------------------
    # VOLUME / ISSUE
    # --------------------------------------------------------

    volume = (
        get_meta(
            soup,
            name="citation_volume"
        )
        or issue_metadata.get(
            "volume",
            ""
        )
    )

    issue = (
        get_meta(
            soup,
            name="citation_issue"
        )
        or issue_metadata.get(
            "issue",
            ""
        )
    )

    # --------------------------------------------------------
    # DOI
    # --------------------------------------------------------

    doi = extract_doi(
        soup
    )

    # --------------------------------------------------------
    # AUTHORS
    # --------------------------------------------------------

    authors = extract_authors(
        soup
    )

    emails = extract_emails(
        soup
    )

    authors_joined = "; ".join(
        authors
    )

    # --------------------------------------------------------
    # ONE ROW PER AUTHOR
    # --------------------------------------------------------

    records = []

    if authors:

        for index, author in enumerate(
            authors
        ):

            # Email-to-author mapping is not always explicit.
            # Only assign by same index when counts match.
            email = ""

            if (
                len(authors)
                == len(emails)
                and index < len(emails)
            ):
                email = emails[index]

            elif (
                len(authors) == 1
                and len(emails) == 1
            ):
                email = emails[0]

            records.append(
                {
                    "journal": journal,
                    "year": year or "",
                    "volume": volume,
                    "issue": issue,
                    "issue_date":
                        issue_metadata.get(
                            "issue_date",
                            ""
                        ),
                    "article_title": title,
                    "author": author,
                    "email": email,
                    "all_authors":
                        authors_joined,
                    "doi": doi,
                    "article_url":
                        article_url,
                    "issue_url":
                        issue_metadata.get(
                            "issue_url",
                            ""
                        ),
                }
            )

    else:

        records.append(
            {
                "journal": journal,
                "year": year or "",
                "volume": volume,
                "issue": issue,
                "issue_date":
                    issue_metadata.get(
                        "issue_date",
                        ""
                    ),
                "article_title": title,
                "author": "",
                "email":
                    emails[0]
                    if len(emails) == 1
                    else "",
                "all_authors": "",
                "doi": doi,
                "article_url":
                    article_url,
                "issue_url":
                    issue_metadata.get(
                        "issue_url",
                        ""
                    ),
            }
        )

    return records


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(records):

    if not records:
        return

    df = pd.DataFrame(
        records
    )

    # --------------------------------------------------------
    # REMOVE DUPLICATES
    # --------------------------------------------------------

    df = df.drop_duplicates(
        subset=[
            "article_url",
            "author"
        ],
        keep="first"
    )

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    if "year" in df.columns:

        df["year_sort"] = pd.to_numeric(
            df["year"],
            errors="coerce"
        )

        df = df.sort_values(
            by=[
                "journal",
                "year_sort",
                "volume",
                "issue",
                "article_title",
                "author",
            ],
            ascending=[
                True,
                True,
                True,
                True,
                True,
                True,
            ]
        )

        df.drop(
            columns=[
                "year_sort"
            ],
            inplace=True
        )

    # --------------------------------------------------------
    # EXPORT
    # --------------------------------------------------------

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Authors",
            index=False
        )

        worksheet = writer.book[
            "Authors"
        ]

        worksheet.freeze_panes = "A2"

        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        widths = {
            "A": 45,
            "B": 12,
            "C": 12,
            "D": 12,
            "E": 22,
            "F": 70,
            "G": 35,
            "H": 40,
            "I": 80,
            "J": 35,
            "K": 75,
            "L": 75,
        }

        for column, width in widths.items():

            worksheet.column_dimensions[
                column
            ].width = width


# ============================================================
# MAIN
# ============================================================

def main():

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    checkpoint = load_json(
        CHECKPOINT_FILE
    )

    completed_articles = set(
        checkpoint.get(
            "completed_articles",
            []
        )
    )

    records = checkpoint.get(
        "records",
        []
    )

    errors = []

    print("=" * 110)
    print("SCIENCEDIRECT JOURNAL SCRAPER")
    print("=" * 110)

    print(
        f"Existing records       : "
        f"{len(records):,}"
    )

    print(
        f"Completed article URLs : "
        f"{len(completed_articles):,}"
    )

    print()

    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=HEADLESS,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--start-maximized",
            ]
        )

        context = browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000
            },
            user_agent=(
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/140.0.0.0 Safari/537.36"
            ),
            locale="en-US",
        )

        page = context.new_page()

        # ====================================================
        # JOURNALS
        # ====================================================

        for journal_index, journal in enumerate(
            JOURNALS,
            start=1
        ):

            print()
            print("=" * 110)

            print(
                f"[{journal_index}/{len(JOURNALS)}] "
                f"{journal['name']}"
            )

            print(
                f"Years: "
                f"{journal['start_year']} - "
                f"{journal['end_year']}"
            )

            print("=" * 110)

            issue_candidates = (
                extract_issue_links(
                    page,
                    journal
                )
            )

            print(
                f"Issue candidates: "
                f"{len(issue_candidates)}"
            )

            # ================================================
            # ISSUES
            # ================================================

            for issue_index, issue_item in enumerate(
                issue_candidates,
                start=1
            ):

                issue_url = issue_item[
                    "url"
                ]

                print()
                print(
                    f"Issue "
                    f"{issue_index}/"
                    f"{len(issue_candidates)}"
                )

                try:

                    article_links, issue_metadata = (
                        extract_article_links(
                            page,
                            issue_url
                        )
                    )

                    issue_year = (
                        issue_metadata.get(
                            "year"
                        )
                    )

                    # ----------------------------------------
                    # YEAR FILTER
                    # ----------------------------------------

                    if issue_year:

                        if not (
                            journal[
                                "start_year"
                            ]
                            <= issue_year
                            <= journal[
                                "end_year"
                            ]
                        ):

                            print(
                                f"Skipping year "
                                f"{issue_year}"
                            )

                            continue

                    print(
                        f"Year: "
                        f"{issue_year}"
                    )

                    print(
                        f"Articles found: "
                        f"{len(article_links)}"
                    )

                    # ========================================
                    # ARTICLES
                    # ========================================

                    for article_index, article_item in enumerate(
                        article_links,
                        start=1
                    ):

                        article_url = article_item[
                            "url"
                        ]

                        if (
                            article_url
                            in completed_articles
                        ):

                            print(
                                f"  [{article_index}/"
                                f"{len(article_links)}] "
                                f"Already done"
                            )

                            continue

                        print(
                            f"  [{article_index}/"
                            f"{len(article_links)}] "
                            f"{article_url}"
                        )

                        try:

                            article_records = (
                                extract_article(
                                    page,
                                    journal[
                                        "name"
                                    ],
                                    article_url,
                                    issue_metadata
                                )
                            )

                            # Final article year validation
                            filtered = []

                            for record in article_records:

                                record_year = (
                                    record.get(
                                        "year"
                                    )
                                )

                                try:
                                    record_year_int = int(
                                        record_year
                                    )
                                except Exception:
                                    record_year_int = None

                                if (
                                    record_year_int
                                    and journal[
                                        "start_year"
                                    ]
                                    <= record_year_int
                                    <= journal[
                                        "end_year"
                                    ]
                                ):
                                    filtered.append(
                                        record
                                    )

                                elif not record_year_int:
                                    filtered.append(
                                        record
                                    )

                            records.extend(
                                filtered
                            )

                            completed_articles.add(
                                article_url
                            )

                            # --------------------------------
                            # CHECKPOINT
                            # --------------------------------

                            save_json(
                                CHECKPOINT_FILE,
                                {
                                    "completed_articles":
                                        list(
                                            completed_articles
                                        ),
                                    "records":
                                        records,
                                    "updated_at":
                                        datetime.now()
                                        .isoformat(),
                                }
                            )

                            # Periodic Excel
                            if (
                                len(
                                    completed_articles
                                )
                                % 10 == 0
                            ):
                                save_excel(
                                    records
                                )

                            sleep_random(
                                ARTICLE_DELAY_MIN,
                                ARTICLE_DELAY_MAX
                            )

                        except Exception as exc:

                            print(
                                f"Article error: "
                                f"{exc}"
                            )

                            errors.append(
                                {
                                    "type":
                                        "article",
                                    "journal":
                                        journal[
                                            "name"
                                        ],
                                    "url":
                                        article_url,
                                    "error":
                                        str(exc),
                                }
                            )

                except Exception as exc:

                    print(
                        f"Issue error: "
                        f"{exc}"
                    )

                    errors.append(
                        {
                            "type":
                                "issue",
                            "journal":
                                journal[
                                    "name"
                                ],
                            "url":
                                issue_url,
                            "error":
                                str(exc),
                        }
                    )

                sleep_random(
                    ISSUE_DELAY_MIN,
                    ISSUE_DELAY_MAX
                )

        browser.close()

    # ========================================================
    # FINAL OUTPUT
    # ========================================================

    save_excel(
        records
    )

    save_json(
        ERROR_FILE,
        errors
    )

    save_json(
        CHECKPOINT_FILE,
        {
            "completed_articles":
                list(
                    completed_articles
                ),
            "records":
                records,
            "updated_at":
                datetime.now()
                .isoformat(),
        }
    )

    print()
    print("=" * 110)
    print("COMPLETED")
    print("=" * 110)

    print(
        f"Articles processed : "
        f"{len(completed_articles):,}"
    )

    print(
        f"Author records     : "
        f"{len(records):,}"
    )

    print(
        f"Errors             : "
        f"{len(errors):,}"
    )

    print(
        f"Excel              : "
        f"{os.path.abspath(OUTPUT_FILE)}"
    )

    print(
        f"Errors JSON        : "
        f"{os.path.abspath(ERROR_FILE)}"
    )


if __name__ == "__main__":
    main()