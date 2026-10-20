import os
import re
import time
import json
from urllib.parse import urljoin

import pandas as pd

from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException


# ============================================================
# CONFIG
# ============================================================

OUTPUT_FOLDER = "sciencedirect_output"

OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "sciencedirect_authors_emails.xlsx"
)

CHECKPOINT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "checkpoint.json"
)

WAIT_BETWEEN_ARTICLES = 4


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
# CONNECT TO EXISTING CHROME
# ============================================================

def connect_chrome():

    options = Options()

    options.add_experimental_option(
        "debuggerAddress",
        "127.0.0.1:9222"
    )

    driver = webdriver.Chrome(
        options=options
    )

    return driver


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
    ).strip()


def extract_year(text):

    match = re.search(
        r"\b(20\d{2})\b",
        text or ""
    )

    if match:
        return int(
            match.group(1)
        )

    return None


# ============================================================
# BLOCK DETECTION
# ============================================================

def is_blocked(driver):

    source = driver.page_source.lower()

    indicators = [
        "there was a problem providing the content you requested",
        "cpe00002",
        "reference number:",
        "access denied",
        "unusual traffic",
    ]

    return any(
        text in source
        for text in indicators
    )


def wait_for_manual_access(driver):

    while is_blocked(driver):

        print()
        print("=" * 90)

        print(
            "SCIENCEDIRECT BLOCK PAGE DETECTED"
        )

        print(
            "Resolve/reload the page manually in Chrome."
        )

        print(
            "The scraper will automatically continue "
            "when the article becomes visible."
        )

        print("=" * 90)

        time.sleep(10)

        try:
            driver.refresh()

        except Exception:
            pass

        time.sleep(5)


# ============================================================
# PAGE LOAD
# ============================================================

def open_page(driver, url):

    print()
    print(
        "Opening:",
        url
    )

    try:

        driver.get(
            url
        )

    except WebDriverException as exc:

        print(
            "Browser error:",
            exc
        )

        return False

    time.sleep(4)

    if is_blocked(driver):

        wait_for_manual_access(
            driver
        )

    return not is_blocked(
        driver
    )


# ============================================================
# SCROLL
# ============================================================

def scroll_page(driver):

    previous_height = 0

    for _ in range(20):

        try:

            current_height = (
                driver.execute_script(
                    "return document.body.scrollHeight"
                )
            )

            if current_height == previous_height:
                break

            driver.execute_script(
                """
                window.scrollTo(
                    0,
                    document.body.scrollHeight
                );
                """
            )

            time.sleep(1)

            previous_height = (
                current_height
            )

        except Exception:
            break


# ============================================================
# ARTICLE URL EXTRACTION
# ============================================================

def extract_article_urls(driver):

    soup = BeautifulSoup(
        driver.page_source,
        "lxml"
    )

    links = {}

    for anchor in soup.find_all(
        "a",
        href=True
    ):

        href = anchor.get(
            "href",
            ""
        )

        if "/science/article/pii/" not in href:
            continue

        url = urljoin(
            "https://www.sciencedirect.com",
            href
        )

        url = url.split(
            "?"
        )[0]

        url = url.split(
            "#"
        )[0]

        links[url] = url

    return list(
        links.values()
    )


# ============================================================
# EMAIL EXTRACTION
# ============================================================

EMAIL_REGEX = re.compile(
    r"""
    \b
    [A-Za-z0-9._%+\-]+
    @
    [A-Za-z0-9.\-]+
    \.
    [A-Za-z]{2,}
    \b
    """,
    re.VERBOSE
)


def extract_emails(soup):

    emails = []

    # --------------------------------------------------------
    # mailto:
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

            email = (
                href
                .split(":", 1)[1]
                .split("?", 1)[0]
                .strip()
                .lower()
            )

            if (
                email
                and email not in emails
            ):
                emails.append(
                    email
                )

    # --------------------------------------------------------
    # visible page text
    # --------------------------------------------------------

    text = soup.get_text(
        " ",
        strip=True
    )

    found = EMAIL_REGEX.findall(
        text
    )

    for email in found:

        email = email.lower()

        if email not in emails:
            emails.append(
                email
            )

    # --------------------------------------------------------
    # complete HTML
    #
    # Sometimes email information exists in attributes / JSON
    # but isn't directly visible.
    # --------------------------------------------------------

    html_source = str(
        soup
    )

    found = EMAIL_REGEX.findall(
        html_source
    )

    for email in found:

        email = (
            email
            .replace("\\u0040", "@")
            .lower()
        )

        if email not in emails:
            emails.append(
                email
            )

    return emails


# ============================================================
# AUTHORS
# ============================================================

def extract_authors(soup):

    authors = []

    # citation_author is generally the safest source

    for meta in soup.find_all(
        "meta",
        attrs={
            "name": "citation_author"
        }
    ):

        author = clean_text(
            meta.get(
                "content",
                ""
            )
        )

        if (
            author
            and author not in authors
        ):
            authors.append(
                author
            )

    # DC.creator fallback

    if not authors:

        for meta in soup.find_all(
            "meta",
            attrs={
                "name": re.compile(
                    r"dc\.creator",
                    re.I
                )
            }
        ):

            author = clean_text(
                meta.get(
                    "content",
                    ""
                )
            )

            if (
                author
                and author not in authors
            ):
                authors.append(
                    author
                )

    return authors


# ============================================================
# AUTHOR / EMAIL MATCHING
# ============================================================

def normalize_name(value):

    value = (
        value
        .lower()
        .strip()
    )

    value = re.sub(
        r"[^a-z]",
        "",
        value
    )

    return value


def email_username(email):

    return (
        email
        .split("@", 1)[0]
        .lower()
    )


def email_name_score(
    author,
    email
):

    author_compact = normalize_name(
        author
    )

    username = email_username(
        email
    )

    username_compact = re.sub(
        r"[^a-z]",
        "",
        username
    )

    if not author_compact:
        return 0

    if not username_compact:
        return 0

    score = 0

    author_parts = re.findall(
        r"[A-Za-z]+",
        author.lower()
    )

    # Full name appears
    if author_compact in username_compact:
        score += 100

    # Username appears inside author
    if username_compact in author_compact:
        score += 80

    # Individual name parts
    for part in author_parts:

        if len(part) < 2:
            continue

        if part in username_compact:
            score += 20

    # First-letter initial matching
    if author_parts:

        first = author_parts[0]
        last = author_parts[-1]

        if username.startswith(
            first[:1]
        ):
            score += 5

        if last in username:
            score += 30

    return score


def match_author_emails(
    authors,
    emails
):

    results = []

    used_emails = set()

    for author in authors:

        best_email = ""
        best_score = 0

        for email in emails:

            if email in used_emails:
                continue

            score = email_name_score(
                author,
                email
            )

            if score > best_score:

                best_score = score
                best_email = email

        # Don't force weak guesses
        if best_score >= 20:

            used_emails.add(
                best_email
            )

        else:

            best_email = ""

        results.append(
            {
                "author": author,
                "email": best_email,
                "email_match_score":
                    best_score,
            }
        )

    return results


# ============================================================
# ARTICLE METADATA
# ============================================================

def meta_value(
    soup,
    name
):

    meta = soup.find(
        "meta",
        attrs={
            "name": name
        }
    )

    if meta:

        return clean_text(
            meta.get(
                "content",
                ""
            )
        )

    return ""


def parse_article(
    driver,
    article_url
):

    if not open_page(
        driver,
        article_url
    ):
        return []

    # Some author details load after initial article content
    time.sleep(2)

    scroll_page(
        driver
    )

    soup = BeautifulSoup(
        driver.page_source,
        "lxml"
    )

    title = meta_value(
        soup,
        "citation_title"
    )

    journal = meta_value(
        soup,
        "citation_journal_title"
    )

    doi = meta_value(
        soup,
        "citation_doi"
    )

    volume = meta_value(
        soup,
        "citation_volume"
    )

    issue = meta_value(
        soup,
        "citation_issue"
    )

    date_value = (
        meta_value(
            soup,
            "citation_publication_date"
        )
        or
        meta_value(
            soup,
            "citation_date"
        )
    )

    year = extract_year(
        date_value
    )

    authors = extract_authors(
        soup
    )

    emails = extract_emails(
        soup
    )

    print(
        f"    Authors : {len(authors)}"
    )

    print(
        f"    Emails  : {len(emails)}"
    )

    matches = match_author_emails(
        authors,
        emails
    )

    rows = []

    for match in matches:

        rows.append(
            {
                "journal":
                    journal,

                "year":
                    year or "",

                "volume":
                    volume,

                "issue":
                    issue,

                "article_title":
                    title,

                "author":
                    match["author"],

                "email":
                    match["email"],

                "email_match_score":
                    match[
                        "email_match_score"
                    ],

                "all_found_emails":
                    "; ".join(
                        emails
                    ),

                "doi":
                    doi,

                "article_url":
                    article_url,
            }
        )

    # If no author metadata exists but email exists
    if not authors:

        for email in emails:

            rows.append(
                {
                    "journal":
                        journal,

                    "year":
                        year or "",

                    "volume":
                        volume,

                    "issue":
                        issue,

                    "article_title":
                        title,

                    "author":
                        "",

                    "email":
                        email,

                    "email_match_score":
                        0,

                    "all_found_emails":
                        "; ".join(
                            emails
                        ),

                    "doi":
                        doi,

                    "article_url":
                        article_url,
                }
            )

    return rows


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(records):

    if not records:
        return

    df = pd.DataFrame(
        records
    )

    df = df.drop_duplicates(
        subset=[
            "article_url",
            "author",
            "email"
        ],
        keep="first"
    )

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Authors",
            index=False
        )

        ws = writer.book[
            "Authors"
        ]

        ws.freeze_panes = "A2"

        ws.auto_filter.ref = (
            ws.dimensions
        )

        widths = {
            "A": 45,
            "B": 12,
            "C": 12,
            "D": 12,
            "E": 70,
            "F": 35,
            "G": 45,
            "H": 20,
            "I": 80,
            "J": 35,
            "K": 80,
        }

        for column, width in widths.items():

            ws.column_dimensions[
                column
            ].width = width


# ============================================================
# CHECKPOINT
# ============================================================

def save_checkpoint(
    records,
    completed
):

    with open(
        CHECKPOINT_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            {
                "records":
                    records,

                "completed":
                    list(
                        completed
                    )
            },
            f,
            ensure_ascii=False,
            indent=2
        )


def load_checkpoint():

    if not os.path.exists(
        CHECKPOINT_FILE
    ):
        return [], set()

    try:

        with open(
            CHECKPOINT_FILE,
            "r",
            encoding="utf-8"
        ) as f:

            data = json.load(
                f
            )

        return (
            data.get(
                "records",
                []
            ),
            set(
                data.get(
                    "completed",
                    []
                )
            )
        )

    except Exception:

        return [], set()


# ============================================================
# MAIN
# ============================================================

def main():

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    records, completed = (
        load_checkpoint()
    )

    driver = connect_chrome()

    print()
    print("=" * 100)
    print("SCIENCEDIRECT EMAIL SCRAPER")
    print("USING EXISTING CHROME SESSION")
    print("=" * 100)

    print(
        f"Existing records: "
        f"{len(records):,}"
    )

    for journal in JOURNALS:

        print()
        print("=" * 100)

        print(
            journal["name"]
        )

        print(
            f"{journal['start_year']} - "
            f"{journal['end_year']}"
        )

        print("=" * 100)

        if not open_page(
            driver,
            journal["url"]
        ):
            continue

        scroll_page(
            driver
        )

        issue_source = BeautifulSoup(
            driver.page_source,
            "lxml"
        )

        # Get possible issue links
        issue_links = []

        for a in issue_source.find_all(
            "a",
            href=True
        ):

            href = a[
                "href"
            ]

            text = clean_text(
                a.get_text(
                    " ",
                    strip=True
                )
            )

            year = extract_year(
                text + " " + href
            )

            if year:

                if not (
                    journal[
                        "start_year"
                    ]
                    <= year
                    <= journal[
                        "end_year"
                    ]
                ):
                    continue

            if (
                "/issue/" in href
                or "/vol/" in href
            ):

                url = urljoin(
                    "https://www.sciencedirect.com",
                    href
                )

                if url not in issue_links:

                    issue_links.append(
                        url
                    )

        print(
            f"Issues found: "
            f"{len(issue_links)}"
        )

        for issue_url in issue_links:

            if not open_page(
                driver,
                issue_url
            ):
                continue

            scroll_page(
                driver
            )

            article_urls = (
                extract_article_urls(
                    driver
                )
            )

            print(
                f"Articles: "
                f"{len(article_urls)}"
            )

            for index, article_url in enumerate(
                article_urls,
                1
            ):

                if article_url in completed:

                    continue

                print()
                print(
                    f"[{index}/"
                    f"{len(article_urls)}]"
                )

                try:

                    article_records = (
                        parse_article(
                            driver,
                            article_url
                        )
                    )

                    # Year filtering
                    valid_records = []

                    for row in article_records:

                        year = row.get(
                            "year"
                        )

                        if year:

                            if not (
                                journal[
                                    "start_year"
                                ]
                                <= int(year)
                                <= journal[
                                    "end_year"
                                ]
                            ):
                                continue

                        valid_records.append(
                            row
                        )

                    records.extend(
                        valid_records
                    )

                    completed.add(
                        article_url
                    )

                    save_checkpoint(
                        records,
                        completed
                    )

                    if (
                        len(completed)
                        % 10 == 0
                    ):

                        save_excel(
                            records
                        )

                    time.sleep(
                        WAIT_BETWEEN_ARTICLES
                    )

                except Exception as exc:

                    print(
                        "ERROR:",
                        exc
                    )

    save_excel(
        records
    )

    save_checkpoint(
        records,
        completed
    )

    print()
    print("=" * 100)
    print("DONE")
    print("=" * 100)

    print(
        f"Records: "
        f"{len(records):,}"
    )

    print(
        f"Excel: "
        f"{OUTPUT_FILE}"
    )


if __name__ == "__main__":
    main()