import os
import re
import json
import time
from datetime import datetime
from urllib.parse import urljoin

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.common.exceptions import (
    TimeoutException,
    WebDriverException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    NoSuchElementException,
)


# ============================================================
# CONFIGURATION
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

ERROR_FILE = os.path.join(
    OUTPUT_FOLDER,
    "errors.json"
)


REMOTE_DEBUGGING_ADDRESS = "127.0.0.1:9222"

PAGE_WAIT_SECONDS = 4

AUTHOR_PANEL_WAIT_SECONDS = 8

ARTICLE_DELAY_SECONDS = 3

ISSUE_DELAY_SECONDS = 2


# ============================================================
# JOURNALS
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
# REGEX
# ============================================================

EMAIL_RE = re.compile(
    r"\b"
    r"[A-Za-z0-9._%+\-]+"
    r"@"
    r"[A-Za-z0-9.\-]+"
    r"\."
    r"[A-Za-z]{2,}"
    r"\b"
)

YEAR_RE = re.compile(
    r"\b(20\d{2})\b"
)


# ============================================================
# BASIC HELPERS
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


def normalize_url(url):

    if not url:
        return ""

    url = url.split("#")[0]
    url = url.split("?")[0]

    return url.strip()


def extract_year(text):

    if not text:
        return None

    match = YEAR_RE.search(
        str(text)
    )

    if match:

        try:
            return int(
                match.group(1)
            )

        except Exception:
            pass

    return None


def save_json(path, data):

    with open(
        path,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            data,
            file,
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
        ) as file:

            return json.load(
                file
            )

    except Exception:
        return {}


# ============================================================
# CONNECT TO OPEN CHROME
# ============================================================

def connect_to_chrome():

    options = Options()

    options.add_experimental_option(
        "debuggerAddress",
        REMOTE_DEBUGGING_ADDRESS
    )

    print(
        f"Connecting to Chrome at "
        f"{REMOTE_DEBUGGING_ADDRESS}..."
    )

    driver = webdriver.Chrome(
        options=options
    )

    print(
        "Connected to existing Chrome session."
    )

    return driver


# ============================================================
# BLOCK / ERROR PAGE DETECTION
# ============================================================

def is_blocked(driver):

    try:

        source = (
            driver.page_source
            or ""
        ).lower()

    except Exception:

        return False

    indicators = [
        "there was a problem providing the content you requested",
        "cpe00002",
        "reference number:",
        "access denied",
        "unusual traffic",
    ]

    return any(
        indicator in source
        for indicator in indicators
    )


def wait_for_manual_access(driver):

    print()
    print("=" * 100)
    print("SCIENCEDIRECT BLOCK PAGE DETECTED")
    print("=" * 100)
    print()
    print(
        "Please resolve the page manually in the open Chrome window."
    )
    print(
        "When the normal ScienceDirect page becomes visible, "
        "the script will continue automatically."
    )
    print()

    while True:

        try:

            if not is_blocked(
                driver
            ):

                print(
                    "Normal ScienceDirect page detected. Continuing..."
                )

                return True

        except Exception:
            pass

        time.sleep(
            5
        )


# ============================================================
# PAGE OPEN
# ============================================================

def open_page(
    driver,
    url,
    wait_seconds=PAGE_WAIT_SECONDS
):

    print(
        f"Opening: {url}"
    )

    try:

        driver.get(
            url
        )

    except WebDriverException as exc:

        print(
            f"Browser error: {exc}"
        )

        return False

    time.sleep(
        wait_seconds
    )

    if is_blocked(
        driver
    ):

        wait_for_manual_access(
            driver
        )

    return True


# ============================================================
# SCROLL
# ============================================================

def scroll_page(
    driver,
    max_scrolls=15
):

    previous_height = 0

    for _ in range(
        max_scrolls
    ):

        try:

            current_height = (
                driver.execute_script(
                    "return document.body.scrollHeight"
                )
            )

            if (
                current_height
                == previous_height
            ):
                break

            driver.execute_script(
                """
                window.scrollTo(
                    0,
                    document.body.scrollHeight
                );
                """
            )

            time.sleep(
                0.8
            )

            previous_height = (
                current_height
            )

        except Exception:
            break


# ============================================================
# ISSUE PAGE
# ============================================================

def extract_issue_links(
    driver,
    journal
):

    issue_links = []

    print()
    print(
        f"Opening journal issues page:"
    )

    print(
        journal["url"]
    )

    if not open_page(
        driver,
        journal["url"]
    ):

        return []

    scroll_page(
        driver
    )

    anchors = driver.find_elements(
        By.CSS_SELECTOR,
        'a[href*="/vol/"][href*="/issue/"]'
    )

    print(
        f"Potential issue links found: "
        f"{len(anchors)}"
    )

    seen = set()

    for anchor in anchors:

        try:

            href = anchor.get_attribute(
                "href"
            )

            text = clean_text(
                anchor.text
            )

            if not href:
                continue

            href = normalize_url(
                href
            )

            if href in seen:
                continue

            possible_year = extract_year(
                text
            )

            # If the year is visible in link text,
            # filter immediately.
            if possible_year:

                if not (
                    journal["start_year"]
                    <= possible_year
                    <= journal["end_year"]
                ):

                    continue

            seen.add(
                href
            )

            issue_links.append(
                href
            )

        except Exception:
            continue

    return issue_links


# ============================================================
# ISSUE METADATA
# ============================================================

def extract_issue_metadata(
    driver
):

    result = {
        "journal": "",
        "year": "",
        "date": "",
        "volume": "",
        "issue": "",
    }

    # --------------------------------------------------------
    # JOURNAL
    # --------------------------------------------------------

    selectors = [
        ".publication-title",
        "h1.journal-title",
    ]

    for selector in selectors:

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector
            )

            if elements:

                text = clean_text(
                    elements[0].text
                )

                if text:

                    result[
                        "journal"
                    ] = text

                    break

        except Exception:
            pass

    # --------------------------------------------------------
    # PAGE TEXT
    # --------------------------------------------------------

    try:

        body_text = driver.find_element(
            By.TAG_NAME,
            "body"
        ).text

    except Exception:

        body_text = ""

    # Date: October 2026

    match = re.search(
        r"Date:\s*([A-Za-z]+\s+20\d{2})",
        body_text,
        re.I
    )

    if match:

        result[
            "date"
        ] = clean_text(
            match.group(1)
        )

    # Year

    year = extract_year(
        result["date"]
    )

    if not year:

        year = extract_year(
            body_text
        )

    if year:

        result[
            "year"
        ] = year

    # Volume

    match = re.search(
        r"\bVolume\s+(\d+)",
        body_text,
        re.I
    )

    if match:

        result[
            "volume"
        ] = match.group(1)

    # Issue

    match = re.search(
        r"\bIssue\s+(\d+)",
        body_text,
        re.I
    )

    if match:

        result[
            "issue"
        ] = match.group(1)

    return result


# ============================================================
# ARTICLE LINKS FROM ISSUE
# ============================================================

def extract_article_urls(
    driver
):

    urls = []

    seen = set()

    elements = driver.find_elements(
        By.CSS_SELECTOR,
        (
            'a.article-content-title'
            '[href*="/science/article/pii/"]'
        )
    )

    for element in elements:

        try:

            href = element.get_attribute(
                "href"
            )

            if not href:
                continue

            href = normalize_url(
                href
            )

            if href in seen:
                continue

            seen.add(
                href
            )

            urls.append(
                href
            )

        except Exception:
            continue

    return urls


# ============================================================
# ARTICLE TITLE
# ============================================================

def extract_article_title(
    driver
):

    selectors = [
        "h1.content-title .title-text",
        "#screen-reader-main-title .title-text",
        "h1.content-title",
        "h1",
    ]

    for selector in selectors:

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector
            )

            for element in elements:

                text = clean_text(
                    element.text
                )

                if text:
                    return text

        except Exception:
            continue

    return ""


# ============================================================
# DOI
# ============================================================

def extract_doi(
    driver
):

    try:

        links = driver.find_elements(
            By.CSS_SELECTOR,
            'a[href*="doi.org/10."]'
        )

        for link in links:

            href = link.get_attribute(
                "href"
            )

            if not href:
                continue

            match = re.search(
                r"10\.\d{4,9}/[^\s?#]+",
                href,
                re.I
            )

            if match:

                return (
                    match
                    .group(0)
                    .rstrip(
                        ".,;)"
                    )
                )

    except Exception:
        pass

    return ""


# ============================================================
# ARTICLE METADATA
# ============================================================

def extract_article_metadata(
    driver
):

    result = {
        "journal": "",
        "year": "",
        "date": "",
        "volume": "",
        "issue": "",
        "doi": "",
    }

    # --------------------------------------------------------
    # JOURNAL
    # --------------------------------------------------------

    try:

        journal_elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                ".publication-title"
            )
        )

        if journal_elements:

            result[
                "journal"
            ] = clean_text(
                journal_elements[
                    0
                ].text
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # METADATA SECTION
    # --------------------------------------------------------

    metadata_text = ""

    try:

        metadata_element = (
            driver.find_element(
                By.CSS_SELECTOR,
                ".publication-metadata"
            )
        )

        metadata_text = (
            metadata_element.text
        )

    except Exception:

        try:

            metadata_text = (
                driver.find_element(
                    By.TAG_NAME,
                    "body"
                ).text
            )

        except Exception:
            pass

    # Date

    date_match = re.search(
        r"Date:\s*([A-Za-z]+\s+20\d{2})",
        metadata_text,
        re.I
    )

    if date_match:

        result[
            "date"
        ] = clean_text(
            date_match.group(1)
        )

    # Year

    year = extract_year(
        result[
            "date"
        ]
    )

    if year:

        result[
            "year"
        ] = year

    # Volume

    volume_match = re.search(
        r"\bVolume\s+(\d+)",
        metadata_text,
        re.I
    )

    if volume_match:

        result[
            "volume"
        ] = volume_match.group(
            1
        )

    # Issue

    issue_match = re.search(
        r"\bIssue\s+(\d+)",
        metadata_text,
        re.I
    )

    if issue_match:

        result[
            "issue"
        ] = issue_match.group(
            1
        )

    # DOI

    result[
        "doi"
    ] = extract_doi(
        driver
    )

    return result


# ============================================================
# CLEAN AUTHOR
# ============================================================

def clean_author_name(
    name
):

    if not name:
        return ""

    name = clean_text(
        name
    )

    degrees = [
        "MD",
        "DM",
        "PhD",
        "Ph.D",
        "RN",
        "MSc",
        "MS",
        "MBBS",
        "MPH",
        "DNP",
        "FRCA",
        "FRCPC",
        "BSN",
        "MSN",
        "MBA",
    ]

    changed = True

    while changed:

        changed = False

        for degree in degrees:

            pattern = (
                r"[\s,]+"
                + re.escape(
                    degree
                )
                + r"\.?$"
            )

            new_name = re.sub(
                pattern,
                "",
                name,
                flags=re.I
            )

            new_name = (
                new_name
                .strip(
                    " ,"
                )
            )

            if (
                new_name
                != name
            ):

                name = new_name

                changed = True

    return clean_text(
        name
    )


# ============================================================
# AUTHOR BUTTONS
# ============================================================

def get_author_buttons(
    driver
):

    return driver.find_elements(
        By.CSS_SELECTOR,
        (
            'button'
            '[data-sd-ui-side-panel-opener="true"]'
            '[data-xocs-content-type="author"]'
        )
    )


def author_name_from_button(
    button
):

    given_name = ""
    surname = ""

    try:

        given_elements = (
            button.find_elements(
                By.CSS_SELECTOR,
                ".given-name"
            )
        )

        if given_elements:

            given_name = clean_text(
                given_elements[
                    0
                ].text
            )

    except Exception:
        pass

    try:

        surname_elements = (
            button.find_elements(
                By.CSS_SELECTOR,
                ".surname"
            )
        )

        if surname_elements:

            surname = clean_text(
                surname_elements[
                    0
                ].text
            )

    except Exception:
        pass

    if (
        given_name
        or surname
    ):

        return clean_author_name(
            f"{given_name} {surname}"
        )

    try:

        return clean_author_name(
            button.text
        )

    except Exception:

        return ""


# ============================================================
# AUTHOR PANEL DETECTION
# ============================================================

def find_author_panel(
    driver
):

    selectors = [
        '[role="dialog"]',
        '[class*="side-panel"]',
        '[class*="SidePanel"]',
        '[class*="drawer"]',
        '[data-testid*="side-panel"]',
    ]

    for selector in selectors:

        try:

            panels = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    selector
                )
            )

            for panel in panels:

                try:

                    if not panel.is_displayed():
                        continue

                    text = clean_text(
                        panel.text
                    )

                    if not text:
                        continue

                    if (
                        "Author" in text
                        or "Correspondence" in text
                        or EMAIL_RE.search(
                            text
                        )
                    ):

                        return panel

                except Exception:
                    continue

        except Exception:
            continue

    return None


def wait_for_author_panel(
    driver
):

    end_time = (
        time.time()
        + AUTHOR_PANEL_WAIT_SECONDS
    )

    while time.time() < end_time:

        panel = find_author_panel(
            driver
        )

        if panel is not None:

            return panel

        time.sleep(
            0.3
        )

    return None


# ============================================================
# EXTRACT EMAIL FROM PANEL
# ============================================================

def extract_email_from_panel(
    panel
):

    if panel is None:
        return ""

    # --------------------------------------------------------
    # MAILTO
    # --------------------------------------------------------

    try:

        links = panel.find_elements(
            By.CSS_SELECTOR,
            'a[href^="mailto:"]'
        )

        for link in links:

            href = (
                link.get_attribute(
                    "href"
                )
                or ""
            )

            email = (
                href
                .replace(
                    "mailto:",
                    ""
                )
                .split(
                    "?"
                )[0]
                .strip()
                .lower()
            )

            if EMAIL_RE.fullmatch(
                email
            ):

                return email

    except Exception:
        pass

    # --------------------------------------------------------
    # PANEL TEXT
    # --------------------------------------------------------

    try:

        text = panel.text

        matches = EMAIL_RE.findall(
            text
        )

        if matches:

            return (
                matches[0]
                .strip()
                .lower()
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # PANEL HTML
    # --------------------------------------------------------

    try:

        html = (
            panel.get_attribute(
                "outerHTML"
            )
            or ""
        )

        matches = EMAIL_RE.findall(
            html
        )

        if matches:

            return (
                matches[0]
                .strip()
                .lower()
            )

    except Exception:
        pass

    return ""


# ============================================================
# AFFILIATION
# ============================================================

def extract_affiliation_from_panel(
    panel,
    author_name,
    email
):

    if panel is None:
        return ""

    try:

        text = (
            panel.text
            or ""
        )

    except Exception:

        return ""

    lines = [
        clean_text(
            line
        )
        for line
        in text.splitlines()
        if clean_text(
            line
        )
    ]

    affiliation_lines = []

    for line in lines:

        lower = (
            line
            .lower()
        )

        if lower == "author":
            continue

        if (
            author_name
            and author_name.lower()
            == lower
        ):
            continue

        if (
            email
            and email.lower()
            == lower
        ):
            continue

        if "view in scopus" in lower:
            continue

        if "orcid" in lower:
            continue

        if lower.startswith(
            "more documents"
        ):
            break

        if (
            "provided by scopus"
            in lower
        ):
            break

        affiliation_lines.append(
            line
        )

    return clean_text(
        " ".join(
            affiliation_lines
        )
    )


# ============================================================
# CLOSE PANEL
# ============================================================

def close_author_panel(
    driver
):

    # --------------------------------------------------------
    # aria-label close
    # --------------------------------------------------------

    selectors = [
        'button[aria-label="Close"]',
        'button[aria-label="close"]',
        'button[title="Close"]',
        'button[title="close"]',
    ]

    for selector in selectors:

        try:

            buttons = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    selector
                )
            )

            for button in buttons:

                if button.is_displayed():

                    driver.execute_script(
                        "arguments[0].click();",
                        button
                    )

                    time.sleep(
                        0.4
                    )

                    return True

        except Exception:
            continue

    # --------------------------------------------------------
    # Search visible buttons for X
    # --------------------------------------------------------

    try:

        buttons = driver.find_elements(
            By.TAG_NAME,
            "button"
        )

        for button in reversed(
            buttons
        ):

            try:

                if not button.is_displayed():
                    continue

                aria = clean_text(
                    button.get_attribute(
                        "aria-label"
                    )
                ).lower()

                title = clean_text(
                    button.get_attribute(
                        "title"
                    )
                ).lower()

                text = clean_text(
                    button.text
                )

                if (
                    "close" in aria
                    or "close" in title
                    or text in [
                        "×",
                        "x"
                    ]
                ):

                    driver.execute_script(
                        "arguments[0].click();",
                        button
                    )

                    time.sleep(
                        0.4
                    )

                    return True

            except Exception:
                continue

    except Exception:
        pass

    # --------------------------------------------------------
    # ESCAPE
    # --------------------------------------------------------

    try:

        driver.find_element(
            By.TAG_NAME,
            "body"
        ).send_keys(
            Keys.ESCAPE
        )

        time.sleep(
            0.4
        )

        return True

    except Exception:
        pass

    return False


# ============================================================
# EXTRACT ALL AUTHORS + EMAIL
# ============================================================

def extract_article_authors(
    driver
):

    results = []

    initial_buttons = (
        get_author_buttons(
            driver
        )
    )

    total_authors = len(
        initial_buttons
    )

    print(
        f"    Author buttons found: "
        f"{total_authors}"
    )

    for index in range(
        total_authors
    ):

        # React can rebuild elements after each click.
        # Always reacquire author buttons.

        buttons = (
            get_author_buttons(
                driver
            )
        )

        if index >= len(
            buttons
        ):
            break

        button = buttons[
            index
        ]

        author_name = (
            author_name_from_button(
                button
            )
        )

        author_id = (
            button.get_attribute(
                "data-xocs-content-id"
            )
            or ""
        )

        print(
            f"      "
            f"[{index + 1}/{total_authors}] "
            f"{author_name} "
            f"({author_id})"
        )

        email = ""
        affiliation = ""

        try:

            driver.execute_script(
                """
                arguments[0].scrollIntoView({
                    block: 'center',
                    behavior: 'instant'
                });
                """,
                button
            )

            time.sleep(
                0.3
            )

            # -----------------------------------------------
            # CLICK
            # -----------------------------------------------

            try:

                button.click()

            except (
                ElementClickInterceptedException,
                StaleElementReferenceException
            ):

                buttons = (
                    get_author_buttons(
                        driver
                    )
                )

                if index < len(
                    buttons
                ):

                    button = (
                        buttons[
                            index
                        ]
                    )

                    driver.execute_script(
                        "arguments[0].click();",
                        button
                    )

            # -----------------------------------------------
            # WAIT PANEL
            # -----------------------------------------------

            panel = (
                wait_for_author_panel(
                    driver
                )
            )

            if panel:

                email = (
                    extract_email_from_panel(
                        panel
                    )
                )

                affiliation = (
                    extract_affiliation_from_panel(
                        panel,
                        author_name,
                        email
                    )
                )

            if email:

                print(
                    f"          EMAIL: "
                    f"{email}"
                )

            else:

                print(
                    "          EMAIL: -"
                )

        except Exception as exc:

            print(
                f"          Author panel error: "
                f"{exc}"
            )

        finally:

            close_author_panel(
                driver
            )

            time.sleep(
                0.5
            )

        results.append(
            {
                "author":
                    author_name,

                "email":
                    email,

                "affiliation":
                    affiliation,

                "author_id":
                    author_id,
            }
        )

    return results


# ============================================================
# PARSE ARTICLE
# ============================================================

def parse_article(
    driver,
    article_url,
    expected_journal=None
):

    print()
    print(
        "-" * 100
    )

    print(
        f"ARTICLE: {article_url}"
    )

    if not open_page(
        driver,
        article_url
    ):

        return []

    # Wait for author area.

    try:

        WebDriverWait(
            driver,
            10
        ).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "#author-group"
                )
            )
        )

    except TimeoutException:

        pass

    title = (
        extract_article_title(
            driver
        )
    )

    metadata = (
        extract_article_metadata(
            driver
        )
    )

    if (
        not metadata[
            "journal"
        ]
        and expected_journal
    ):

        metadata[
            "journal"
        ] = expected_journal

    print(
        f"    Title  : {title}"
    )

    print(
        f"    Journal: "
        f"{metadata['journal']}"
    )

    print(
        f"    Date   : "
        f"{metadata['date']}"
    )

    print(
        f"    Volume : "
        f"{metadata['volume']}"
    )

    print(
        f"    Issue  : "
        f"{metadata['issue']}"
    )

    print(
        f"    DOI    : "
        f"{metadata['doi']}"
    )

    authors = (
        extract_article_authors(
            driver
        )
    )

    rows = []

    for author in authors:

        rows.append(
            {
                "journal":
                    metadata[
                        "journal"
                    ],

                "year":
                    metadata[
                        "year"
                    ],

                "date":
                    metadata[
                        "date"
                    ],

                "volume":
                    metadata[
                        "volume"
                    ],

                "issue":
                    metadata[
                        "issue"
                    ],

                "article_title":
                    title,

                "author":
                    author[
                        "author"
                    ],

                "email":
                    author[
                        "email"
                    ],

                "affiliation":
                    author[
                        "affiliation"
                    ],

                "author_id":
                    author[
                        "author_id"
                    ],

                "doi":
                    metadata[
                        "doi"
                    ],

                "article_url":
                    article_url,
            }
        )

    return rows


# ============================================================
# CHECKPOINT
# ============================================================

def load_checkpoint():

    data = load_json(
        CHECKPOINT_FILE
    )

    records = data.get(
        "records",
        []
    )

    completed_articles = set(
        data.get(
            "completed_articles",
            []
        )
    )

    return (
        records,
        completed_articles
    )


def save_checkpoint(
    records,
    completed_articles
):

    save_json(
        CHECKPOINT_FILE,
        {
            "updated_at":
                datetime.now()
                .isoformat(),

            "records":
                records,

            "completed_articles":
                sorted(
                    completed_articles
                ),
        }
    )


# ============================================================
# EXCEL
# ============================================================

def save_excel(
    records
):

    if not records:

        print(
            "No records to save."
        )

        return

    df = pd.DataFrame(
        records
    )

    required_columns = [
        "journal",
        "year",
        "date",
        "volume",
        "issue",
        "article_title",
        "author",
        "email",
        "affiliation",
        "author_id",
        "doi",
        "article_url",
    ]

    for column in required_columns:

        if column not in df.columns:

            df[
                column
            ] = ""

    df = df[
        required_columns
    ]

    # Remove duplicate author/article records

    df.drop_duplicates(
        subset=[
            "article_url",
            "author"
        ],
        keep="first",
        inplace=True
    )

    # Prioritize rows with emails when sorting

    df[
        "_has_email"
    ] = (
        df[
            "email"
        ]
        .astype(str)
        .str.strip()
        .ne("")
    )

    df.sort_values(
        by=[
            "journal",
            "year",
            "volume",
            "issue",
            "article_title",
            "_has_email",
            "author",
        ],
        ascending=[
            True,
            True,
            True,
            True,
            True,
            False,
            True,
        ],
        inplace=True
    )

    df.drop(
        columns=[
            "_has_email"
        ],
        inplace=True
    )

    # --------------------------------------------------------
    # EMAIL ONLY SHEET
    # --------------------------------------------------------

    email_df = (
        df[
            df[
                "email"
            ]
            .astype(str)
            .str.strip()
            .ne("")
        ]
        .copy()
    )

    # Remove duplicate email addresses

    email_df.drop_duplicates(
        subset=[
            "email"
        ],
        keep="first",
        inplace=True
    )

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="All Authors",
            index=False
        )

        email_df.to_excel(
            writer,
            sheet_name="Authors With Email",
            index=False
        )

        workbook = (
            writer.book
        )

        for sheet_name in [
            "All Authors",
            "Authors With Email",
        ]:

            worksheet = (
                workbook[
                    sheet_name
                ]
            )

            worksheet.freeze_panes = (
                "A2"
            )

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            widths = {
                "A": 45,
                "B": 10,
                "C": 18,
                "D": 10,
                "E": 10,
                "F": 70,
                "G": 35,
                "H": 45,
                "I": 80,
                "J": 15,
                "K": 35,
                "L": 85,
            }

            for column, width in (
                widths.items()
            ):

                worksheet.column_dimensions[
                    column
                ].width = width

    print()
    print(
        f"Excel saved: "
        f"{os.path.abspath(OUTPUT_FILE)}"
    )

    print(
        f"Total author rows: "
        f"{len(df):,}"
    )

    print(
        f"Rows with email: "
        f"{len(email_df):,}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    records, completed_articles = (
        load_checkpoint()
    )

    errors = load_json(
        ERROR_FILE
    )

    if not isinstance(
        errors,
        list
    ):
        errors = []

    print("=" * 110)
    print("SCIENCEDIRECT AUTHOR + EMAIL SCRAPER")
    print("=" * 110)

    print()
    print(
        f"Existing records : "
        f"{len(records):,}"
    )

    print(
        f"Completed articles: "
        f"{len(completed_articles):,}"
    )

    print()

    driver = connect_to_chrome()

    try:

        for journal_index, journal in enumerate(
            JOURNALS,
            start=1
        ):

            print()
            print("=" * 110)

            print(
                f"[{journal_index}/"
                f"{len(JOURNALS)}] "
                f"{journal['name']}"
            )

            print(
                f"Years: "
                f"{journal['start_year']} - "
                f"{journal['end_year']}"
            )

            print("=" * 110)

            # =================================================
            # ISSUES
            # =================================================

            issue_links = (
                extract_issue_links(
                    driver,
                    journal
                )
            )

            print(
                f"Issue links collected: "
                f"{len(issue_links)}"
            )

            for issue_index, issue_url in enumerate(
                issue_links,
                start=1
            ):

                print()
                print(
                    f"[Issue "
                    f"{issue_index}/"
                    f"{len(issue_links)}]"
                )

                print(
                    issue_url
                )

                try:

                    if not open_page(
                        driver,
                        issue_url
                    ):

                        continue

                    time.sleep(
                        ISSUE_DELAY_SECONDS
                    )

                    issue_metadata = (
                        extract_issue_metadata(
                            driver
                        )
                    )

                    issue_year = (
                        issue_metadata.get(
                            "year"
                        )
                    )

                    print(
                        f"Issue year: "
                        f"{issue_year}"
                    )

                    # ------------------------------------------
                    # YEAR FILTER
                    # ------------------------------------------

                    if issue_year:

                        try:

                            issue_year_int = int(
                                issue_year
                            )

                        except Exception:

                            issue_year_int = None

                        if issue_year_int:

                            if not (
                                journal[
                                    "start_year"
                                ]
                                <= issue_year_int
                                <= journal[
                                    "end_year"
                                ]
                            ):

                                print(
                                    "Skipping issue outside "
                                    "requested year range."
                                )

                                continue

                    # ------------------------------------------
                    # ARTICLE LINKS
                    # ------------------------------------------

                    article_urls = (
                        extract_article_urls(
                            driver
                        )
                    )

                    print(
                        f"Articles found: "
                        f"{len(article_urls)}"
                    )

                    for article_index, article_url in enumerate(
                        article_urls,
                        start=1
                    ):

                        article_url = (
                            normalize_url(
                                article_url
                            )
                        )

                        if (
                            article_url
                            in completed_articles
                        ):

                            print(
                                f"[{article_index}/"
                                f"{len(article_urls)}] "
                                f"Already processed"
                            )

                            continue

                        print()
                        print(
                            f"[Article "
                            f"{article_index}/"
                            f"{len(article_urls)}]"
                        )

                        try:

                            article_records = (
                                parse_article(
                                    driver,
                                    article_url,
                                    expected_journal=
                                        journal[
                                            "name"
                                        ]
                                )
                            )

                            # ----------------------------------
                            # FINAL YEAR FILTER
                            # ----------------------------------

                            accepted_records = []

                            for record in article_records:

                                year = record.get(
                                    "year"
                                )

                                if year:

                                    try:

                                        year_int = int(
                                            year
                                        )

                                    except Exception:

                                        year_int = None

                                    if year_int:

                                        if not (
                                            journal[
                                                "start_year"
                                            ]
                                            <= year_int
                                            <= journal[
                                                "end_year"
                                            ]
                                        ):

                                            continue

                                accepted_records.append(
                                    record
                                )

                            records.extend(
                                accepted_records
                            )

                            completed_articles.add(
                                article_url
                            )

                            # ----------------------------------
                            # SAVE CHECKPOINT AFTER EVERY ARTICLE
                            # ----------------------------------

                            save_checkpoint(
                                records,
                                completed_articles
                            )

                            # ----------------------------------
                            # SAVE EXCEL EVERY 10 ARTICLES
                            # ----------------------------------

                            if (
                                len(
                                    completed_articles
                                )
                                % 10 == 0
                            ):

                                save_excel(
                                    records
                                )

                            time.sleep(
                                ARTICLE_DELAY_SECONDS
                            )

                        except Exception as exc:

                            print(
                                f"ARTICLE ERROR: "
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

                                    "timestamp":
                                        datetime.now()
                                        .isoformat(),
                                }
                            )

                            save_json(
                                ERROR_FILE,
                                errors
                            )

                except Exception as exc:

                    print(
                        f"ISSUE ERROR: "
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

                            "timestamp":
                                datetime.now()
                                .isoformat(),
                        }
                    )

                    save_json(
                        ERROR_FILE,
                        errors
                    )

    except KeyboardInterrupt:

        print()
        print(
            "Stopped manually."
        )

    finally:

        # =====================================================
        # FINAL SAVE
        # =====================================================

        save_checkpoint(
            records,
            completed_articles
        )

        save_json(
            ERROR_FILE,
            errors
        )

        save_excel(
            records
        )

        print()
        print("=" * 110)
        print("FINISHED")
        print("=" * 110)

        print(
            f"Completed articles : "
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
            f"Checkpoint         : "
            f"{os.path.abspath(CHECKPOINT_FILE)}"
        )

        print(
            f"Errors JSON        : "
            f"{os.path.abspath(ERROR_FILE)}"
        )


if __name__ == "__main__":
    main()