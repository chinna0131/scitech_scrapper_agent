import os
import re
import json
import time
import random
from datetime import datetime

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
    NoSuchWindowException,
    InvalidSessionIdException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
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


# ------------------------------------------------------------
# TIMING
# ------------------------------------------------------------

PAGE_LOAD_WAIT = 4

ISSUE_LOAD_WAIT = 4

AUTHOR_PANEL_OPEN_TIMEOUT = 15

AUTHOR_EMAIL_TIMEOUT = 25

AUTHOR_EMAIL_POLL_INTERVAL = 0.5

ARTICLE_DELAY_MIN = 6

ARTICLE_DELAY_MAX = 12

COOLDOWN_AFTER_ARTICLES = 10

COOLDOWN_SECONDS = 60

SAVE_EVERY_N_ARTICLES = 5


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
# CUSTOM EXCEPTIONS
# ============================================================

class BrowserSessionLost(Exception):
    pass


class ScienceDirectBlocked(Exception):
    pass


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_text(value):

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
    ).strip()


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

    if not match:
        return None

    try:

        return int(
            match.group(1)
        )

    except Exception:

        return None


def polite_delay():

    seconds = random.uniform(
        ARTICLE_DELAY_MIN,
        ARTICLE_DELAY_MAX
    )

    print(
        f"Waiting {seconds:.1f}s..."
    )

    time.sleep(
        seconds
    )


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
# CHROME CONNECTION
# ============================================================

def connect_to_chrome():

    print(
        f"Connecting to Chrome at "
        f"{REMOTE_DEBUGGING_ADDRESS}..."
    )

    options = Options()

    options.add_experimental_option(
        "debuggerAddress",
        REMOTE_DEBUGGING_ADDRESS
    )

    driver = webdriver.Chrome(
        options=options
    )

    handles = driver.window_handles

    if not handles:

        raise BrowserSessionLost(
            "Chrome is running but no browser window is available."
        )

    driver.switch_to.window(
        handles[-1]
    )

    print(
        "Connected to existing Chrome session."
    )

    return driver


# ============================================================
# DRIVER HEALTH
# ============================================================

def driver_is_alive(driver):

    if driver is None:
        return False

    try:

        handles = driver.window_handles

        if not handles:
            return False

        driver.switch_to.window(
            handles[-1]
        )

        _ = driver.current_url

        return True

    except (
        NoSuchWindowException,
        InvalidSessionIdException,
        WebDriverException,
    ):

        return False

    except Exception:

        return False


def ensure_driver_alive(driver):

    if not driver_is_alive(
        driver
    ):

        raise BrowserSessionLost(
            "Chrome/Selenium browser window is no longer available."
        )


# ============================================================
# SCIENCEDIRECT BLOCK DETECTION
# ============================================================

def is_sciencedirect_blocked(driver):

    try:

        html = (
            driver.page_source
            or ""
        ).lower()

        title = (
            driver.title
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

    for marker in indicators:

        if (
            marker in html
            or marker in title
        ):

            return True

    return False


# ============================================================
# CHECK NORMAL SCIENCEDIRECT CONTENT
# ============================================================

def is_normal_sciencedirect_page(driver):

    try:

        current_url = (
            driver.current_url
            or ""
        ).lower()

        html = (
            driver.page_source
            or ""
        ).lower()

        if (
            "sciencedirect.com"
            not in current_url
        ):

            return False

        indicators = [
            "publication-title",
            "author-group",
            "article-content-title",
            "journal-title",
        ]

        return any(
            indicator in html
            for indicator in indicators
        )

    except Exception:

        return False


# ============================================================
# MANUAL BLOCK RECOVERY
# ============================================================

def wait_for_manual_sciencedirect_recovery(
    driver,
    records,
    completed_articles
):

    print()
    print(
        "=" * 100
    )

    print(
        "SCIENCEDIRECT BLOCK DETECTED - SCRAPER PAUSED"
    )

    print(
        "=" * 100
    )

    print()

    print(
        "Saving checkpoint before pausing..."
    )

    save_checkpoint(
        records,
        completed_articles
    )

    save_excel(
        records
    )

    print()
    print(
        "IMPORTANT:"
    )

    print(
        "Do NOT close the Chrome window."
    )

    print(
        "Do NOT keep refreshing repeatedly."
    )

    print()
    print(
        "In Chrome manually restore normal ScienceDirect access."
    )

    print(
        "Once a normal ScienceDirect article/journal page is visible,"
    )

    print(
        "the scraper will continue automatically."
    )

    print()

    while True:

        if not driver_is_alive(
            driver
        ):

            raise BrowserSessionLost(
                "Chrome window was closed while scraper was paused."
            )

        if (
            not is_sciencedirect_blocked(
                driver
            )
            and
            is_normal_sciencedirect_page(
                driver
            )
        ):

            print()
            print(
                "ScienceDirect access restored."
            )

            print(
                "Resuming scraper..."
            )

            time.sleep(
                3
            )

            return True

        time.sleep(
            5
        )


# ============================================================
# OPEN PAGE
# ============================================================

def open_page(
    driver,
    url,
    records,
    completed_articles,
    wait_seconds=PAGE_LOAD_WAIT
):

    ensure_driver_alive(
        driver
    )

    print(
        f"Opening: {url}"
    )

    try:

        driver.get(
            url
        )

    except (
        NoSuchWindowException,
        InvalidSessionIdException,
    ) as exc:

        raise BrowserSessionLost(
            str(exc)
        )

    except WebDriverException as exc:

        if not driver_is_alive(
            driver
        ):

            raise BrowserSessionLost(
                str(exc)
            )

        raise

    time.sleep(
        wait_seconds
    )

    ensure_driver_alive(
        driver
    )

    # ========================================================
    # BLOCK DETECTED
    # ========================================================

    if is_sciencedirect_blocked(
        driver
    ):

        wait_for_manual_sciencedirect_recovery(
            driver,
            records,
            completed_articles
        )

        # ----------------------------------------------------
        # User probably moved to another page manually.
        # Retry requested URL ONCE.
        # ----------------------------------------------------

        print(
            "Retrying requested URL once..."
        )

        driver.get(
            url
        )

        time.sleep(
            wait_seconds
        )

        ensure_driver_alive(
            driver
        )

        if is_sciencedirect_blocked(
            driver
        ):

            raise ScienceDirectBlocked(
                "ScienceDirect blocked the requested URL again."
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

        ensure_driver_alive(
            driver
        )

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

        except WebDriverException:

            break


# ============================================================
# ISSUE LINKS
# ============================================================

def extract_issue_links(
    driver,
    journal,
    records,
    completed_articles
):

    print()
    print(
        "Opening journal issues page:"
    )

    print(
        journal[
            "url"
        ]
    )

    open_page(
        driver,
        journal[
            "url"
        ],
        records,
        completed_articles,
        PAGE_LOAD_WAIT
    )

    scroll_page(
        driver
    )

    links = []

    seen = set()

    anchors = driver.find_elements(
        By.CSS_SELECTOR,
        'a[href*="/vol/"][href*="/issue/"]'
    )

    print(
        f"Potential issue links found: "
        f"{len(anchors)}"
    )

    for anchor in anchors:

        try:

            href = (
                anchor.get_attribute(
                    "href"
                )
                or ""
            )

            if not href:
                continue

            href = normalize_url(
                href
            )

            if href in seen:
                continue

            text = clean_text(
                anchor.text
            )

            possible_year = (
                extract_year(
                    text
                )
            )

            if possible_year:

                if not (
                    journal[
                        "start_year"
                    ]
                    <= possible_year
                    <= journal[
                        "end_year"
                    ]
                ):

                    continue

            seen.add(
                href
            )

            links.append(
                href
            )

        except Exception:

            continue

    return links


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

    try:

        elements = driver.find_elements(
            By.CSS_SELECTOR,
            ".publication-title"
        )

        if elements:

            result[
                "journal"
            ] = clean_text(
                elements[
                    0
                ].text
            )

    except Exception:
        pass

    try:

        body_text = (
            driver.find_element(
                By.TAG_NAME,
                "body"
            ).text
        )

    except Exception:

        body_text = ""

    # --------------------------------------------------------
    # ISSUE URL
    # --------------------------------------------------------

    try:

        url = (
            driver.current_url
            or ""
        )

    except Exception:

        url = ""

    # --------------------------------------------------------
    # YEAR
    # --------------------------------------------------------

    result[
        "year"
    ] = (
        extract_year(
            body_text
        )
        or ""
    )

    # --------------------------------------------------------
    # URL VOLUME/ISSUE IS VERY RELIABLE
    # --------------------------------------------------------

    match = re.search(
        r"/vol/(\d+)/issue/(\d+)",
        url,
        re.I
    )

    if match:

        result[
            "volume"
        ] = match.group(
            1
        )

        result[
            "issue"
        ] = match.group(
            2
        )

    # --------------------------------------------------------
    # PAGE VOLUME/ISSUE FALLBACK
    # --------------------------------------------------------

    if (
        not result[
            "volume"
        ]
        or
        not result[
            "issue"
        ]
    ):

        match = re.search(
            r"Volume\s+(\d+)\s*,?\s*Issue\s+(\d+)",
            body_text,
            re.I
        )

        if match:

            result[
                "volume"
            ] = match.group(
                1
            )

            result[
                "issue"
            ] = match.group(
                2
            )

    return result


# ============================================================
# ARTICLE URLS
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

            href = (
                element.get_attribute(
                    "href"
                )
                or ""
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
    ]

    for selector in selectors:

        try:

            elements = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    selector
                )
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

            href = (
                link.get_attribute(
                    "href"
                )
                or ""
            )

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

        elements = driver.find_elements(
            By.CSS_SELECTOR,
            ".publication-title"
        )

        if elements:

            result[
                "journal"
            ] = clean_text(
                elements[
                    0
                ].text
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # PUBLICATION METADATA
    # --------------------------------------------------------

    try:

        metadata_text = (
            driver.find_element(
                By.CSS_SELECTOR,
                ".publication-metadata"
            ).text
        )

    except Exception:

        metadata_text = ""

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    match = re.search(
        r"Date:\s*([A-Za-z]+\s+20\d{2})",
        metadata_text,
        re.I
    )

    if match:

        result[
            "date"
        ] = clean_text(
            match.group(
                1
            )
        )

    result[
        "year"
    ] = (
        extract_year(
            result[
                "date"
            ]
        )
        or ""
    )

    # --------------------------------------------------------
    # VOLUME / ISSUE
    # --------------------------------------------------------

    match = re.search(
        r"Volume\s*:?\s*"
        r"(?:Volume\s*)?"
        r"(\d+)"
        r"\s*,?\s*Issue\s+(\d+)",
        metadata_text,
        re.I
    )

    if match:

        result[
            "volume"
        ] = match.group(
            1
        )

        result[
            "issue"
        ] = match.group(
            2
        )

    # --------------------------------------------------------
    # DOI
    # --------------------------------------------------------

    result[
        "doi"
    ] = extract_doi(
        driver
    )

    return result


# ============================================================
# CLEAN AUTHOR NAME
# ============================================================

def clean_author_name(
    name
):

    if not name:
        return ""

    name = clean_text(
        name
    )

    # Remove common qualifications from the END only.

    degree_pattern = re.compile(
        r"""
        (?:
            \s*,?\s*
            (?:
                MD|
                DM|
                RN|
                PhD|
                Ph\.D\.?|
                MSc|
                MS|
                MBBS|
                MPH|
                DNP|
                FRCA|
                FRCPC|
                BSN|
                MSN|
                MBA|
                APRN|
                CCRP|
                CPAN|
                CPE|
                CNS|
                FNP-BC|
                NE-BC|
                FAAN|
                FASPAN|
                FAIHA
            )
        )+
        \s*$
        """,
        re.I | re.X
    )

    previous = None

    while (
        previous
        != name
    ):

        previous = name

        name = degree_pattern.sub(
            "",
            name
        ).strip(
            " ,"
        )

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

    given = ""

    surname = ""

    try:

        elements = button.find_elements(
            By.CSS_SELECTOR,
            ".given-name"
        )

        if elements:

            given = clean_text(
                elements[
                    0
                ].text
            )

    except Exception:
        pass

    try:

        elements = button.find_elements(
            By.CSS_SELECTOR,
            ".surname"
        )

        if elements:

            surname = clean_text(
                elements[
                    0
                ].text
            )

    except Exception:
        pass

    if (
        given
        or surname
    ):

        return clean_author_name(
            f"{given} {surname}"
        )

    try:

        return clean_author_name(
            button.text
        )

    except Exception:

        return ""


# ============================================================
# EMAIL / CONTACT ICON
# ============================================================

def author_has_email_icon(
    button
):

    # Exact current ScienceDirect icon.

    try:

        icons = button.find_elements(
            By.CSS_SELECTOR,
            "svg.icon-envelope"
        )

        if icons:

            return True

    except Exception:
        pass

    try:

        html = (
            button.get_attribute(
                "outerHTML"
            )
            or ""
        ).lower()

        if (
            "icon-envelope"
            in html
        ):

            return True

        if (
            "author email or social media contact details icon"
            in html
        ):

            return True

    except Exception:
        pass

    return False


# ============================================================
# AUTHOR PANEL
# ============================================================

def find_author_panel(
    driver
):

    try:

        panels = driver.find_elements(
            By.CSS_SELECTOR,
            (
                "#side-panel."
                "side-panel-is-expanded"
            )
        )

        for panel in panels:

            if panel.is_displayed():

                return panel

    except Exception:
        pass

    return None


def wait_for_author_panel(
    driver
):

    try:

        panel = WebDriverWait(
            driver,
            AUTHOR_PANEL_OPEN_TIMEOUT
        ).until(
            EC.visibility_of_element_located(
                (
                    By.CSS_SELECTOR,
                    (
                        "#side-panel."
                        "side-panel-is-expanded"
                    )
                )
            )
        )

        return panel

    except TimeoutException:

        return None


# ============================================================
# WAIT FOR EMAIL
# ============================================================

def wait_for_author_email(
    driver
):

    end_time = (
        time.time()
        + AUTHOR_EMAIL_TIMEOUT
    )

    while (
        time.time()
        < end_time
    ):

        ensure_driver_alive(
            driver
        )

        # ----------------------------------------------------
        # EXACT SCIENCEDIRECT EMAIL
        # ----------------------------------------------------

        try:

            links = driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel."
                    "side-panel-is-expanded "
                    ".e-address "
                    'a[href^="mailto:"]'
                )
            )

            for link in links:

                try:

                    if not link.is_displayed():
                        continue

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

                    continue

        except Exception:
            pass

        # ----------------------------------------------------
        # TEXT FALLBACK
        # ----------------------------------------------------

        try:

            elements = driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel."
                    "side-panel-is-expanded "
                    ".e-address"
                )
            )

            for element in elements:

                matches = EMAIL_RE.findall(
                    element.text
                    or ""
                )

                if matches:

                    return (
                        matches[
                            0
                        ]
                        .strip()
                        .lower()
                    )

        except Exception:
            pass

        time.sleep(
            AUTHOR_EMAIL_POLL_INTERVAL
        )

    return ""


# ============================================================
# PANEL AUTHOR NAME
# ============================================================

def extract_panel_author_name(
    driver
):

    try:

        element = driver.find_element(
            By.CSS_SELECTOR,
            (
                "#side-panel."
                "side-panel-is-expanded "
                ".side-panel-children "
                ".side-panel-author > h2"
            )
        )

        return clean_author_name(
            clean_text(
                element.text
            )
        )

    except Exception:

        return ""


# ============================================================
# AFFILIATION
# ============================================================

def extract_panel_affiliation(
    driver
):

    selectors = [
        (
            "#side-panel."
            "side-panel-is-expanded "
            ".side-panel-children > "
            ".side-panel-author > "
            ".affiliation"
        ),
        (
            "#side-panel."
            "side-panel-is-expanded "
            ".affiliation"
        ),
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
# CORRESPONDENCE ADDRESS
# ============================================================

def extract_correspondence_address(
    driver
):

    try:

        elements = driver.find_elements(
            By.CSS_SELECTOR,
            (
                "#side-panel."
                "side-panel-is-expanded "
                ".correspondence span"
            )
        )

        for element in elements:

            text = clean_text(
                element.text
            )

            if not text:
                continue

            text = re.sub(
                r"^Address correspondence to\s*:\s*",
                "",
                text,
                flags=re.I
            )

            return clean_text(
                text
            )

    except Exception:
        pass

    return ""


# ============================================================
# CLOSE AUTHOR PANEL
# ============================================================

def close_author_panel(
    driver
):

    if not driver_is_alive(
        driver
    ):

        return False

    try:

        buttons = driver.find_elements(
            By.CSS_SELECTOR,
            (
                "#side-panel "
                "button.side-panel-close-btn"
            )
        )

        for button in buttons:

            try:

                if not button.is_displayed():
                    continue

                driver.execute_script(
                    "arguments[0].click();",
                    button
                )

                try:

                    WebDriverWait(
                        driver,
                        5
                    ).until(
                        EC.invisibility_of_element_located(
                            (
                                By.CSS_SELECTOR,
                                (
                                    "#side-panel."
                                    "side-panel-is-expanded"
                                )
                            )
                        )
                    )

                except Exception:

                    pass

                time.sleep(
                    0.5
                )

                return True

            except Exception:

                continue

    except Exception:
        pass

    # --------------------------------------------------------
    # ESC FALLBACK
    # --------------------------------------------------------

    try:

        driver.find_element(
            By.TAG_NAME,
            "body"
        ).send_keys(
            Keys.ESCAPE
        )

        time.sleep(
            0.5
        )

    except Exception:
        pass

    return False


# ============================================================
# OPEN AUTHOR PANEL
# ============================================================

def open_author_panel(
    driver,
    author_index
):

    ensure_driver_alive(
        driver
    )

    buttons = get_author_buttons(
        driver
    )

    if (
        author_index
        >= len(
            buttons
        )
    ):

        return None

    button = buttons[
        author_index
    ]

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        button
    )

    time.sleep(
        0.5
    )

    # JS click works better for ScienceDirect's React control.

    driver.execute_script(
        "arguments[0].click();",
        button
    )

    return wait_for_author_panel(
        driver
    )


# ============================================================
# EXTRACT ARTICLE AUTHORS
# ============================================================

def extract_article_authors(
    driver
):

    results = []

    buttons = get_author_buttons(
        driver
    )

    total_authors = len(
        buttons
    )

    print(
        f"    Author buttons found: "
        f"{total_authors}"
    )

    for index in range(
        total_authors
    ):

        ensure_driver_alive(
            driver
        )

        buttons = get_author_buttons(
            driver
        )

        if (
            index
            >= len(
                buttons
            )
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

        has_contact_icon = (
            author_has_email_icon(
                button
            )
        )

        print()
        print(
            f"      [{index + 1}/"
            f"{total_authors}] "
            f"{author_name} "
            f"({author_id})"
        )

        print(
            f"          Contact icon: "
            f"{'YES' if has_contact_icon else 'NO'}"
        )

        email = ""

        affiliation = ""

        correspondence_address = ""

        # ====================================================
        # NO CONTACT ICON
        # ====================================================

        if not has_contact_icon:

            print(
                "          EMAIL: -"
            )

            results.append(
                {
                    "author":
                        author_name,

                    "email":
                        "",

                    "affiliation":
                        "",

                    "correspondence_address":
                        "",

                    "has_contact_icon":
                        "No",

                    "author_id":
                        author_id,
                }
            )

            continue

        # ====================================================
        # CONTACT AUTHOR
        # ====================================================

        try:

            print(
                "          Clicking author..."
            )

            panel = open_author_panel(
                driver,
                index
            )

            # ------------------------------------------------
            # RETRY CLICK ONCE
            # ------------------------------------------------

            if panel is None:

                print(
                    "          First click did not open panel."
                )

                close_author_panel(
                    driver
                )

                time.sleep(
                    1
                )

                print(
                    "          Retrying author click..."
                )

                panel = open_author_panel(
                    driver,
                    index
                )

            if panel is None:

                print(
                    "          PANEL NOT FOUND"
                )

            else:

                print(
                    "          Panel opened."
                )

                # --------------------------------------------
                # PANEL AUTHOR
                # --------------------------------------------

                panel_author = (
                    extract_panel_author_name(
                        driver
                    )
                )

                if panel_author:

                    author_name = (
                        panel_author
                    )

                    print(
                        f"          Panel author: "
                        f"{panel_author}"
                    )

                # --------------------------------------------
                # EMAIL
                # --------------------------------------------

                print(
                    "          Waiting for email..."
                )

                email = wait_for_author_email(
                    driver
                )

                if email:

                    print(
                        f"          EMAIL FOUND: "
                        f"{email}"
                    )

                else:

                    print(
                        "          EMAIL NOT FOUND"
                    )

                # --------------------------------------------
                # AFFILIATION
                # --------------------------------------------

                affiliation = (
                    extract_panel_affiliation(
                        driver
                    )
                )

                if affiliation:

                    print(
                        f"          Affiliation: "
                        f"{affiliation[:120]}"
                    )

                # --------------------------------------------
                # CORRESPONDENCE
                # --------------------------------------------

                correspondence_address = (
                    extract_correspondence_address(
                        driver
                    )
                )

        except (
            NoSuchWindowException,
            InvalidSessionIdException,
        ) as exc:

            raise BrowserSessionLost(
                str(exc)
            )

        except WebDriverException as exc:

            if not driver_is_alive(
                driver
            ):

                raise BrowserSessionLost(
                    str(exc)
                )

            print(
                f"          AUTHOR ERROR: "
                f"{exc}"
            )

        except Exception as exc:

            print(
                f"          AUTHOR ERROR: "
                f"{exc}"
            )

        finally:

            if driver_is_alive(
                driver
            ):

                close_author_panel(
                    driver
                )

                time.sleep(
                    0.8
                )

        results.append(
            {
                "author":
                    author_name,

                "email":
                    email,

                "affiliation":
                    affiliation,

                "correspondence_address":
                    correspondence_address,

                "has_contact_icon":
                    "Yes",

                "author_id":
                    author_id,
            }
        )

    return results


# ============================================================
# ARTICLE PARSER
# ============================================================

def parse_article(
    driver,
    article_url,
    journal,
    issue_metadata,
    records,
    completed_articles
):

    print()
    print(
        "-" * 100
    )

    print(
        f"ARTICLE: {article_url}"
    )

    # ========================================================
    # OPEN
    # ========================================================

    open_page(
        driver,
        article_url,
        records,
        completed_articles,
        PAGE_LOAD_WAIT
    )

    ensure_driver_alive(
        driver
    )

    # ========================================================
    # WAIT FOR ARTICLE STRUCTURE
    # ========================================================

    try:

        WebDriverWait(
            driver,
            10
        ).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    "h1"
                )
            )
        )

    except TimeoutException:

        pass

    # ========================================================
    # METADATA
    # ========================================================

    title = extract_article_title(
        driver
    )

    metadata = extract_article_metadata(
        driver
    )

    # --------------------------------------------------------
    # ISSUE FALLBACKS
    # --------------------------------------------------------

    if not metadata[
        "journal"
    ]:

        metadata[
            "journal"
        ] = journal[
            "name"
        ]

    if not metadata[
        "year"
    ]:

        metadata[
            "year"
        ] = issue_metadata.get(
            "year",
            ""
        )

    if not metadata[
        "date"
    ]:

        metadata[
            "date"
        ] = issue_metadata.get(
            "date",
            ""
        )

    if not metadata[
        "volume"
    ]:

        metadata[
            "volume"
        ] = issue_metadata.get(
            "volume",
            ""
        )

    if not metadata[
        "issue"
    ]:

        metadata[
            "issue"
        ] = issue_metadata.get(
            "issue",
            ""
        )

    print(
        f"    Title  : "
        f"{title}"
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

    # ========================================================
    # AUTHORS
    # ========================================================

    authors = extract_article_authors(
        driver
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

                "correspondence_address":
                    author[
                        "correspondence_address"
                    ],

                "has_contact_icon":
                    author[
                        "has_contact_icon"
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

    # A page with 0 authors can still be legitimate:
    # Editorial Board, TOC, advertisements etc.

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
        normalize_url(
            url
        )
        for url in data.get(
            "completed_articles",
            []
        )
        if url
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
# REMOVE ARTICLE RECORDS
# ============================================================

def remove_article_records(
    records,
    article_url
):

    article_url = normalize_url(
        article_url
    )

    return [
        record
        for record in records
        if normalize_url(
            record.get(
                "article_url",
                ""
            )
        )
        != article_url
    ]


# ============================================================
# SHOULD REPROCESS?
# ============================================================

def article_needs_reprocess(
    article_url,
    records
):

    article_url = normalize_url(
        article_url
    )

    rows = [
        record
        for record in records
        if normalize_url(
            record.get(
                "article_url",
                ""
            )
        )
        == article_url
    ]

    # No author rows can be a legitimate non-author item.
    if not rows:
        return False

    for row in rows:

        contact = (
            str(
                row.get(
                    "has_contact_icon",
                    ""
                )
            )
            .strip()
            .lower()
        )

        email = (
            str(
                row.get(
                    "email",
                    ""
                )
            )
            .strip()
        )

        if (
            contact == "yes"
            and
            not email
        ):

            return True

    return False


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(
    records
):

    if not records:

        print(
            "No author records available for Excel."
        )

        return

    df = pd.DataFrame(
        records
    )

    columns = [
        "journal",
        "year",
        "date",
        "volume",
        "issue",
        "article_title",
        "author",
        "email",
        "affiliation",
        "correspondence_address",
        "has_contact_icon",
        "author_id",
        "doi",
        "article_url",
    ]

    for column in columns:

        if column not in df.columns:

            df[
                column
            ] = ""

    df = df[
        columns
    ]

    # --------------------------------------------------------
    # CLEAN
    # --------------------------------------------------------

    df[
        "email"
    ] = (
        df[
            "email"
        ]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df[
        "author"
    ] = (
        df[
            "author"
        ]
        .fillna("")
        .astype(str)
        .map(
            clean_author_name
        )
    )

    # --------------------------------------------------------
    # DEDUPE
    # --------------------------------------------------------

    df.drop_duplicates(
        subset=[
            "article_url",
            "author",
            "email",
        ],
        keep="last",
        inplace=True
    )

    # --------------------------------------------------------
    # EMAIL ONLY
    # --------------------------------------------------------

    email_df = (
        df[
            df[
                "email"
            ].ne("")
        ]
        .copy()
    )

    email_df.drop_duplicates(
        subset=[
            "email"
        ],
        keep="first",
        inplace=True
    )

    # --------------------------------------------------------
    # SORT
    # --------------------------------------------------------

    df.sort_values(
        by=[
            "journal",
            "year",
            "volume",
            "issue",
            "article_title",
            "author",
        ],
        inplace=True,
        na_position="last"
    )

    email_df.sort_values(
        by=[
            "journal",
            "year",
            "author",
        ],
        inplace=True,
        na_position="last"
    )

    # --------------------------------------------------------
    # WRITE
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

        workbook = writer.book

        widths = {
            "A": 45,
            "B": 10,
            "C": 20,
            "D": 10,
            "E": 10,
            "F": 75,
            "G": 38,
            "H": 45,
            "I": 85,
            "J": 100,
            "K": 20,
            "L": 15,
            "M": 40,
            "N": 90,
        }

        for sheet_name in [
            "All Authors",
            "Authors With Email",
        ]:

            worksheet = workbook[
                sheet_name
            ]

            worksheet.freeze_panes = (
                "A2"
            )

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            for (
                column,
                width
            ) in widths.items():

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

    print(
        "=" * 110
    )

    print(
        "SCIENCEDIRECT AUTHOR + EMAIL SCRAPER"
    )

    print(
        "=" * 110
    )

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

    driver = None

    successful_this_session = 0

    stop_reason = ""

    try:

        driver = connect_to_chrome()

        # ====================================================
        # JOURNALS
        # ====================================================

        for (
            journal_index,
            journal
        ) in enumerate(
            JOURNALS,
            start=1
        ):

            print()
            print(
                "=" * 110
            )

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

            print(
                "=" * 110
            )

            # ------------------------------------------------
            # ISSUE LIST
            # ------------------------------------------------

            issue_links = (
                extract_issue_links(
                    driver,
                    journal,
                    records,
                    completed_articles
                )
            )

            print(
                f"Issue links collected: "
                f"{len(issue_links)}"
            )

            # =================================================
            # ISSUES
            # =================================================

            for (
                issue_index,
                issue_url
            ) in enumerate(
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

                open_page(
                    driver,
                    issue_url,
                    records,
                    completed_articles,
                    ISSUE_LOAD_WAIT
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
                    f"Issue year   : "
                    f"{issue_year}"
                )

                print(
                    f"Issue volume : "
                    f"{issue_metadata.get('volume')}"
                )

                print(
                    f"Issue number : "
                    f"{issue_metadata.get('issue')}"
                )

                # --------------------------------------------
                # YEAR FILTER
                # --------------------------------------------

                if issue_year:

                    try:

                        issue_year_int = int(
                            issue_year
                        )

                    except Exception:

                        issue_year_int = None

                    if (
                        issue_year_int
                        and
                        not (
                            journal[
                                "start_year"
                            ]
                            <= issue_year_int
                            <= journal[
                                "end_year"
                            ]
                        )
                    ):

                        print(
                            "Skipping issue outside requested years."
                        )

                        continue

                # --------------------------------------------
                # ARTICLE URLS
                # --------------------------------------------

                article_urls = (
                    extract_article_urls(
                        driver
                    )
                )

                print(
                    f"Articles found: "
                    f"{len(article_urls)}"
                )

                # ============================================
                # ARTICLES
                # ============================================

                for (
                    article_index,
                    article_url
                ) in enumerate(
                    article_urls,
                    start=1
                ):

                    article_url = (
                        normalize_url(
                            article_url
                        )
                    )

                    already_done = (
                        article_url
                        in completed_articles
                    )

                    needs_reprocess = False

                    if already_done:

                        needs_reprocess = (
                            article_needs_reprocess(
                                article_url,
                                records
                            )
                        )

                    if (
                        already_done
                        and
                        not needs_reprocess
                    ):

                        print(
                            f"[Article "
                            f"{article_index}/"
                            f"{len(article_urls)}] "
                            f"Already processed"
                        )

                        continue

                    if needs_reprocess:

                        print()
                        print(
                            f"[Article "
                            f"{article_index}/"
                            f"{len(article_urls)}] "
                            f"REPROCESSING - missing contact email"
                        )

                    else:

                        print()
                        print(
                            f"[Article "
                            f"{article_index}/"
                            f"{len(article_urls)}]"
                        )

                    # ----------------------------------------
                    # IMPORTANT:
                    # DO NOT DELETE OLD DATA UNTIL NEW
                    # PROCESSING HAS ACTUALLY SUCCEEDED.
                    # ----------------------------------------

                    try:

                        article_records = (
                            parse_article(
                                driver,
                                article_url,
                                journal,
                                issue_metadata,
                                records,
                                completed_articles
                            )
                        )

                    except ScienceDirectBlocked as exc:

                        stop_reason = str(
                            exc
                        )

                        print()
                        print(
                            "ScienceDirect blocked the URL again."
                        )

                        print(
                            "Stopping safely instead of repeatedly retrying."
                        )

                        raise

                    except BrowserSessionLost as exc:

                        stop_reason = str(
                            exc
                        )

                        print()
                        print(
                            "Browser session lost."
                        )

                        print(
                            "Stopping safely. Restart Chrome and rerun."
                        )

                        raise

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
                                    str(
                                        exc
                                    ),

                                "timestamp":
                                    datetime.now()
                                    .isoformat(),
                            }
                        )

                        save_json(
                            ERROR_FILE,
                            errors
                        )

                        # CRITICAL:
                        # failed article is NOT completed.

                        completed_articles.discard(
                            article_url
                        )

                        save_checkpoint(
                            records,
                            completed_articles
                        )

                        continue

                    # ========================================
                    # ARTICLE SUCCESSFUL
                    # ========================================

                    # Replace previous records only AFTER
                    # successful parsing.

                    records = (
                        remove_article_records(
                            records,
                            article_url
                        )
                    )

                    # ----------------------------------------
                    # YEAR FILTER
                    # ----------------------------------------

                    accepted = []

                    for record in article_records:

                        record_year = (
                            record.get(
                                "year"
                            )
                        )

                        if record_year:

                            try:

                                year_int = int(
                                    record_year
                                )

                            except Exception:

                                year_int = None

                            if (
                                year_int
                                and
                                not (
                                    journal[
                                        "start_year"
                                    ]
                                    <= year_int
                                    <= journal[
                                        "end_year"
                                    ]
                                )
                            ):

                                continue

                        accepted.append(
                            record
                        )

                    records.extend(
                        accepted
                    )

                    # ----------------------------------------
                    # ONLY NOW MARK COMPLETED
                    # ----------------------------------------

                    completed_articles.add(
                        article_url
                    )

                    successful_this_session += 1

                    print(
                        "    Article completed successfully."
                    )

                    # ----------------------------------------
                    # CHECKPOINT AFTER EVERY ARTICLE
                    # ----------------------------------------

                    save_checkpoint(
                        records,
                        completed_articles
                    )

                    # ----------------------------------------
                    # SAVE EXCEL
                    # ----------------------------------------

                    if (
                        successful_this_session
                        % SAVE_EVERY_N_ARTICLES
                        == 0
                    ):

                        save_excel(
                            records
                        )

                    # ----------------------------------------
                    # PERIODIC LONGER COOLDOWN
                    # ----------------------------------------

                    if (
                        successful_this_session
                        % COOLDOWN_AFTER_ARTICLES
                        == 0
                    ):

                        print()
                        print(
                            f"Processed "
                            f"{successful_this_session} "
                            f"articles this session."
                        )

                        print(
                            f"Cooling down for "
                            f"{COOLDOWN_SECONDS} seconds..."
                        )

                        save_checkpoint(
                            records,
                            completed_articles
                        )

                        save_excel(
                            records
                        )

                        time.sleep(
                            COOLDOWN_SECONDS
                        )

                    else:

                        polite_delay()

    except KeyboardInterrupt:

        stop_reason = (
            "Stopped manually with Ctrl+C."
        )

        print()
        print(
            stop_reason
        )

    except ScienceDirectBlocked:

        pass

    except BrowserSessionLost as exc:

        stop_reason = (
            f"Browser session lost: {exc}"
        )

    except Exception as exc:

        stop_reason = (
            f"Fatal error: {exc}"
        )

        print()
        print(
            stop_reason
        )

        errors.append(
            {
                "type":
                    "fatal",

                "error":
                    str(
                        exc
                    ),

                "timestamp":
                    datetime.now()
                    .isoformat(),
            }
        )

    finally:

        print()
        print(
            "Saving final checkpoint..."
        )

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
        print(
            "=" * 110
        )

        print(
            "FINISHED / PAUSED SAFELY"
        )

        print(
            "=" * 110
        )

        print(
            f"Completed articles : "
            f"{len(completed_articles):,}"
        )

        print(
            f"Author records     : "
            f"{len(records):,}"
        )

        email_count = sum(
            1
            for record in records
            if str(
                record.get(
                    "email",
                    ""
                )
            ).strip()
        )

        print(
            f"Emails captured    : "
            f"{email_count:,}"
        )

        print(
            f"Errors             : "
            f"{len(errors):,}"
        )

        if stop_reason:

            print(
                f"Stop reason        : "
                f"{stop_reason}"
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