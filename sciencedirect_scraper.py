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


# ============================================================
# CHROME
# ============================================================

REMOTE_DEBUGGING_ADDRESS = "127.0.0.1:9222"


# ============================================================
# NORMAL REQUEST PACING
# ============================================================

PAGE_LOAD_WAIT_MIN = 4
PAGE_LOAD_WAIT_MAX = 7

ISSUE_LOAD_WAIT_MIN = 4
ISSUE_LOAD_WAIT_MAX = 7

ARTICLE_DELAY_MIN = 8
ARTICLE_DELAY_MAX = 15

AUTHOR_CLICK_DELAY_MIN = 1.0
AUTHOR_CLICK_DELAY_MAX = 2.5

AUTHOR_PANEL_OPEN_TIMEOUT = 15
AUTHOR_EMAIL_TIMEOUT = 25
AUTHOR_EMAIL_POLL_INTERVAL = 0.5


# ============================================================
# PERIODIC COOLDOWN
# ============================================================

COOLDOWN_AFTER_ARTICLES = 10
COOLDOWN_SECONDS_MIN = 45
COOLDOWN_SECONDS_MAX = 90


# ============================================================
# BLOCK RETRY
# ============================================================

# When CPE00001 / CPE00002 occurs:
#
# retry 1 -> wait 1 minute
# retry 2 -> wait 3 minutes
# retry 3 -> wait 5 minutes
# retry 4 -> wait 10 minutes
# retry 5 -> wait 15 minutes

BLOCK_BACKOFF_STEPS = [
    60,
    180,
    300,
    600,
    900,
]

MAX_BLOCK_RETRIES = len(
    BLOCK_BACKOFF_STEPS
)


# ============================================================
# SAVE FREQUENCY
# ============================================================

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
# EXCEPTIONS
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


def random_wait(
    minimum,
    maximum,
    message=None
):

    seconds = random.uniform(
        minimum,
        maximum
    )

    if message:

        print(
            f"{message}: "
            f"{seconds:.1f}s"
        )

    time.sleep(
        seconds
    )


def save_json(
    path,
    data
):

    with open(
        path,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            data,
            file,
            ensure_ascii=False,
            indent=2
        )


def load_json(path):

    if not os.path.exists(
        path
    ):

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
            "Chrome is running but no usable tab/window exists."
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

        handles = (
            driver.window_handles
        )

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


def ensure_driver_alive(
    driver
):

    if not driver_is_alive(
        driver
    ):

        raise BrowserSessionLost(
            "Chrome/Selenium window is no longer available."
        )


# ============================================================
# TRY RECONNECT
# ============================================================

def reconnect_to_chrome(
    old_driver=None
):

    print()
    print(
        "=" * 100
    )

    print(
        "ATTEMPTING TO RECONNECT TO CHROME"
    )

    print(
        "=" * 100
    )

    try:

        if old_driver:

            old_driver.quit()

    except Exception:

        pass

    for attempt in range(
        1,
        6
    ):

        try:

            print(
                f"Reconnect attempt "
                f"{attempt}/5..."
            )

            driver = (
                connect_to_chrome()
            )

            if driver_is_alive(
                driver
            ):

                print(
                    "Chrome reconnection successful."
                )

                return driver

        except Exception as exc:

            print(
                f"Reconnect failed: "
                f"{exc}"
            )

        time.sleep(
            5
        )

    raise BrowserSessionLost(
        "Unable to reconnect to Chrome."
    )


# ============================================================
# BLOCK DETECTION
# ============================================================

def is_sciencedirect_blocked(
    driver
):

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
        "cpe00001",
        "cpe00002",
        "reference number:",
        "unusual traffic",
        "access denied",
    ]

    for indicator in indicators:

        if (
            indicator in html
            or
            indicator in title
        ):

            return True

    return False


# ============================================================
# CHECK IF REAL ARTICLE LOADED
# ============================================================

def is_article_page_loaded(
    driver
):

    try:

        if is_sciencedirect_blocked(
            driver
        ):

            return False

        title_elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "h1.content-title, "
                    "#screen-reader-main-title"
                )
            )
        )

        publication = (
            driver.find_elements(
                By.CSS_SELECTOR,
                ".publication-title"
            )
        )

        return bool(
            title_elements
            and publication
        )

    except Exception:

        return False


# ============================================================
# AUTOMATIC BLOCK WAIT / RETRY
# ============================================================

def wait_and_retry_blocked_url(
    driver,
    url,
    records,
    completed_articles
):

    print()
    print(
        "=" * 100
    )

    print(
        "SCIENCEDIRECT BLOCK DETECTED"
    )

    print(
        "=" * 100
    )

    print(
        "Saving current progress..."
    )

    save_checkpoint(
        records,
        completed_articles
    )

    save_excel(
        records
    )

    for (
        retry_number,
        delay_seconds
    ) in enumerate(
        BLOCK_BACKOFF_STEPS,
        start=1
    ):

        print()
        print(
            "-" * 100
        )

        print(
            f"Block retry "
            f"{retry_number}/"
            f"{MAX_BLOCK_RETRIES}"
        )

        print(
            f"Waiting "
            f"{delay_seconds} seconds..."
        )

        print(
            "-" * 100
        )

        time.sleep(
            delay_seconds
        )

        # ----------------------------------------------------
        # CHECK BROWSER
        # ----------------------------------------------------

        if not driver_is_alive(
            driver
        ):

            try:

                driver = (
                    reconnect_to_chrome(
                        driver
                    )
                )

            except Exception:

                raise BrowserSessionLost(
                    "Chrome disappeared during block cooldown."
                )

        # ----------------------------------------------------
        # RETRY SAME URL ONCE
        # ----------------------------------------------------

        try:

            print(
                f"Retrying: {url}"
            )

            driver.get(
                url
            )

            random_wait(
                PAGE_LOAD_WAIT_MIN,
                PAGE_LOAD_WAIT_MAX,
                "Waiting after retry"
            )

        except (
            NoSuchWindowException,
            InvalidSessionIdException,
        ):

            driver = (
                reconnect_to_chrome(
                    driver
                )
            )

            continue

        except WebDriverException:

            if not driver_is_alive(
                driver
            ):

                driver = (
                    reconnect_to_chrome(
                        driver
                    )
                )

            continue

        # ----------------------------------------------------
        # ACCESS RESTORED
        # ----------------------------------------------------

        if not is_sciencedirect_blocked(
            driver
        ):

            print()
            print(
                "ScienceDirect access restored."
            )

            return (
                driver,
                True
            )

        print(
            "Still blocked."
        )

        save_checkpoint(
            records,
            completed_articles
        )

    # ========================================================
    # FAILED ALL RETRIES
    # ========================================================

    print()
    print(
        "=" * 100
    )

    print(
        "SCIENCEDIRECT REMAINED BLOCKED"
    )

    print(
        "=" * 100
    )

    print(
        "Progress has been saved."
    )

    print(
        "Current URL remains unfinished."
    )

    return (
        driver,
        False
    )


# ============================================================
# OPEN URL
# ============================================================

def open_page(
    driver,
    url,
    records,
    completed_articles,
    minimum_wait=PAGE_LOAD_WAIT_MIN,
    maximum_wait=PAGE_LOAD_WAIT_MAX,
):

    if not driver_is_alive(
        driver
    ):

        driver = reconnect_to_chrome(
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
    ):

        driver = reconnect_to_chrome(
            driver
        )

        driver.get(
            url
        )

    except WebDriverException as exc:

        if not driver_is_alive(
            driver
        ):

            driver = (
                reconnect_to_chrome(
                    driver
                )
            )

            driver.get(
                url
            )

        else:

            raise exc

    random_wait(
        minimum_wait,
        maximum_wait,
        "Page load wait"
    )

    ensure_driver_alive(
        driver
    )

    # ========================================================
    # BLOCK HANDLING
    # ========================================================

    if is_sciencedirect_blocked(
        driver
    ):

        (
            driver,
            recovered
        ) = wait_and_retry_blocked_url(
            driver,
            url,
            records,
            completed_articles
        )

        if not recovered:

            raise ScienceDirectBlocked(
                "ScienceDirect remained blocked "
                "after all cooldown retries."
            )

    return driver


# ============================================================
# SCROLL
# ============================================================

def scroll_page(
    driver,
    max_scrolls=12
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
                random.uniform(
                    0.8,
                    1.5
                )
            )

            previous_height = (
                current_height
            )

        except Exception:

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
        journal["url"]
    )

    driver = open_page(
        driver,
        journal["url"],
        records,
        completed_articles
    )

    scroll_page(
        driver
    )

    links = []

    seen = set()

    anchors = (
        driver.find_elements(
            By.CSS_SELECTOR,
            'a[href*="/vol/"][href*="/issue/"]'
        )
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

            seen.add(
                href
            )

            links.append(
                href
            )

        except Exception:

            continue

    return (
        driver,
        links
    )


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

        elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                ".publication-title"
            )
        )

        if elements:

            result[
                "journal"
            ] = clean_text(
                elements[0].text
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

    try:

        current_url = (
            driver.current_url
            or ""
        )

    except Exception:

        current_url = ""

    # --------------------------------------------------------
    # URL VOLUME / ISSUE
    # --------------------------------------------------------

    match = re.search(
        r"/vol/(\d+)/issue/(\d+)",
        current_url,
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
    # DATE
    # --------------------------------------------------------

    match = re.search(
        r"(January|February|March|April|May|June|"
        r"July|August|September|October|November|December)"
        r"\s+(20\d{2})",
        body_text,
        re.I
    )

    if match:

        result[
            "date"
        ] = clean_text(
            match.group(0)
        )

        result[
            "year"
        ] = int(
            match.group(2)
        )

    else:

        result[
            "year"
        ] = (
            extract_year(
                body_text
            )
            or ""
        )

    return result


# ============================================================
# ARTICLE URLS FROM ISSUE
# ============================================================

def extract_article_urls(
    driver
):

    urls = []

    seen = set()

    elements = (
        driver.find_elements(
            By.CSS_SELECTOR,
            (
                'a.article-content-title'
                '[href*="/science/article/pii/"]'
            )
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
        (
            "h1.content-title "
            ".title-text"
        ),
        (
            "#screen-reader-main-title "
            ".title-text"
        ),
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

        links = (
            driver.find_elements(
                By.CSS_SELECTOR,
                'a[href*="doi.org/10."]'
            )
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

    try:

        elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                ".publication-title"
            )
        )

        if elements:

            result[
                "journal"
            ] = clean_text(
                elements[0].text
            )

    except Exception:

        pass

    metadata_text = ""

    try:

        metadata_text = (
            driver.find_element(
                By.CSS_SELECTOR,
                ".publication-metadata"
            ).text
        )

    except Exception:

        pass

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    date_match = re.search(
        r"Date:\s*"
        r"([A-Za-z]+\s+20\d{2})",
        metadata_text,
        re.I
    )

    if date_match:

        result[
            "date"
        ] = clean_text(
            date_match.group(
                1
            )
        )

    result[
        "year"
    ] = (
        extract_year(
            result["date"]
        )
        or ""
    )

    # --------------------------------------------------------
    # VOLUME / ISSUE
    # --------------------------------------------------------

    volume_issue = re.search(
        r"Volume\s*:?\s*"
        r"(?:Volume\s*)?"
        r"(\d+)"
        r"\s*,?\s*Issue\s+(\d+)",
        metadata_text,
        re.I
    )

    if volume_issue:

        result[
            "volume"
        ] = volume_issue.group(
            1
        )

        result[
            "issue"
        ] = volume_issue.group(
            2
        )

    result[
        "doi"
    ] = extract_doi(
        driver
    )

    return result


# ============================================================
# AUTHOR NAME CLEANUP
# ============================================================

def clean_author_name(
    name
):

    if not name:

        return ""

    name = clean_text(
        name
    )

    qualifications = [
        "MD",
        "DM",
        "RN",
        "PhD",
        "Ph.D",
        "MSc",
        "MS",
        "MBBS",
        "MPH",
        "DNP",
        "BSN",
        "MSN",
        "MBA",
        "FRCA",
        "FRCPC",
        "APRN",
        "CPAN",
        "CNS",
        "CCRP",
        "CFRE",
        "CPE",
        "FAAN",
        "FASPAN",
        "FAIHA",
        "NE-BC",
        "FNP-BC",
    ]

    changed = True

    while changed:

        changed = False

        for qualification in (
            qualifications
        ):

            pattern = (
                r"[\s,]+"
                + re.escape(
                    qualification
                )
                + r"\.?$"
            )

            updated = re.sub(
                pattern,
                "",
                name,
                flags=re.I
            ).strip(
                " ,"
            )

            if updated != name:

                name = updated

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

        given = (
            button.find_elements(
                By.CSS_SELECTOR,
                ".given-name"
            )
        )

        if given:

            given_name = clean_text(
                given[0].text
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
# EMAIL ICON
# ============================================================

def author_has_email_icon(
    button
):

    try:

        icons = (
            button.find_elements(
                By.CSS_SELECTOR,
                "svg.icon-envelope"
            )
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

        panels = (
            driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel."
                    "side-panel-is-expanded"
                )
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

        return WebDriverWait(
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

    except TimeoutException:

        return None


# ============================================================
# EMAIL
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

        try:

            links = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    (
                        "#side-panel."
                        "side-panel-is-expanded "
                        ".e-address "
                        'a[href^="mailto:"]'
                    )
                )
            )

            for link in links:

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
                    .split("?")[0]
                    .strip()
                    .lower()
                )

                if EMAIL_RE.fullmatch(
                    email
                ):

                    return email

        except Exception:

            pass

        try:

            elements = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    (
                        "#side-panel."
                        "side-panel-is-expanded "
                        ".e-address"
                    )
                )
            )

            for element in elements:

                matches = (
                    EMAIL_RE.findall(
                        element.text
                        or ""
                    )
                )

                if matches:

                    return (
                        matches[0]
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
# PANEL AUTHOR
# ============================================================

def extract_panel_author_name(
    driver
):

    try:

        element = (
            driver.find_element(
                By.CSS_SELECTOR,
                (
                    "#side-panel."
                    "side-panel-is-expanded "
                    ".side-panel-children "
                    ".side-panel-author > h2"
                )
            )
        )

        return clean_author_name(
            element.text
        )

    except Exception:

        return ""


# ============================================================
# PANEL AFFILIATION
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
# CORRESPONDENCE ADDRESS
# ============================================================

def extract_correspondence_address(
    driver
):

    try:

        elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel."
                    "side-panel-is-expanded "
                    ".correspondence span"
                )
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

        buttons = (
            driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel "
                    "button.side-panel-close-btn"
                )
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
        >= len(buttons)
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

    random_wait(
        AUTHOR_CLICK_DELAY_MIN,
        AUTHOR_CLICK_DELAY_MAX,
        "          Author click wait"
    )

    driver.execute_script(
        "arguments[0].click();",
        button
    )

    return wait_for_author_panel(
        driver
    )


# ============================================================
# ARTICLE AUTHORS
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

        # ----------------------------------------------------
        # NO CONTACT ICON
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # CONTACT AUTHOR
        # ----------------------------------------------------

        try:

            print(
                "          Opening author panel..."
            )

            panel = (
                open_author_panel(
                    driver,
                    index
                )
            )

            # -----------------------------------------------
            # RETRY ONCE
            # -----------------------------------------------

            if panel is None:

                print(
                    "          First panel click failed."
                )

                close_author_panel(
                    driver
                )

                time.sleep(
                    1
                )

                panel = (
                    open_author_panel(
                        driver,
                        index
                    )
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
                # PANEL NAME
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
                        f"{author_name}"
                    )

                # --------------------------------------------
                # EMAIL
                # --------------------------------------------

                print(
                    "          Waiting for email..."
                )

                email = (
                    wait_for_author_email(
                        driver
                    )
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
                # ADDRESS
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
                    random.uniform(
                        0.8,
                        1.5
                    )
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
# PARSE ARTICLE
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

    driver = open_page(
        driver,
        article_url,
        records,
        completed_articles
    )

    ensure_driver_alive(
        driver
    )

    # --------------------------------------------------------
    # ARTICLE WAIT
    # --------------------------------------------------------

    try:

        WebDriverWait(
            driver,
            10
        ).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    (
                        "h1.content-title, "
                        "#screen-reader-main-title"
                    )
                )
            )
        )

    except TimeoutException:

        pass

    if is_sciencedirect_blocked(
        driver
    ):

        raise ScienceDirectBlocked(
            "Article is still blocked."
        )

    title = extract_article_title(
        driver
    )

    metadata = (
        extract_article_metadata(
            driver
        )
    )

    # --------------------------------------------------------
    # FALLBACKS
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

    return (
        driver,
        rows
    )


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
# REMOVE OLD ARTICLE DATA
# ============================================================

def remove_article_records(
    records,
    article_url
):

    article_url = normalize_url(
        article_url
    )

    return [
        row
        for row in records
        if normalize_url(
            row.get(
                "article_url",
                ""
            )
        )
        != article_url
    ]


# ============================================================
# REPROCESS CHECK
# ============================================================

def article_needs_reprocess(
    article_url,
    records
):

    article_url = normalize_url(
        article_url
    )

    rows = [
        row
        for row in records
        if normalize_url(
            row.get(
                "article_url",
                ""
            )
        )
        == article_url
    ]

    # If article is in completed list and contains zero
    # author rows, it may legitimately be an ad/TOC/etc.
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
            and not email
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
            "No author records to write."
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
    # EXCEL
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
            "G": 40,
            "H": 45,
            "I": 90,
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

    (
        records,
        completed_articles
    ) = load_checkpoint()

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

        driver = (
            connect_to_chrome()
        )

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

            (
                driver,
                issue_links
            ) = extract_issue_links(
                driver,
                journal,
                records,
                completed_articles
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

                driver = open_page(
                    driver,
                    issue_url,
                    records,
                    completed_articles,
                    ISSUE_LOAD_WAIT_MIN,
                    ISSUE_LOAD_WAIT_MAX
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
                    f"Issue date   : "
                    f"{issue_metadata.get('date')}"
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
                        and not (
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
                            "Skipping issue outside "
                            "requested year range."
                        )

                        continue

                # --------------------------------------------
                # ARTICLES
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

                    already_completed = (
                        article_url
                        in completed_articles
                    )

                    needs_reprocess = False

                    if already_completed:

                        needs_reprocess = (
                            article_needs_reprocess(
                                article_url,
                                records
                            )
                        )

                    if (
                        already_completed
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
                            f"REPROCESSING - missing email"
                        )

                    else:

                        print()
                        print(
                            f"[Article "
                            f"{article_index}/"
                            f"{len(article_urls)}]"
                        )

                    # ========================================
                    # PROCESS ARTICLE
                    # ========================================

                    try:

                        (
                            driver,
                            new_records
                        ) = parse_article(
                            driver,
                            article_url,
                            journal,
                            issue_metadata,
                            records,
                            completed_articles
                        )

                    except ScienceDirectBlocked as exc:

                        stop_reason = (
                            str(exc)
                        )

                        print()
                        print(
                            "=" * 100
                        )

                        print(
                            "SCIENCEDIRECT STILL BLOCKED"
                        )

                        print(
                            "=" * 100
                        )

                        print(
                            "The current article was NOT "
                            "marked completed."
                        )

                        print(
                            "Stopping safely."
                        )

                        raise

                    except BrowserSessionLost as exc:

                        stop_reason = (
                            str(exc)
                        )

                        print(
                            f"Browser session error: "
                            f"{exc}"
                        )

                        raise

                    except Exception as exc:

                        print(
                            f"ARTICLE ERROR: "
                            f"{exc}"
                        )

                        completed_articles.discard(
                            article_url
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

                        save_checkpoint(
                            records,
                            completed_articles
                        )

                        continue

                    # ========================================
                    # ARTICLE SUCCESS
                    # ========================================

                    # Do not remove old article rows until
                    # the new scrape completed successfully.

                    records = (
                        remove_article_records(
                            records,
                            article_url
                        )
                    )

                    accepted_records = []

                    for row in new_records:

                        row_year = (
                            row.get(
                                "year"
                            )
                        )

                        if row_year:

                            try:

                                row_year_int = int(
                                    row_year
                                )

                            except Exception:

                                row_year_int = None

                            if (
                                row_year_int
                                and not (
                                    journal[
                                        "start_year"
                                    ]
                                    <= row_year_int
                                    <= journal[
                                        "end_year"
                                    ]
                                )
                            ):

                                continue

                        accepted_records.append(
                            row
                        )

                    records.extend(
                        accepted_records
                    )

                    completed_articles.add(
                        article_url
                    )

                    successful_this_session += 1

                    print(
                        "    Article completed successfully."
                    )

                    # ----------------------------------------
                    # CHECKPOINT AFTER EVERY SUCCESS
                    # ----------------------------------------

                    save_checkpoint(
                        records,
                        completed_articles
                    )

                    # ----------------------------------------
                    # EXCEL PERIODIC
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
                    # COOLDOWN
                    # ----------------------------------------

                    if (
                        successful_this_session
                        % COOLDOWN_AFTER_ARTICLES
                        == 0
                    ):

                        cooldown = (
                            random.uniform(
                                COOLDOWN_SECONDS_MIN,
                                COOLDOWN_SECONDS_MAX
                            )
                        )

                        print()
                        print(
                            f"Processed "
                            f"{successful_this_session} "
                            f"articles this session."
                        )

                        print(
                            f"Cooling down for "
                            f"{cooldown:.1f}s..."
                        )

                        save_checkpoint(
                            records,
                            completed_articles
                        )

                        save_excel(
                            records
                        )

                        time.sleep(
                            cooldown
                        )

                    else:

                        random_wait(
                            ARTICLE_DELAY_MIN,
                            ARTICLE_DELAY_MAX,
                            "Before next article"
                        )

    except KeyboardInterrupt:

        stop_reason = (
            "Stopped manually."
        )

        print()
        print(
            stop_reason
        )

    except ScienceDirectBlocked:

        pass

    except BrowserSessionLost as exc:

        stop_reason = (
            f"Browser session lost: "
            f"{exc}"
        )

        print()
        print(
            stop_reason
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

        # ====================================================
        # FINAL SAVE
        # ====================================================

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