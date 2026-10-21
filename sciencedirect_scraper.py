import os
import re
import json
import time
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

PAGE_LOAD_WAIT = 4

ISSUE_LOAD_WAIT = 3

ARTICLE_DELAY = 2

AUTHOR_PANEL_OPEN_TIMEOUT = 15

AUTHOR_EMAIL_TIMEOUT = 25

AUTHOR_EMAIL_POLL_INTERVAL = 0.5

SAVE_EVERY_N_ARTICLES = 10


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
# CONNECT TO CHROME
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
# BLOCK DETECTION
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
    print("SCIENCEDIRECT ACCESS/BLOCK PAGE DETECTED")
    print("=" * 100)

    print(
        "Check the open Chrome window manually."
    )

    print(
        "Once the real ScienceDirect page is visible, "
        "the scraper will continue automatically."
    )

    while True:

        try:
            if not is_blocked(
                driver
            ):
                print(
                    "ScienceDirect content detected. Continuing..."
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
    wait_seconds=PAGE_LOAD_WAIT
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

            time.sleep(
                0.8
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
    journal
):
    print()
    print(
        "Opening journal issues page:"
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

            possible_year = extract_year(
                text
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
                elements[0].text
            )

    except Exception:
        pass

    try:
        body_text = driver.find_element(
            By.TAG_NAME,
            "body"
        ).text

    except Exception:
        body_text = ""

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

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

    result[
        "year"
    ] = (
        extract_year(
            result[
                "date"
            ]
        )
        or
        extract_year(
            body_text
        )
        or ""
    )

    # --------------------------------------------------------
    # VOLUME / ISSUE
    # --------------------------------------------------------

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

    else:

        volume_match = re.search(
            r"\bVolume\s+(\d+)",
            body_text,
            re.I
        )

        issue_match = re.search(
            r"\bIssue\s+(\d+)",
            body_text,
            re.I
        )

        if volume_match:
            result[
                "volume"
            ] = volume_match.group(
                1
            )

        if issue_match:
            result[
                "issue"
            ] = issue_match.group(
                1
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
        "h1",
    ]

    for selector in selectors:

        try:
            elements = driver.find_elements(
                By.CSS_SELECTOR,
                selector
            )

            for element in elements:

                value = clean_text(
                    element.text
                )

                if value:
                    return value

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
                    match.group(0)
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
                elements[0].text
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # METADATA
    # --------------------------------------------------------

    metadata_text = ""

    try:
        metadata_text = (
            driver.find_element(
                By.CSS_SELECTOR,
                ".publication-metadata"
            ).text
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
            match.group(1)
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
        r"Volume\s+(\d+)\s*,?\s*Issue\s+(\d+)",
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

    else:

        match = re.search(
            r"\bVolume\s+(\d+)",
            metadata_text,
            re.I
        )

        if match:

            result[
                "volume"
            ] = match.group(
                1
            )

        match = re.search(
            r"\bIssue\s+(\d+)",
            metadata_text,
            re.I
        )

        if match:

            result[
                "issue"
            ] = match.group(
                1
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
            ).strip(
                " ,"
            )

            if new_name != name:

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
        given = button.find_elements(
            By.CSS_SELECTOR,
            ".given-name"
        )

        if given:
            given_name = clean_text(
                given[0].text
            )

    except Exception:
        pass

    try:
        surnames = button.find_elements(
            By.CSS_SELECTOR,
            ".surname"
        )

        if surnames:
            surname = clean_text(
                surnames[0].text
            )

    except Exception:
        pass

    if given_name or surname:

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
# CONTACT ICON
# ============================================================

def author_has_email_icon(
    button
):
    """
    Exact ScienceDirect contact marker:
        svg.icon-envelope
    """

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

        if "icon-envelope" in html:
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
# EXACT AUTHOR PANEL
# ============================================================

def find_author_panel(
    driver
):
    try:
        panels = driver.find_elements(
            By.CSS_SELECTOR,
            "#side-panel.side-panel-is-expanded"
        )

        for panel in panels:

            try:
                if panel.is_displayed():
                    return panel

            except Exception:
                continue

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
                    "#side-panel.side-panel-is-expanded"
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

        # ----------------------------------------------------
        # EXACT EMAIL SELECTOR
        # ----------------------------------------------------

        try:
            links = driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel.side-panel-is-expanded "
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
        # E-ADDRESS TEXT FALLBACK
        # ----------------------------------------------------

        try:
            elements = driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel.side-panel-is-expanded "
                    ".e-address"
                )
            )

            for element in elements:

                text = (
                    element.text
                    or ""
                )

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
                "#side-panel.side-panel-is-expanded "
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
# CORRESPONDENCE ADDRESS
# ============================================================

def extract_correspondence_address(
    driver
):
    try:
        elements = driver.find_elements(
            By.CSS_SELECTOR,
            (
                "#side-panel.side-panel-is-expanded "
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
# AFFILIATION
# ============================================================

def extract_panel_affiliation(
    driver
):
    selectors = [
        (
            "#side-panel.side-panel-is-expanded "
            ".side-panel-children > "
            ".side-panel-author > "
            ".affiliation"
        ),
        (
            "#side-panel.side-panel-is-expanded "
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
# CLOSE PANEL
# ============================================================

def close_author_panel(
    driver
):
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

    # ESC fallback

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
    """
    Re-fetches the button because React frequently rebuilds
    author DOM nodes.
    """

    buttons = get_author_buttons(
        driver
    )

    if author_index >= len(
        buttons
    ):
        return None

    button = buttons[
        author_index
    ]

    try:
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

        driver.execute_script(
            "arguments[0].click();",
            button
        )

    except Exception:
        return None

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

        panel = None

        try:
            print(
                "          Clicking author..."
            )

            panel = open_author_panel(
                driver,
                index
            )

            # ------------------------------------------------
            # RETRY ONCE
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
                    "          Retrying..."
                )

                panel = open_author_panel(
                    driver,
                    index
                )

            # ------------------------------------------------
            # PANEL OPEN
            # ------------------------------------------------

            if panel is None:

                print(
                    "          PANEL NOT FOUND"
                )

            else:

                print(
                    "          Panel opened."
                )

                # --------------------------------------------
                # AUTHOR NAME FROM PANEL
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
                # CORRESPONDENCE ADDRESS
                # --------------------------------------------

                correspondence_address = (
                    extract_correspondence_address(
                        driver
                    )
                )

        except Exception as exc:

            print(
                f"          AUTHOR ERROR: "
                f"{exc}"
            )

        finally:

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
    issue_metadata
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

    time.sleep(
        1
    )

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

    title = extract_article_title(
        driver
    )

    metadata = extract_article_metadata(
        driver
    )

    # --------------------------------------------------------
    # FALLBACKS FROM ISSUE
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
# CHECK WHETHER ARTICLE NEEDS REPROCESSING
# ============================================================

def article_needs_reprocess(
    article_url,
    records
):
    """
    Reprocess an already-completed article when a contact
    author was detected but its email remained blank.
    """

    matching = [
        row
        for row in records
        if normalize_url(
            row.get(
                "article_url",
                ""
            )
        ) == normalize_url(
            article_url
        )
    ]

    if not matching:
        return True

    for row in matching:

        if (
            str(
                row.get(
                    "has_contact_icon",
                    ""
                )
            ).strip().lower()
            == "yes"
            and
            not str(
                row.get(
                    "email",
                    ""
                )
            ).strip()
        ):

            return True

    return False


# ============================================================
# REMOVE OLD ARTICLE RECORDS BEFORE REPROCESS
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
        ) != article_url
    ]


# ============================================================
# SAVE EXCEL
# ============================================================

def save_excel(
    records
):
    if not records:

        print(
            "No records available for Excel."
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
    # CLEAN EMAIL
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

    # --------------------------------------------------------
    # CLEAN AUTHOR
    # --------------------------------------------------------

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
    # REMOVE DUPLICATES
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
    # EMAIL SHEET
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
    # WRITE EXCEL
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
            "G": 35,
            "H": 45,
            "I": 85,
            "J": 100,
            "K": 20,
            "L": 15,
            "M": 38,
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

    driver = connect_to_chrome()

    try:

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
            # ISSUES
            # ------------------------------------------------

            issue_links = extract_issue_links(
                driver,
                journal
            )

            print(
                f"Issue links collected: "
                f"{len(issue_links)}"
            )

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

                try:

                    if not open_page(
                        driver,
                        issue_url,
                        ISSUE_LOAD_WAIT
                    ):
                        continue

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

                    # ----------------------------------------
                    # YEAR FILTER
                    # ----------------------------------------

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
                                <=
                                issue_year_int
                                <=
                                journal[
                                    "end_year"
                                ]
                            )
                        ):

                            print(
                                "Skipping issue outside "
                                "requested year range."
                            )

                            continue

                    # ----------------------------------------
                    # ARTICLES
                    # ----------------------------------------

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

                        # ====================================
                        # COMPLETED / REPROCESS DECISION
                        # ====================================

                        already_completed = (
                            article_url
                            in completed_articles
                        )

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

                        if (
                            already_completed
                            and
                            needs_reprocess
                        ):

                            print()
                            print(
                                f"[Article "
                                f"{article_index}/"
                                f"{len(article_urls)}] "
                                f"REPROCESSING - "
                                f"missing contact email"
                            )

                            records = (
                                remove_article_records(
                                    records,
                                    article_url
                                )
                            )

                            completed_articles.discard(
                                article_url
                            )

                        else:

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
                                    journal,
                                    issue_metadata
                                )
                            )

                            accepted_records = []

                            # --------------------------------
                            # YEAR FILTER
                            # --------------------------------

                            for record in (
                                article_records
                            ):

                                record_year = (
                                    record.get(
                                        "year"
                                    )
                                )

                                if record_year:

                                    try:
                                        record_year_int = int(
                                            record_year
                                        )

                                    except Exception:
                                        record_year_int = None

                                    if (
                                        record_year_int
                                        and
                                        not (
                                            journal[
                                                "start_year"
                                            ]
                                            <=
                                            record_year_int
                                            <=
                                            journal[
                                                "end_year"
                                            ]
                                        )
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

                            # --------------------------------
                            # CHECKPOINT
                            # --------------------------------

                            save_checkpoint(
                                records,
                                completed_articles
                            )

                            # --------------------------------
                            # PERIODIC EXCEL
                            # --------------------------------

                            if (
                                len(
                                    completed_articles
                                )
                                %
                                SAVE_EVERY_N_ARTICLES
                                == 0
                            ):

                                save_excel(
                                    records
                                )

                            time.sleep(
                                ARTICLE_DELAY
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

    except KeyboardInterrupt:

        print()
        print(
            "Stopped manually."
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
            "FINISHED"
        )

        print(
            "=" * 110
        )

        print(
            f"Completed articles : "
            f"{len(completed_articles):,}"
        )

        print(
            f"Raw author records : "
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