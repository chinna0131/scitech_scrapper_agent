import os
import re
import json
import time
import random
import socket
import subprocess
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
)


# ============================================================
# OUTPUT PATHS
# ============================================================

OUTPUT_FOLDER = "sciencedirect_output"

OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "sciencedirect_authors_emails.xlsx",
)

CHECKPOINT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "checkpoint.json",
)

ERROR_FILE = os.path.join(
    OUTPUT_FOLDER,
    "errors.json",
)


# ============================================================
# CHROME CONFIG
# ============================================================

CHROME_EXE = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe"
)

CHROME_PROFILE_DIR = (
    r"C:\chrome-sciencedirect"
)

REMOTE_DEBUGGING_PORT = 9222

REMOTE_DEBUGGING_ADDRESS = (
    f"127.0.0.1:{REMOTE_DEBUGGING_PORT}"
)

chrome_process = None


# ============================================================
# NORMAL PACING
# ============================================================

PAGE_LOAD_WAIT_MIN = 4
PAGE_LOAD_WAIT_MAX = 7

ISSUE_LOAD_WAIT_MIN = 4
ISSUE_LOAD_WAIT_MAX = 7

ARTICLE_DELAY_MIN = 4
ARTICLE_DELAY_MAX = 7

AUTHOR_CLICK_DELAY_MIN = 0.8
AUTHOR_CLICK_DELAY_MAX = 1.5

AUTHOR_PANEL_OPEN_TIMEOUT = 15

AUTHOR_EMAIL_TIMEOUT = 25

AUTHOR_EMAIL_POLL_INTERVAL = 0.5


# ============================================================
# PERIODIC COOLDOWN
# ============================================================

COOLDOWN_AFTER_ARTICLES = 20

COOLDOWN_SECONDS_MIN = 30
COOLDOWN_SECONDS_MAX = 45


# ============================================================
# BLOCK RECOVERY
# ============================================================

# First block retry:
#   close Chrome
#   wait 3 minutes
#   fresh Chrome
#
# Second block retry:
#   close Chrome
#   wait 7 minutes
#   fresh Chrome
#
# Still blocked:
#   stop safely
#   checkpoint remains intact

BLOCK_RESTART_WAITS = [
    180,   # 3 minutes
    420,   # 7 minutes
]

MAX_BLOCK_RESTARTS = len(
    BLOCK_RESTART_WAITS
)


# ============================================================
# SAVE FREQUENCY
# ============================================================

SAVE_EVERY_N_ARTICLES = 5


# ============================================================
# JOURNAL CONFIG
# ============================================================

JOURNALS = [
    {
        "name":
            "Journal of PeriAnesthesia Nursing",

        "url":
            "https://www.sciencedirect.com/journal/"
            "journal-of-perianesthesia-nursing/issues",

        "start_year":
            2025,

        "end_year":
            2026,
    },

    {
        "name":
            "Journal of Clinical Anesthesia",

        "url":
            "https://www.sciencedirect.com/journal/"
            "journal-of-clinical-anesthesia/issues",

        "start_year":
            2025,

        "end_year":
            2027,
    },

    {
        "name":
            "Journal of Cardiothoracic and Vascular Anesthesia",

        "url":
            "https://www.sciencedirect.com/journal/"
            "journal-of-cardiothoracic-and-vascular-anesthesia/"
            "issues",

        "start_year":
            2025,

        "end_year":
            2026,
    },

    {
        "name":
            "Journal of Anesthesia and Translational Medicine",

        "url":
            "https://www.sciencedirect.com/journal/"
            "journal-of-anesthesia-and-translational-medicine/"
            "issues",

        "start_year":
            2025,

        "end_year":
            2026,
    },

    {
        "name":
            "International Journal of Obstetric Anesthesia",

        "url":
            "https://www.sciencedirect.com/journal/"
            "international-journal-of-obstetric-anesthesia/"
            "issues",

        "start_year":
            2025,

        "end_year":
            2026,
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


def random_wait(
    minimum,
    maximum,
    message=None,
):

    seconds = random.uniform(
        minimum,
        maximum,
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
    data,
):

    with open(
        path,
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            data,
            file,
            ensure_ascii=False,
            indent=2,
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
            encoding="utf-8",
        ) as file:

            return json.load(
                file
            )

    except Exception:

        return {}


# ============================================================
# CHROME DEBUG PORT
# ============================================================

def chrome_debug_port_open():

    try:

        with socket.create_connection(
            (
                "127.0.0.1",
                REMOTE_DEBUGGING_PORT,
            ),
            timeout=2,
        ):

            return True

    except OSError:

        return False


# ============================================================
# FIND DEDICATED CHROME PIDS
# ============================================================

def find_sciencedirect_chrome_pids():

    powershell = f"""
$processes = Get-CimInstance Win32_Process |
Where-Object {{
    $_.Name -eq 'chrome.exe' -and
    $_.CommandLine -and
    (
        $_.CommandLine -like '*--remote-debugging-port={REMOTE_DEBUGGING_PORT}*' -or
        $_.CommandLine -like '*--user-data-dir={CHROME_PROFILE_DIR}*'
    )
}}

$processes |
ForEach-Object {{
    Write-Output $_.ProcessId
}}
"""

    try:

        result = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                powershell,
            ],
            capture_output=True,
            text=True,
            timeout=20,
        )

        pids = []

        for line in (
            result.stdout
            or ""
        ).splitlines():

            line = line.strip()

            if line.isdigit():

                pids.append(
                    int(line)
                )

        return list(
            dict.fromkeys(
                pids
            )
        )

    except Exception as exc:

        print(
            "Could not inspect Chrome "
            f"processes: {exc}"
        )

        return []


# ============================================================
# FORCE CLOSE ONLY SCRAPER CHROME
# ============================================================

def force_close_sciencedirect_chrome():

    pids = (
        find_sciencedirect_chrome_pids()
    )

    if not pids:

        print(
            "No dedicated ScienceDirect "
            "Chrome PID found."
        )

        return True

    print(
        "Dedicated Chrome PID(s): "
        f"{pids}"
    )

    for pid in pids:

        try:

            subprocess.run(
                [
                    "taskkill",
                    "/PID",
                    str(pid),
                    "/T",
                    "/F",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )

        except Exception as exc:

            print(
                f"Unable to close PID "
                f"{pid}: {exc}"
            )

    time.sleep(
        3
    )

    return (
        not chrome_debug_port_open()
    )


# ============================================================
# CLOSE DEDICATED CHROME
# ============================================================

def close_sciencedirect_chrome(
    driver=None,
):

    global chrome_process

    print()
    print(
        "=" * 100
    )

    print(
        "CLOSING DEDICATED SCIENCEDIRECT CHROME"
    )

    print(
        "=" * 100
    )

    # --------------------------------------------------------
    # METHOD 1: DevTools Browser.close
    # --------------------------------------------------------

    if driver is not None:

        try:

            print(
                "Requesting Browser.close "
                "through DevTools..."
            )

            driver.execute_cdp_cmd(
                "Browser.close",
                {},
            )

            time.sleep(
                2
            )

        except Exception as exc:

            print(
                "Browser.close response: "
                f"{exc}"
            )

        try:

            driver.quit()

        except Exception:

            pass

    # --------------------------------------------------------
    # WAIT FOR PORT TO DISAPPEAR
    # --------------------------------------------------------

    print(
        "Waiting for port 9222 to close..."
    )

    for second in range(
        1,
        8,
    ):

        if not chrome_debug_port_open():

            print(
                "Chrome closed successfully."
            )

            chrome_process = None

            time.sleep(
                2
            )

            return True

        print(
            "    Chrome still running "
            f"{second}/7"
        )

        time.sleep(
            1
        )

    # --------------------------------------------------------
    # FORCE-CLOSE ONLY DEDICATED PROFILE
    # --------------------------------------------------------

    print(
        "Graceful close failed."
    )

    print(
        "Force-closing dedicated "
        "ScienceDirect Chrome..."
    )

    force_close_sciencedirect_chrome()

    for second in range(
        1,
        15,
    ):

        if not chrome_debug_port_open():

            print(
                "Dedicated Chrome fully closed."
            )

            chrome_process = None

            time.sleep(
                2
            )

            return True

        print(
            "    Waiting for port "
            f"9222: {second}/14"
        )

        time.sleep(
            1
        )

    print(
        "WARNING: port 9222 "
        "is still active."
    )

    return False


# ============================================================
# START DEDICATED CHROME
# ============================================================

def start_sciencedirect_chrome():

    global chrome_process

    # --------------------------------------------------------
    # OLD INSTANCE?
    # --------------------------------------------------------

    if chrome_debug_port_open():

        print(
            "Old dedicated Chrome detected."
        )

        close_sciencedirect_chrome(
            None
        )

    if chrome_debug_port_open():

        force_close_sciencedirect_chrome()

        time.sleep(
            3
        )

    if chrome_debug_port_open():

        raise RuntimeError(
            "Unable to close previous "
            "Chrome on port 9222."
        )

    # --------------------------------------------------------
    # START
    # --------------------------------------------------------

    print()
    print(
        "=" * 100
    )

    print(
        "STARTING NEW SCIENCEDIRECT CHROME"
    )

    print(
        "=" * 100
    )

    if not os.path.exists(
        CHROME_EXE
    ):

        raise RuntimeError(
            "Chrome not found: "
            f"{CHROME_EXE}"
        )

    chrome_process = subprocess.Popen(
        [
            CHROME_EXE,

            (
                "--remote-debugging-port="
                f"{REMOTE_DEBUGGING_PORT}"
            ),

            (
                "--user-data-dir="
                f"{CHROME_PROFILE_DIR}"
            ),

            "--no-first-run",

            "--no-default-browser-check",

            "https://www.sciencedirect.com/",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    print(
        "Chrome PID: "
        f"{chrome_process.pid}"
    )

    for second in range(
        1,
        31,
    ):

        if chrome_debug_port_open():

            print(
                "Chrome debugging port "
                "9222 is ready."
            )

            time.sleep(
                3
            )

            return True

        print(
            "    Waiting for Chrome "
            f"{second}/30"
        )

        time.sleep(
            1
        )

    raise RuntimeError(
        "Chrome did not become ready."
    )


# ============================================================
# CONNECT SELENIUM
# ============================================================

def connect_to_chrome():

    if not chrome_debug_port_open():

        start_sciencedirect_chrome()

    print(
        "Connecting Selenium to "
        f"{REMOTE_DEBUGGING_ADDRESS}..."
    )

    options = Options()

    options.add_experimental_option(
        "debuggerAddress",
        REMOTE_DEBUGGING_ADDRESS,
    )

    for attempt in range(
        1,
        6,
    ):

        try:

            driver = webdriver.Chrome(
                options=options
            )

            handles = (
                driver.window_handles
            )

            if not handles:

                raise RuntimeError(
                    "Chrome has no browser tab."
                )

            driver.switch_to.window(
                handles[-1]
            )

            print(
                "Connected to Chrome."
            )

            return driver

        except Exception as exc:

            print(
                "Chrome connection "
                f"{attempt}/5 failed: "
                f"{exc}"
            )

            time.sleep(
                2
            )

    raise BrowserSessionLost(
        "Unable to connect Selenium "
        "to Chrome."
    )


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


def ensure_driver_alive(driver):

    if not driver_is_alive(
        driver
    ):

        raise BrowserSessionLost(
            "Chrome/Selenium window "
            "is unavailable."
        )


# ============================================================
# BLOCK DETECTION
# ============================================================

def is_sciencedirect_blocked(
    driver,
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
        (
            "there was a problem providing "
            "the content you requested"
        ),
        "cpe00001",
        "cpe00002",
        "reference number:",
        "access denied",
        "unusual traffic",
    ]

    return any(
        indicator in html
        or indicator in title
        for indicator
        in indicators
    )


# ============================================================
# CHECKPOINT
# ============================================================

def save_checkpoint(
    records,
    completed_articles,
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
        },
    )


def load_checkpoint():

    data = load_json(
        CHECKPOINT_FILE
    )

    records = data.get(
        "records",
        [],
    )

    completed_articles = set(
        normalize_url(
            url
        )
        for url
        in data.get(
            "completed_articles",
            [],
        )
        if url
    )

    return (
        records,
        completed_articles,
    )


# ============================================================
# AUTHOR NAME CLEANING
# ============================================================

def clean_author_name(
    name,
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

        for qualification in qualifications:

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
                flags=re.I,
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
# SAVE EXCEL
# ============================================================

def save_excel(records):

    if not records:

        print(
            "No author records "
            "available for Excel."
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
        inplace=True,
    )

    # --------------------------------------------------------
    # EMAIL-ONLY SHEET
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
        inplace=True,
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
        na_position="last",
    )

    email_df.sort_values(
        by=[
            "journal",
            "year",
            "author",
        ],
        inplace=True,
        na_position="last",
    )

    # --------------------------------------------------------
    # WRITE
    # --------------------------------------------------------

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="All Authors",
            index=False,
        )

        email_df.to_excel(
            writer,
            sheet_name="Authors With Email",
            index=False,
        )

        workbook = (
            writer.book
        )

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
                width,
            ) in widths.items():

                worksheet.column_dimensions[
                    column
                ].width = width

    print()
    print(
        "Excel saved: "
        f"{os.path.abspath(OUTPUT_FILE)}"
    )

    print(
        "Total author rows: "
        f"{len(df):,}"
    )

    print(
        "Rows with email: "
        f"{len(email_df):,}"
    )


# ============================================================
# BLOCK RECOVERY
# ============================================================

def restart_after_sciencedirect_block(
    driver,
    blocked_url,
    records,
    completed_articles,
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
        "Blocked URL:"
    )

    print(
        blocked_url
    )

    print()
    print(
        "Saving checkpoint and Excel..."
    )

    save_checkpoint(
        records,
        completed_articles,
    )

    save_excel(
        records
    )

    # ========================================================
    # TWO FAST RECOVERY ATTEMPTS
    # ========================================================

    for (
        attempt,
        wait_seconds,
    ) in enumerate(
        BLOCK_RESTART_WAITS,
        start=1,
    ):

        print()
        print(
            "=" * 100
        )

        print(
            "BLOCK RECOVERY "
            f"{attempt}/"
            f"{MAX_BLOCK_RESTARTS}"
        )

        print(
            "=" * 100
        )

        # ----------------------------------------------------
        # CLOSE OLD CHROME
        # ----------------------------------------------------

        closed = (
            close_sciencedirect_chrome(
                driver
            )
        )

        driver = None

        if not closed:

            print(
                "Chrome did not close "
                "normally."
            )

            force_close_sciencedirect_chrome()

            time.sleep(
                3
            )

        # ----------------------------------------------------
        # ENSURE PORT CLOSED
        # ----------------------------------------------------

        if chrome_debug_port_open():

            force_close_sciencedirect_chrome()

            time.sleep(
                3
            )

        if chrome_debug_port_open():

            print(
                "Could not fully close old Chrome."
            )

            continue

        print(
            "Chrome is fully closed."
        )

        # ----------------------------------------------------
        # WAIT
        # ----------------------------------------------------

        minutes = (
            wait_seconds
            / 60
        )

        print(
            f"Waiting {minutes:g} minute(s) "
            f"before restart..."
        )

        remaining = (
            wait_seconds
        )

        while remaining > 0:

            step = min(
                30,
                remaining,
            )

            time.sleep(
                step
            )

            remaining -= step

            if remaining:

                print(
                    "    "
                    f"{remaining}s remaining..."
                )

        # ----------------------------------------------------
        # START FRESH CHROME
        # ----------------------------------------------------

        try:

            start_sciencedirect_chrome()

            driver = (
                connect_to_chrome()
            )

        except Exception as exc:

            print(
                "Chrome restart failed: "
                f"{exc}"
            )

            driver = None

            continue

        # ----------------------------------------------------
        # RETRY EXACT SAME URL
        # ----------------------------------------------------

        print()
        print(
            "Retrying blocked URL:"
        )

        print(
            blocked_url
        )

        try:

            driver.get(
                blocked_url
            )

            random_wait(
                PAGE_LOAD_WAIT_MIN,
                PAGE_LOAD_WAIT_MAX,
                "Retry page load wait",
            )

        except Exception as exc:

            print(
                "Retry navigation failed: "
                f"{exc}"
            )

            continue

        # ----------------------------------------------------
        # SUCCESS?
        # ----------------------------------------------------

        if not is_sciencedirect_blocked(
            driver
        ):

            print()
            print(
                "=" * 100
            )

            print(
                "SCIENCEDIRECT ACCESS RESTORED"
            )

            print(
                "=" * 100
            )

            return (
                driver,
                True,
            )

        print(
            "Still blocked after "
            "fresh Chrome."
        )

        save_checkpoint(
            records,
            completed_articles,
        )

    # ========================================================
    # TWO ATTEMPTS FAILED
    # ========================================================

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
        "Saving progress and exiting."
    )

    print(
        "Current URL remains unfinished."
    )

    save_checkpoint(
        records,
        completed_articles,
    )

    save_excel(
        records
    )

    close_sciencedirect_chrome(
        driver
    )

    return (
        None,
        False,
    )


# ============================================================
# OPEN PAGE
# ============================================================

def open_page(
    driver,
    url,
    records,
    completed_articles,
    minimum_wait=PAGE_LOAD_WAIT_MIN,
    maximum_wait=PAGE_LOAD_WAIT_MAX,
):

    # --------------------------------------------------------
    # ENSURE CHROME
    # --------------------------------------------------------

    if not driver_is_alive(
        driver
    ):

        if not chrome_debug_port_open():

            start_sciencedirect_chrome()

        driver = (
            connect_to_chrome()
        )

    print()
    print(
        f"Opening: {url}"
    )

    # --------------------------------------------------------
    # NAVIGATE
    # --------------------------------------------------------

    try:

        driver.get(
            url
        )

    except (
        NoSuchWindowException,
        InvalidSessionIdException,
    ):

        close_sciencedirect_chrome(
            driver
        )

        start_sciencedirect_chrome()

        driver = (
            connect_to_chrome()
        )

        driver.get(
            url
        )

    except WebDriverException as exc:

        if not driver_is_alive(
            driver
        ):

            close_sciencedirect_chrome(
                driver
            )

            start_sciencedirect_chrome()

            driver = (
                connect_to_chrome()
            )

            driver.get(
                url
            )

        else:

            raise exc

    random_wait(
        minimum_wait,
        maximum_wait,
        "Page load wait",
    )

    ensure_driver_alive(
        driver
    )

    # --------------------------------------------------------
    # BLOCK?
    # --------------------------------------------------------

    if is_sciencedirect_blocked(
        driver
    ):

        (
            driver,
            recovered,
        ) = restart_after_sciencedirect_block(
            driver,
            url,
            records,
            completed_articles,
        )

        if not recovered:

            raise ScienceDirectBlocked(
                "ScienceDirect remained blocked "
                "after 3-minute and 7-minute recovery."
            )

    return driver


# ============================================================
# SCROLL ISSUE PAGE
# ============================================================

def scroll_page(
    driver,
    max_scrolls=12,
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
                    1.5,
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
    completed_articles,
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

    driver = (
        open_page(
            driver,
            journal[
                "url"
            ],
            records,
            completed_articles,
        )
    )

    scroll_page(
        driver
    )

    links = []
    seen = set()

    anchors = (
        driver.find_elements(
            By.CSS_SELECTOR,
            (
                'a[href*="/vol/"]'
                '[href*="/issue/"]'
            ),
        )
    )

    print(
        "Potential issue links found: "
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

    return (
        driver,
        links,
    )


# ============================================================
# ISSUE METADATA
# ============================================================

def extract_issue_metadata(
    driver,
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
                ".publication-title",
            )
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
                "body",
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
    # VOLUME + ISSUE FROM URL
    # --------------------------------------------------------

    match = re.search(
        r"/vol/(\d+)/issue/(\d+)",
        current_url,
        re.I,
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
        (
            r"(January|February|March|April|May|June|"
            r"July|August|September|October|November|December)"
            r"\s+(20\d{2})"
        ),
        body_text,
        re.I,
    )

    if match:

        result[
            "date"
        ] = clean_text(
            match.group(
                0
            )
        )

        result[
            "year"
        ] = int(
            match.group(
                2
            )
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
# ARTICLE LINKS
# ============================================================

def extract_article_urls(
    driver,
):

    urls = []
    seen = set()

    elements = (
        driver.find_elements(
            By.CSS_SELECTOR,
            (
                'a.article-content-title'
                '[href*="/science/article/pii/"]'
            ),
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
# TITLE
# ============================================================

def extract_article_title(
    driver,
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
        "h1",
    ]

    for selector in selectors:

        try:

            elements = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    selector,
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
    driver,
):

    try:

        links = (
            driver.find_elements(
                By.CSS_SELECTOR,
                'a[href*="doi.org/10."]',
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
                re.I,
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
    driver,
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

        elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                ".publication-title",
            )
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
    # METADATA TEXT
    # --------------------------------------------------------

    try:

        metadata_text = (
            driver.find_element(
                By.CSS_SELECTOR,
                ".publication-metadata",
            ).text
        )

    except Exception:

        try:

            metadata_text = (
                driver.find_element(
                    By.TAG_NAME,
                    "body",
                ).text
            )

        except Exception:

            metadata_text = ""

    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    match = re.search(
        r"Date:\s*"
        r"([A-Za-z]+\s+20\d{2})",
        metadata_text,
        re.I,
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
    # VOLUME + ISSUE
    # --------------------------------------------------------

    match = re.search(
        (
            r"Volume\s*:?\s*"
            r"(?:Volume\s*)?"
            r"(\d+)"
            r"\s*,?\s*Issue\s+(\d+)"
        ),
        metadata_text,
        re.I,
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

    result[
        "doi"
    ] = extract_doi(
        driver
    )

    return result


# ============================================================
# AUTHOR BUTTONS
# ============================================================

def get_author_buttons(
    driver,
):

    return driver.find_elements(
        By.CSS_SELECTOR,
        (
            'button'
            '[data-sd-ui-side-panel-opener="true"]'
            '[data-xocs-content-type="author"]'
        ),
    )


def author_name_from_button(
    button,
):

    given_name = ""
    surname = ""

    try:

        elements = (
            button.find_elements(
                By.CSS_SELECTOR,
                ".given-name",
            )
        )

        if elements:

            given_name = clean_text(
                elements[
                    0
                ].text
            )

    except Exception:

        pass

    try:

        elements = (
            button.find_elements(
                By.CSS_SELECTOR,
                ".surname",
            )
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
    button,
):

    try:

        icons = (
            button.find_elements(
                By.CSS_SELECTOR,
                "svg.icon-envelope",
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
            "author email or social media "
            "contact details icon"
            in html
        ):

            return True

    except Exception:

        pass

    return False


# ============================================================
# AUTHOR PANEL
# ============================================================

def wait_for_author_panel(
    driver,
):

    try:

        return WebDriverWait(
            driver,
            AUTHOR_PANEL_OPEN_TIMEOUT,
        ).until(
            EC.visibility_of_element_located(
                (
                    By.CSS_SELECTOR,
                    (
                        "#side-panel."
                        "side-panel-is-expanded"
                    ),
                )
            )
        )

    except TimeoutException:

        return None


# ============================================================
# EMAIL
# ============================================================

def wait_for_author_email(
    driver,
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
        # EXACT mailto:
        # ----------------------------------------------------

        try:

            links = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    (
                        "#side-panel."
                        "side-panel-is-expanded "
                        ".e-address "
                        'a[href^="mailto:"]'
                    ),
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
                        "",
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

        # ----------------------------------------------------
        # TEXT FALLBACK
        # ----------------------------------------------------

        try:

            elements = (
                driver.find_elements(
                    By.CSS_SELECTOR,
                    (
                        "#side-panel."
                        "side-panel-is-expanded "
                        ".e-address"
                    ),
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
# PANEL AUTHOR
# ============================================================

def extract_panel_author_name(
    driver,
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
                ),
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
# PANEL AFFILIATION
# ============================================================

def extract_panel_affiliation(
    driver,
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
                    selector,
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
    driver,
):

    try:

        elements = (
            driver.find_elements(
                By.CSS_SELECTOR,
                (
                    "#side-panel."
                    "side-panel-is-expanded "
                    ".correspondence span"
                ),
            )
        )

        for element in elements:

            text = clean_text(
                element.text
            )

            if not text:

                continue

            text = re.sub(
                (
                    r"^Address correspondence "
                    r"to\s*:\s*"
                ),
                "",
                text,
                flags=re.I,
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
    driver,
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
                ),
            )
        )

        for button in buttons:

            try:

                if not button.is_displayed():

                    continue

                driver.execute_script(
                    "arguments[0].click();",
                    button,
                )

                try:

                    WebDriverWait(
                        driver,
                        5,
                    ).until(
                        EC.invisibility_of_element_located(
                            (
                                By.CSS_SELECTOR,
                                (
                                    "#side-panel."
                                    "side-panel-is-expanded"
                                ),
                            )
                        )
                    )

                except Exception:

                    pass

                time.sleep(
                    0.4
                )

                return True

            except Exception:

                continue

    except Exception:

        pass

    try:

        driver.find_element(
            By.TAG_NAME,
            "body",
        ).send_keys(
            Keys.ESCAPE
        )

        time.sleep(
            0.4
        )

    except Exception:

        pass

    return False


# ============================================================
# OPEN AUTHOR PANEL
# ============================================================

def open_author_panel(
    driver,
    author_index,
):

    ensure_driver_alive(
        driver
    )

    buttons = (
        get_author_buttons(
            driver
        )
    )

    if (
        author_index
        >= len(
            buttons
        )
    ):

        return None

    button = (
        buttons[
            author_index
        ]
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        button,
    )

    random_wait(
        AUTHOR_CLICK_DELAY_MIN,
        AUTHOR_CLICK_DELAY_MAX,
        "          Author click wait",
    )

    driver.execute_script(
        "arguments[0].click();",
        button,
    )

    return (
        wait_for_author_panel(
            driver
        )
    )


# ============================================================
# EXTRACT ARTICLE AUTHORS
# ============================================================

def extract_article_authors(
    driver,
):

    results = []

    buttons = (
        get_author_buttons(
            driver
        )
    )

    total_authors = len(
        buttons
    )

    print(
        "    Author buttons found: "
        f"{total_authors}"
    )

    for index in range(
        total_authors
    ):

        ensure_driver_alive(
            driver
        )

        # ScienceDirect/React may recreate the elements.
        buttons = (
            get_author_buttons(
                driver
            )
        )

        if (
            index
            >= len(
                buttons
            )
        ):

            break

        button = (
            buttons[
                index
            ]
        )

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
            "          Contact icon: "
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
                    index,
                )
            )

            # -----------------------------------------------
            # ONE PANEL RETRY
            # -----------------------------------------------

            if panel is None:

                print(
                    "          First panel "
                    "attempt failed."
                )

                close_author_panel(
                    driver
                )

                time.sleep(
                    0.8
                )

                panel = (
                    open_author_panel(
                        driver,
                        index,
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
                        "          Panel author: "
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
                        "          EMAIL FOUND: "
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
                        "          Affiliation: "
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
                "          AUTHOR ERROR: "
                f"{exc}"
            )

        except Exception as exc:

            print(
                "          AUTHOR ERROR: "
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
                        0.5,
                        1.0,
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
    completed_articles,
):

    print()
    print(
        "-" * 100
    )

    print(
        f"ARTICLE: {article_url}"
    )

    driver = (
        open_page(
            driver,
            article_url,
            records,
            completed_articles,
        )
    )

    ensure_driver_alive(
        driver
    )

    try:

        WebDriverWait(
            driver,
            10,
        ).until(
            EC.presence_of_element_located(
                (
                    By.CSS_SELECTOR,
                    (
                        "h1.content-title, "
                        "#screen-reader-main-title, "
                        "h1"
                    ),
                )
            )
        )

    except TimeoutException:

        pass

    if is_sciencedirect_blocked(
        driver
    ):

        raise ScienceDirectBlocked(
            "Article remained blocked."
        )

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
            "",
        )

    if not metadata[
        "date"
    ]:

        metadata[
            "date"
        ] = issue_metadata.get(
            "date",
            "",
        )

    if not metadata[
        "volume"
    ]:

        metadata[
            "volume"
        ] = issue_metadata.get(
            "volume",
            "",
        )

    if not metadata[
        "issue"
    ]:

        metadata[
            "issue"
        ] = issue_metadata.get(
            "issue",
            "",
        )

    print(
        "    Title  : "
        f"{title}"
    )

    print(
        "    Journal: "
        f"{metadata['journal']}"
    )

    print(
        "    Date   : "
        f"{metadata['date']}"
    )

    print(
        "    Volume : "
        f"{metadata['volume']}"
    )

    print(
        "    Issue  : "
        f"{metadata['issue']}"
    )

    print(
        "    DOI    : "
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
        rows,
    )


# ============================================================
# REMOVE EXISTING ARTICLE RECORDS
# ============================================================

def remove_article_records(
    records,
    article_url,
):

    article_url = (
        normalize_url(
            article_url
        )
    )

    return [
        row
        for row in records
        if normalize_url(
            row.get(
                "article_url",
                "",
            )
        )
        != article_url
    ]


# ============================================================
# REPROCESS CHECK
# ============================================================

def article_needs_reprocess(
    article_url,
    records,
):

    article_url = (
        normalize_url(
            article_url
        )
    )

    rows = [
        row
        for row in records
        if normalize_url(
            row.get(
                "article_url",
                "",
            )
        )
        == article_url
    ]

    # Legitimate page with zero author rows.
    if not rows:

        return False

    for row in rows:

        contact = (
            str(
                row.get(
                    "has_contact_icon",
                    "",
                )
            )
            .strip()
            .lower()
        )

        email = (
            str(
                row.get(
                    "email",
                    "",
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
# MAIN
# ============================================================

def main():

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True,
    )

    (
        records,
        completed_articles,
    ) = load_checkpoint()

    errors = (
        load_json(
            ERROR_FILE
        )
    )

    if not isinstance(
        errors,
        list,
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
        "Existing records : "
        f"{len(records):,}"
    )

    print(
        "Completed articles: "
        f"{len(completed_articles):,}"
    )

    print()

    driver = None

    successful_this_session = 0

    stop_reason = ""

    try:

        # ====================================================
        # SCRIPT OWNS CHROME
        # ====================================================

        start_sciencedirect_chrome()

        driver = (
            connect_to_chrome()
        )

        # ====================================================
        # JOURNALS
        # ====================================================

        for (
            journal_index,
            journal,
        ) in enumerate(
            JOURNALS,
            start=1,
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
                "Years: "
                f"{journal['start_year']} - "
                f"{journal['end_year']}"
            )

            print(
                "=" * 110
            )

            (
                driver,
                issue_links,
            ) = extract_issue_links(
                driver,
                journal,
                records,
                completed_articles,
            )

            print(
                "Issue links collected: "
                f"{len(issue_links)}"
            )

            # =================================================
            # ISSUES
            # =================================================

            for (
                issue_index,
                issue_url,
            ) in enumerate(
                issue_links,
                start=1,
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

                driver = (
                    open_page(
                        driver,
                        issue_url,
                        records,
                        completed_articles,
                        ISSUE_LOAD_WAIT_MIN,
                        ISSUE_LOAD_WAIT_MAX,
                    )
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
                    "Issue date   : "
                    f"{issue_metadata.get('date')}"
                )

                print(
                    "Issue year   : "
                    f"{issue_year}"
                )

                print(
                    "Issue volume : "
                    f"{issue_metadata.get('volume')}"
                )

                print(
                    "Issue number : "
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
                            "Skipping issue outside "
                            "requested year range."
                        )

                        continue

                # --------------------------------------------
                # ARTICLE LIST
                # --------------------------------------------

                article_urls = (
                    extract_article_urls(
                        driver
                    )
                )

                print(
                    "Articles found: "
                    f"{len(article_urls)}"
                )

                # ============================================
                # ARTICLES
                # ============================================

                for (
                    article_index,
                    article_url,
                ) in enumerate(
                    article_urls,
                    start=1,
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
                                records,
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
                            "Already processed"
                        )

                        continue

                    if needs_reprocess:

                        print()
                        print(
                            f"[Article "
                            f"{article_index}/"
                            f"{len(article_urls)}] "
                            "REPROCESSING - "
                            "missing contact email"
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
                            new_records,
                        ) = parse_article(
                            driver,
                            article_url,
                            journal,
                            issue_metadata,
                            records,
                            completed_articles,
                        )

                    except ScienceDirectBlocked as exc:

                        stop_reason = (
                            str(exc)
                        )

                        completed_articles.discard(
                            article_url
                        )

                        save_checkpoint(
                            records,
                            completed_articles,
                        )

                        print()
                        print(
                            "ScienceDirect still blocked."
                        )

                        print(
                            "Current article remains "
                            "unfinished."
                        )

                        raise

                    except BrowserSessionLost as exc:

                        stop_reason = (
                            str(exc)
                        )

                        completed_articles.discard(
                            article_url
                        )

                        save_checkpoint(
                            records,
                            completed_articles,
                        )

                        raise

                    except Exception as exc:

                        print(
                            "ARTICLE ERROR: "
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
                            errors,
                        )

                        save_checkpoint(
                            records,
                            completed_articles,
                        )

                        continue

                    # ========================================
                    # ARTICLE SUCCESS
                    # ========================================

                    records = (
                        remove_article_records(
                            records,
                            article_url,
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
                                and
                                not (
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

                    # ----------------------------------------
                    # ONLY SUCCESS = COMPLETE
                    # ----------------------------------------

                    completed_articles.add(
                        article_url
                    )

                    successful_this_session += 1

                    print(
                        "    Article completed successfully."
                    )

                    save_checkpoint(
                        records,
                        completed_articles,
                    )

                    # ----------------------------------------
                    # PERIODIC EXCEL SAVE
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
                    # PERIODIC COOLDOWN
                    # ----------------------------------------

                    if (
                        successful_this_session
                        % COOLDOWN_AFTER_ARTICLES
                        == 0
                    ):

                        cooldown = (
                            random.uniform(
                                COOLDOWN_SECONDS_MIN,
                                COOLDOWN_SECONDS_MAX,
                            )
                        )

                        print()
                        print(
                            "Processed "
                            f"{successful_this_session} "
                            "articles this session."
                        )

                        print(
                            "Cooling down for "
                            f"{cooldown:.1f}s..."
                        )

                        save_checkpoint(
                            records,
                            completed_articles,
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
                            "Before next article",
                        )

    except KeyboardInterrupt:

        stop_reason = (
            "Stopped manually."
        )

        print()
        print(
            stop_reason
        )

    except ScienceDirectBlocked as exc:

        if not stop_reason:

            stop_reason = (
                str(exc)
            )

    except BrowserSessionLost as exc:

        stop_reason = (
            "Browser session lost: "
            f"{exc}"
        )

    except Exception as exc:

        stop_reason = (
            "Fatal error: "
            f"{exc}"
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
            completed_articles,
        )

        save_json(
            ERROR_FILE,
            errors,
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
                    "",
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
            "Completed articles : "
            f"{len(completed_articles):,}"
        )

        print(
            "Author records     : "
            f"{len(records):,}"
        )

        print(
            "Emails captured    : "
            f"{email_count:,}"
        )

        print(
            "Errors             : "
            f"{len(errors):,}"
        )

        if stop_reason:

            print(
                "Stop reason        : "
                f"{stop_reason}"
            )

        print(
            "Excel              : "
            f"{os.path.abspath(OUTPUT_FILE)}"
        )

        print(
            "Checkpoint         : "
            f"{os.path.abspath(CHECKPOINT_FILE)}"
        )

        print(
            "Errors JSON        : "
            f"{os.path.abspath(ERROR_FILE)}"
        )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    main()