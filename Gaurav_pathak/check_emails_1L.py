import os
import re
import sys
import html
import math
import unicodedata
from datetime import date, timedelta

import pandas as pd

try:
    from email_validator import validate_email, EmailNotValidError
except ImportError:
    print("Missing package: email-validator")
    print("Install using:")
    print("pip install pandas openpyxl email-validator dnspython")
    sys.exit(1)


# ============================================================
# CONFIGURATION
# ============================================================

INPUT_FILE = "input.xlsx"

OUTPUT_FOLDER = "cleaned_output"

EMAIL_COLUMN = "email"
NAME_COLUMN = "firstname"

# 6,000 VALID emails per output/day
VALID_EMAILS_PER_DAY = 6000

# True:
#   checks DNS / domain deliverability
#
# False:
#   syntax-only validation
CHECK_DELIVERABILITY = True

# Remove duplicate emails after validation
REMOVE_DUPLICATES = True

# Starting date for daily output files.
# Uses today's machine date.
START_DATE = date.today()


# ============================================================
# EMAIL CLEANING
# ============================================================

def clean_email(value):
    """
    Clean common formatting problems without changing
    the underlying mailbox unnecessarily.
    """

    if pd.isna(value):
        return ""

    email = str(value)

    # Decode HTML entities
    email = html.unescape(email)

    # Unicode normalization
    email = unicodedata.normalize("NFKC", email)

    # Remove invisible characters
    email = (
        email.replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
        .replace("\xa0", " ")
    )

    email = email.strip()

    # Remove mailto:
    email = re.sub(
        r"^\s*mailto\s*:\s*",
        "",
        email,
        flags=re.IGNORECASE
    )

    # Remove surrounding junk
    email = email.strip(
        " <>[](){}\"'`"
    )

    # Remove whitespace inside email
    email = re.sub(
        r"\s+",
        "",
        email
    )

    # Fix escaped @
    email = email.replace(
        "\\@",
        "@"
    )

    # HTML encoded @
    email = email.replace(
        "&#64;",
        "@"
    )

    # Remove common trailing punctuation
    email = email.rstrip(
        ",;:"
    )

    return email.lower()


# ============================================================
# EMAIL VALIDATION
# ============================================================

def validate_email_address(email):
    """
    Returns:

        is_valid
        normalized_email

    Invalid reason is not returned because invalid records
    will not be stored in the final output.
    """

    if not email:
        return False, ""

    if "@" not in email:
        return False, ""

    if email.count("@") != 1:
        return False, ""

    local_part, domain = email.rsplit(
        "@",
        1
    )

    if not local_part:
        return False, ""

    if not domain:
        return False, ""

    if ".." in local_part:
        return False, ""

    if local_part.startswith("."):
        return False, ""

    if local_part.endswith("."):
        return False, ""

    try:
        result = validate_email(
            email,
            check_deliverability=CHECK_DELIVERABILITY
        )

        normalized_email = (
            result.normalized
            .strip()
            .lower()
        )

        return True, normalized_email

    except EmailNotValidError:
        return False, ""

    except Exception:
        return False, ""


# ============================================================
# NAME CLEANING
# ============================================================

def split_camel_case(text):
    """
    Examples:

        PradoCarla
            -> Prado Carla

        MacedoAriane
            -> Macedo Ariane

        DeArruda
            -> De Arruda

        HouJibo
            -> Hou Jibo
    """

    if not text:
        return ""

    # ABCJohn -> ABC John
    text = re.sub(
        r"([A-Z]+)([A-Z][a-z])",
        r"\1 \2",
        text
    )

    # PradoCarla -> Prado Carla
    text = re.sub(
        r"([a-zà-ÿ])([A-ZÀ-Ý])",
        r"\1 \2",
        text
    )

    return text


def clean_name(value):
    """
    Clean the provided firstname/name.

    Keeps:
        Unicode letters
        spaces
        hyphens
        apostrophes
        periods
    """

    if pd.isna(value):
        return ""

    name = str(value)

    # Decode HTML
    name = html.unescape(name)

    # Normalize Unicode
    name = unicodedata.normalize(
        "NFKC",
        name
    )

    # Invisible characters
    name = (
        name.replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
        .replace("\xa0", " ")
    )

    # Remove HTML tags
    name = re.sub(
        r"<[^>]+>",
        " ",
        name
    )

    # Remove URLs
    name = re.sub(
        r"https?://\S+|www\.\S+",
        " ",
        name,
        flags=re.IGNORECASE
    )

    # Remove email addresses from name
    name = re.sub(
        r"\b[A-Za-z0-9._%+\-]+"
        r"@[A-Za-z0-9.\-]+"
        r"\.[A-Za-z]{2,}\b",
        " ",
        name
    )

    # Replace separators
    name = re.sub(
        r"[_/\\|;,]+",
        " ",
        name
    )

    # Remove brackets
    name = (
        name.replace("(", " ")
        .replace(")", " ")
        .replace("[", " ")
        .replace("]", " ")
        .replace("{", " ")
        .replace("}", " ")
    )

    # Remove numbers
    name = re.sub(
        r"\d+",
        " ",
        name
    )

    # Split CamelCase
    name = split_camel_case(
        name
    )

    cleaned_chars = []

    for char in name:
        category = unicodedata.category(
            char
        )

        if (
            category.startswith("L")
            or char in [
                " ",
                "-",
                "'",
                "’",
                "."
            ]
        ):
            cleaned_chars.append(char)

        else:
            cleaned_chars.append(" ")

    name = "".join(
        cleaned_chars
    )

    # Standard apostrophe
    name = name.replace(
        "’",
        "'"
    )

    # Remove duplicate whitespace
    name = re.sub(
        r"\s+",
        " ",
        name
    ).strip()

    # Repeated punctuation
    name = re.sub(
        r"-{2,}",
        "-",
        name
    )

    name = re.sub(
        r"'{2,}",
        "'",
        name
    )

    name = name.strip(
        " .-'"
    )

    return name


# ============================================================
# GET NAME FROM EMAIL
# ============================================================

def name_from_email(email):
    """
    Generate a fallback name when the original firstname
    is empty.

    Examples:

        sarah.skromanis@utas.edu.au
            -> Sarah Skromanis

        barbara.bacova@savba.sk
            -> Barbara Bacova

    Avoids guessing IDs such as:
        plh
        xds3
        aa7885
    """

    if not email or "@" not in email:
        return ""

    local = email.split(
        "@",
        1
    )[0]

    # Do not guess usernames containing numbers
    if any(
        char.isdigit()
        for char in local
    ):
        return ""

    local = re.sub(
        r"[._\-]+",
        " ",
        local
    )

    local = re.sub(
        r"\s+",
        " ",
        local
    ).strip()

    if not local:
        return ""

    parts = local.split()

    # Don't guess very short account IDs
    if (
        len(parts) == 1
        and len(parts[0]) <= 3
    ):
        return ""

    return " ".join(
        word.capitalize()
        for word in parts
    )


# ============================================================
# PROCESS SINGLE ROW
# ============================================================

def process_row(row):
    """
    Returns None if the email is invalid.

    Returns:
        {
            email,
            name
        }

    when valid.
    """

    raw_email = row.get(
        EMAIL_COLUMN,
        ""
    )

    raw_name = row.get(
        NAME_COLUMN,
        ""
    )

    # --------------------------------------------------------
    # CLEAN EMAIL
    # --------------------------------------------------------

    email = clean_email(
        raw_email
    )

    # --------------------------------------------------------
    # VALIDATE EMAIL
    # --------------------------------------------------------

    is_valid, normalized_email = (
        validate_email_address(
            email
        )
    )

    # Completely discard invalid emails
    if not is_valid:
        return None

    email = normalized_email

    # --------------------------------------------------------
    # CLEAN NAME
    # --------------------------------------------------------

    name = clean_name(
        raw_name
    )

    # If original name is empty,
    # try to derive it from email
    if not name:
        name = name_from_email(
            email
        )

    return {
        "email": email,
        "name": name
    }


# ============================================================
# FORMAT EXCEL
# ============================================================

def format_excel(writer):
    workbook = writer.book

    for sheet_name in workbook.sheetnames:
        worksheet = workbook[
            sheet_name
        ]

        worksheet.freeze_panes = "A2"

        if (
            worksheet.max_row >= 1
            and worksheet.max_column >= 1
        ):
            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

        worksheet.column_dimensions[
            "A"
        ].width = 45

        worksheet.column_dimensions[
            "B"
        ].width = 35


# ============================================================
# SAVE ONE DAILY FILE
# ============================================================

def save_daily_file(
    daily_df,
    output_date,
    day_number
):
    """
    Saves only valid email + name records.
    """

    date_string = output_date.strftime(
        "%Y-%m-%d"
    )

    output_name = (
        f"valid_emails_{date_string}.xlsx"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        output_name
    )

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        daily_df.to_excel(
            writer,
            sheet_name="Valid Emails",
            index=False
        )

        format_excel(
            writer
        )

    return output_path


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 100)
    print("EMAIL + NAME CLEANER / VALIDATOR")
    print("6,000 VALID EMAILS PER DAY")
    print("INVALID EMAILS ARE DISCARDED")
    print("=" * 100)

    # --------------------------------------------------------
    # CHECK INPUT
    # --------------------------------------------------------

    if not os.path.exists(
        INPUT_FILE
    ):
        print()
        print(
            f"ERROR: Input file not found: "
            f"{INPUT_FILE}"
        )

        sys.exit(1)

    # --------------------------------------------------------
    # OUTPUT DIRECTORY
    # --------------------------------------------------------

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    print()
    print(
        f"Input file        : "
        f"{INPUT_FILE}"
    )

    print(
        f"Output folder     : "
        f"{OUTPUT_FOLDER}"
    )

    print(
        f"Valid/day         : "
        f"{VALID_EMAILS_PER_DAY:,}"
    )

    print(
        f"Start date        : "
        f"{START_DATE}"
    )

    print(
        "DNS validation    : "
        + (
            "Enabled"
            if CHECK_DELIVERABILITY
            else "Disabled"
        )
    )

    print(
        "Remove duplicates : "
        + (
            "Yes"
            if REMOVE_DUPLICATES
            else "No"
        )
    )

    # --------------------------------------------------------
    # READ EXCEL
    # --------------------------------------------------------

    try:
        df = pd.read_excel(
            INPUT_FILE,
            dtype=str
        )

    except Exception as exc:
        print()
        print(
            f"ERROR reading Excel: "
            f"{exc}"
        )

        sys.exit(1)

    # --------------------------------------------------------
    # NORMALIZE COLUMN NAMES
    # --------------------------------------------------------

    df.columns = [
        str(column)
        .strip()
        .lower()
        for column in df.columns
    ]

    print()
    print(
        "Columns:",
        list(df.columns)
    )

    # --------------------------------------------------------
    # REQUIRED COLUMNS
    # --------------------------------------------------------

    missing_columns = []

    if EMAIL_COLUMN not in df.columns:
        missing_columns.append(
            EMAIL_COLUMN
        )

    if NAME_COLUMN not in df.columns:
        missing_columns.append(
            NAME_COLUMN
        )

    if missing_columns:
        print()
        print(
            "ERROR: Missing required column(s):",
            ", ".join(
                missing_columns
            )
        )

        sys.exit(1)

    # --------------------------------------------------------
    # REMOVE COMPLETELY EMPTY ROWS
    # --------------------------------------------------------

    df = df[
        ~(
            df[EMAIL_COLUMN]
            .fillna("")
            .astype(str)
            .str.strip()
            .eq("")
            &
            df[NAME_COLUMN]
            .fillna("")
            .astype(str)
            .str.strip()
            .eq("")
        )
    ].copy()

    df.reset_index(
        drop=True,
        inplace=True
    )

    total_input_records = len(
        df
    )

    if total_input_records == 0:
        print()
        print(
            "No records found."
        )

        return

    print()
    print(
        f"Total input records: "
        f"{total_input_records:,}"
    )

    print()
    print(
        "Validating all records..."
    )
    print()

    # --------------------------------------------------------
    # PROCESS ALL RECORDS FIRST
    # --------------------------------------------------------

    valid_records = []

    invalid_count = 0

    for index, row in df.iterrows():

        result = process_row(
            row
        )

        if result is None:
            invalid_count += 1

        else:
            valid_records.append(
                result
            )

        processed_count = (
            index + 1
        )

        # Progress every 100 records
        if (
            processed_count % 100 == 0
            or processed_count
            == total_input_records
        ):
            print(
                f"\rProcessed: "
                f"{processed_count:,}/"
                f"{total_input_records:,} | "
                f"Valid: "
                f"{len(valid_records):,} | "
                f"Invalid: "
                f"{invalid_count:,}",
                end="",
                flush=True
            )

    print()
    print()

    # --------------------------------------------------------
    # CREATE VALID DATAFRAME
    # --------------------------------------------------------

    if not valid_records:
        print(
            "No valid email addresses were found."
        )

        return

    valid_df = pd.DataFrame(
        valid_records,
        columns=[
            "email",
            "name"
        ]
    )

    valid_before_dedupe = len(
        valid_df
    )

    # --------------------------------------------------------
    # REMOVE DUPLICATE EMAILS
    # --------------------------------------------------------

    duplicate_count = 0

    if REMOVE_DUPLICATES:

        valid_df = (
            valid_df
            .drop_duplicates(
                subset=["email"],
                keep="first"
            )
            .reset_index(
                drop=True
            )
        )

        duplicate_count = (
            valid_before_dedupe
            - len(valid_df)
        )

    total_valid = len(
        valid_df
    )

    # --------------------------------------------------------
    # CALCULATE DAILY FILES
    # --------------------------------------------------------

    total_days = math.ceil(
        total_valid
        / VALID_EMAILS_PER_DAY
    )

    print("=" * 100)
    print("VALIDATION COMPLETE")
    print("=" * 100)

    print(
        f"Input records       : "
        f"{total_input_records:,}"
    )

    print(
        f"Valid before dedupe : "
        f"{valid_before_dedupe:,}"
    )

    print(
        f"Invalid discarded   : "
        f"{invalid_count:,}"
    )

    print(
        f"Duplicates removed  : "
        f"{duplicate_count:,}"
    )

    print(
        f"Final valid emails  : "
        f"{total_valid:,}"
    )

    print(
        f"Emails per day      : "
        f"{VALID_EMAILS_PER_DAY:,}"
    )

    print(
        f"Total daily files   : "
        f"{total_days:,}"
    )

    print()

    # --------------------------------------------------------
    # SAVE 6,000 VALID EMAILS PER DAY
    # --------------------------------------------------------

    generated_files = []

    for day_index in range(
        total_days
    ):
        start_index = (
            day_index
            * VALID_EMAILS_PER_DAY
        )

        end_index = min(
            start_index
            + VALID_EMAILS_PER_DAY,
            total_valid
        )

        daily_df = (
            valid_df
            .iloc[
                start_index:end_index
            ]
            .copy()
            .reset_index(
                drop=True
            )
        )

        output_date = (
            START_DATE
            + timedelta(
                days=day_index
            )
        )

        output_path = save_daily_file(
            daily_df=daily_df,
            output_date=output_date,
            day_number=day_index + 1
        )

        generated_files.append(
            output_path
        )

        print(
            f"Day {day_index + 1:03d} | "
            f"{output_date} | "
            f"{len(daily_df):,} valid emails | "
            f"{output_path}"
        )

    # --------------------------------------------------------
    # FINAL SUMMARY
    # --------------------------------------------------------

    print()
    print("=" * 100)
    print("COMPLETED")
    print("=" * 100)

    print(
        f"Input records      : "
        f"{total_input_records:,}"
    )

    print(
        f"Invalid discarded  : "
        f"{invalid_count:,}"
    )

    print(
        f"Duplicates removed : "
        f"{duplicate_count:,}"
    )

    print(
        f"Valid exported     : "
        f"{total_valid:,}"
    )

    print(
        f"Excel files        : "
        f"{len(generated_files):,}"
    )

    print()
    print(
        "Generated daily files:"
    )

    for file_path in generated_files:
        print(
            f"  - "
            f"{os.path.abspath(file_path)}"
        )


if __name__ == "__main__":
    main()