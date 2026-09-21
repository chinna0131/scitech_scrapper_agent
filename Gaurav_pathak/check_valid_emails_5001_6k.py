import os
import re
import sys
import html
import math
import unicodedata
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

BATCH_SIZE = 500

# True  = validate domain / DNS / MX
# False = syntax-only validation
CHECK_DELIVERABILITY = True


# ============================================================
# EMAIL CLEANING
# ============================================================

def clean_email(value):
    if pd.isna(value):
        return ""

    email = str(value)

    email = html.unescape(email)
    email = unicodedata.normalize("NFKC", email)

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

    # Remove surrounding characters
    email = email.strip(" <>[](){}\"'`")

    # Remove spaces inside email
    email = re.sub(r"\s+", "", email)

    # Escaped @
    email = email.replace("\\@", "@")

    # HTML encoded @
    email = email.replace("&#64;", "@")

    # Trailing punctuation
    email = email.rstrip(",;:")

    return email.lower()


# ============================================================
# EMAIL VALIDATION
# ============================================================

def validate_email_address(email):

    if not email:
        return "No", "Email is empty", ""

    if "@" not in email:
        return "No", "Missing @ symbol", email

    if email.count("@") != 1:
        return "No", "Email contains multiple @ symbols", email

    local_part, domain = email.rsplit("@", 1)

    if not local_part:
        return "No", "Missing username before @", email

    if not domain:
        return "No", "Missing domain after @", email

    if ".." in local_part:
        return "No", "Username contains consecutive dots", email

    if local_part.startswith("."):
        return "No", "Username starts with a dot", email

    if local_part.endswith("."):
        return "No", "Username ends with a dot", email

    try:

        result = validate_email(
            email,
            check_deliverability=CHECK_DELIVERABILITY
        )

        normalized_email = result.normalized.lower()

        return "Yes", "", normalized_email

    except EmailNotValidError as exc:

        reason = str(exc)
        reason_lower = reason.lower()

        if "does not exist" in reason_lower:
            reason = "Email domain does not exist"

        elif "does not accept email" in reason_lower:
            reason = "Email domain does not accept email"

        elif "not deliverable" in reason_lower:
            reason = "Email domain is not deliverable"

        elif "must have an @" in reason_lower:
            reason = "Missing @ symbol"

        else:
            reason = f"Invalid email: {reason}"

        return "No", reason, email

    except Exception as exc:

        return (
            "No",
            f"Validation error: {str(exc)}",
            email
        )


# ============================================================
# NAME CLEANING
# ============================================================

def split_camel_case(text):

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

    if pd.isna(value):
        return ""

    name = str(value)

    name = html.unescape(name)
    name = unicodedata.normalize("NFKC", name)

    name = (
        name.replace("\u200b", "")
        .replace("\u200c", "")
        .replace("\u200d", "")
        .replace("\ufeff", "")
        .replace("\xa0", " ")
    )

    # HTML tags
    name = re.sub(
        r"<[^>]+>",
        " ",
        name
    )

    # URLs
    name = re.sub(
        r"https?://\S+|www\.\S+",
        " ",
        name,
        flags=re.IGNORECASE
    )

    # Emails inside name
    name = re.sub(
        r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b",
        " ",
        name
    )

    # Separators
    name = re.sub(
        r"[_/\\|;,]+",
        " ",
        name
    )

    # Brackets
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

    # CamelCase splitting
    name = split_camel_case(name)

    cleaned = []

    for char in name:

        category = unicodedata.category(char)

        if (
            category.startswith("L")
            or char in [" ", "-", "'", "’", "."]
        ):
            cleaned.append(char)

        else:
            cleaned.append(" ")

    name = "".join(cleaned)

    name = name.replace("’", "'")

    # Remove duplicate whitespace
    name = re.sub(
        r"\s+",
        " ",
        name
    ).strip()

    # Duplicate punctuation
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

    name = name.strip(" .-'")

    return name


# ============================================================
# CREATE NAME FROM EMAIL IF NAME EMPTY
# ============================================================

def name_from_email(email):

    if not email or "@" not in email:
        return ""

    local = email.split("@", 1)[0]

    # Don't use IDs containing digits
    if any(char.isdigit() for char in local):
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

    # Avoid short IDs like:
    # plh
    # abc
    if len(parts) == 1 and len(parts[0]) <= 3:
        return ""

    return " ".join(
        word.capitalize()
        for word in parts
    )


# ============================================================
# PROCESS SINGLE ROW
# ============================================================

def process_row(row):

    raw_email = row.get(
        EMAIL_COLUMN,
        ""
    )

    raw_name = row.get(
        NAME_COLUMN,
        ""
    )

    # ------------------------------------
    # EMAIL
    # ------------------------------------

    email = clean_email(raw_email)

    is_valid, reason, normalized_email = (
        validate_email_address(email)
    )

    if normalized_email:
        email = normalized_email

    # ------------------------------------
    # NAME
    # ------------------------------------

    name = clean_name(raw_name)

    if not name:
        generated_name = name_from_email(email)

        if generated_name:
            name = generated_name

    return pd.Series({
        "email": email,
        "name": name,
        "is_valid_email": is_valid,
        "reason": reason,
    })


# ============================================================
# FORMAT EXCEL
# ============================================================

def format_excel(writer):

    workbook = writer.book

    for sheet_name in workbook.sheetnames:

        worksheet = workbook[sheet_name]

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions

        worksheet.column_dimensions["A"].width = 45
        worksheet.column_dimensions["B"].width = 35
        worksheet.column_dimensions["C"].width = 18
        worksheet.column_dimensions["D"].width = 65


# ============================================================
# SAVE ONE BATCH
# ============================================================

def save_batch(batch_df, batch_number):

    output_name = (
        f"cleaned_validated_emails_part_"
        f"{batch_number:03d}.xlsx"
    )

    output_path = os.path.join(
        OUTPUT_FOLDER,
        output_name
    )

    valid_df = batch_df[
        batch_df["is_valid_email"] == "Yes"
    ].copy()

    invalid_df = batch_df[
        batch_df["is_valid_email"] == "No"
    ].copy()

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl"
    ) as writer:

        # All records
        batch_df.to_excel(
            writer,
            sheet_name="Cleaned Data",
            index=False
        )

        # Valid only
        valid_df.to_excel(
            writer,
            sheet_name="Valid Emails",
            index=False
        )

        # Invalid only
        invalid_df.to_excel(
            writer,
            sheet_name="Invalid Emails",
            index=False
        )

        format_excel(writer)

    return output_path


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 100)
    print("EMAIL + NAME CLEANER / VALIDATOR")
    print("500 RECORDS PER OUTPUT EXCEL")
    print("=" * 100)

    # --------------------------------------------------------
    # INPUT CHECK
    # --------------------------------------------------------

    if not os.path.exists(INPUT_FILE):

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
    print(f"Input file      : {INPUT_FILE}")
    print(f"Output folder   : {OUTPUT_FOLDER}")
    print(f"Batch size      : {BATCH_SIZE}")
    print(
        "DNS validation  : "
        + (
            "Enabled"
            if CHECK_DELIVERABILITY
            else "Disabled"
        )
    )

    # --------------------------------------------------------
    # READ INPUT
    # --------------------------------------------------------

    try:

        df = pd.read_excel(
            INPUT_FILE,
            dtype=str
        )

    except Exception as exc:

        print()
        print(
            f"ERROR reading input Excel: {exc}"
        )

        sys.exit(1)

    # Normalize headings
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
            "ERROR: Missing columns:",
            ", ".join(missing_columns)
        )

        sys.exit(1)

    # --------------------------------------------------------
    # REMOVE FULLY EMPTY ROWS
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

    total_records = len(df)

    if total_records == 0:

        print()
        print("No records found.")

        return

    total_batches = math.ceil(
        total_records / BATCH_SIZE
    )

    print()
    print(
        f"Total records    : "
        f"{total_records:,}"
    )

    print(
        f"Total Excel files: "
        f"{total_batches:,}"
    )

    print()

    # --------------------------------------------------------
    # PROCESS BATCHES
    # --------------------------------------------------------

    overall_valid = 0
    overall_invalid = 0

    generated_files = []

    for batch_number in range(
        1,
        total_batches + 1
    ):

        start_index = (
            (batch_number - 1)
            * BATCH_SIZE
        )

        end_index = min(
            start_index + BATCH_SIZE,
            total_records
        )

        raw_batch = df.iloc[
            start_index:end_index
        ].copy()

        print(
            f"[Batch {batch_number}/{total_batches}] "
            f"Processing rows "
            f"{start_index + 1:,} - {end_index:,}"
        )

        # Process records
        processed_batch = raw_batch.apply(
            process_row,
            axis=1
        )

        # Count validation results
        valid_count = (
            processed_batch[
                "is_valid_email"
            ]
            .eq("Yes")
            .sum()
        )

        invalid_count = (
            processed_batch[
                "is_valid_email"
            ]
            .eq("No")
            .sum()
        )

        overall_valid += valid_count
        overall_invalid += invalid_count

        # Save batch
        output_path = save_batch(
            processed_batch,
            batch_number
        )

        generated_files.append(
            output_path
        )

        print(
            f"    Records : "
            f"{len(processed_batch):,}"
        )

        print(
            f"    Valid   : "
            f"{valid_count:,}"
        )

        print(
            f"    Invalid : "
            f"{invalid_count:,}"
        )

        print(
            f"    Saved   : "
            f"{output_path}"
        )

        print()

    # --------------------------------------------------------
    # FINAL SUMMARY
    # --------------------------------------------------------

    print("=" * 100)
    print("COMPLETED")
    print("=" * 100)

    print(
        f"Total records : "
        f"{total_records:,}"
    )

    print(
        f"Valid emails  : "
        f"{overall_valid:,}"
    )

    print(
        f"Invalid emails: "
        f"{overall_invalid:,}"
    )

    print(
        f"Excel files   : "
        f"{len(generated_files):,}"
    )

    if total_records:

        valid_percentage = (
            overall_valid
            / total_records
        ) * 100

        print(
            f"Valid %       : "
            f"{valid_percentage:.2f}%"
        )

    print()
    print("Generated files:")

    for file_path in generated_files:
        print(
            f"  - {os.path.abspath(file_path)}"
        )


if __name__ == "__main__":
    main()