"""
Utility helpers for metadata extraction, header normalisation
and safe value conversion.
"""

import re
from pathlib import Path

import pandas as pd

from config import AIRLINE_CODES, COLUMN_ALIASES, INVALID_ROW_KEYWORDS


def clean_text(value):
    """Convert a cell value to a clean string."""
    if pd.isna(value):
        return ""

    return str(value).strip()


def normalise_text(value):
    """
    Normalise text for case-insensitive header comparison.
    Underscores, dots and repeated spaces are handled.
    """
    text = clean_text(value).upper()
    text = text.replace("_", " ")
    text = text.replace(".", "")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def build_alias_lookup():
    """Create alias -> target column lookup."""
    lookup = {}

    for target_column, aliases in COLUMN_ALIASES.items():
        lookup[normalise_text(target_column)] = target_column

        for alias in aliases:
            lookup[normalise_text(alias)] = target_column

    return lookup


ALIAS_LOOKUP = build_alias_lookup()


def map_column_name(column_name):
    """Map an airline-specific column header to target schema."""
    normalised = normalise_text(column_name)

    return ALIAS_LOOKUP.get(normalised)


def extract_direction(file_path):
    """
    Derive Import or Export from the top-level folder.
    """
    path = Path(file_path)

    for part in path.parts:
        value = part.lower()

        if value == "export":
            return "Export"

        if value == "import":
            return "Import"

    return ""


def extract_airline(file_path):
    """
    Derive airline code from folder names first,
    then fall back to the filename.
    """
    path = Path(file_path)

    for part in reversed(path.parts[:-1]):
        code = part.upper().strip()

        if code in AIRLINE_CODES:
            return AIRLINE_CODES[code]

    filename = path.stem.upper()

    for code in AIRLINE_CODES:
        pattern = rf"(^|[^A-Z0-9]){re.escape(code)}([^A-Z0-9]|$)"

        if re.search(pattern, filename):
            return AIRLINE_CODES[code]

    return ""


def extract_month(file_path):
    """
    Extract month and year from filenames such as:
    FEB'2026
    JAN'2026
    FEB 2026
    FEB2026

    Returns a consistent value such as FEB-2026.
    """
    filename = Path(file_path).stem.upper()

    month_pattern = (
        r"\b("
        r"JAN(?:UARY)?|"
        r"FEB(?:RUARY)?|"
        r"MAR(?:CH)?|"
        r"APR(?:IL)?|"
        r"MAY|"
        r"JUN(?:E)?|"
        r"JUL(?:Y)?|"
        r"AUG(?:UST)?|"
        r"SEP(?:TEMBER)?|"
        r"OCT(?:OBER)?|"
        r"NOV(?:EMBER)?|"
        r"DEC(?:EMBER)?"
        r")"
        r"[\s'_-]*(20\d{2})\b"
    )

    match = re.search(month_pattern, filename)

    if not match:
        return ""

    month_name = match.group(1)[:3]
    year = match.group(2)

    return f"{month_name}-{year}"


def extract_fortnight(file_path, sheet_name=""):
    """
    Derive fortnight from filename or sheet name.

    Examples:
    (01-15)
    (16-28)
    ANNEX 01-15
    ANNEX 16-28
    """
    search_values = [
        Path(file_path).stem.upper(),
        str(sheet_name).upper(),
    ]

    for value in search_values:
        match = re.search(
            r"(?:\(|\b)(0?1|16)\s*[-–]\s*(15|28|29|30|31)(?:\)|\b)",
            value,
        )

        if match:
            start = int(match.group(1))
            end = int(match.group(2))

            return f"{start:02d}-{end:02d}"

    return "Monthly"


def is_invalid_line_item(row):
    """
    Detect trailing TOTAL / SUBTOTAL rows.
    """
    values = [
        normalise_text(value)
        for value in row.tolist()
        if clean_text(value)
    ]

    if not values:
        return True

    for value in values:
        for keyword in INVALID_ROW_KEYWORDS:
            if value == normalise_text(keyword):
                return True

    return False


def safe_numeric(value):
    """
    Convert Excel values safely to numbers.

    Handles:
    1,250.50
    ₹500
    blanks
    NaN
    """
    if pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = clean_text(value)

    if not text:
        return 0.0

    text = text.replace(",", "")
    text = text.replace("₹", "")
    text = text.replace("INR", "")
    text = text.strip()

    try:
        return float(text)

    except (TypeError, ValueError):
        return 0.0


def safe_date(value):
    """Convert a value to a pandas datetime when possible."""
    if pd.isna(value) or clean_text(value) == "":
        return pd.NaT

    try:
        return pd.to_datetime(value, errors="coerce")

    except (TypeError, ValueError):
        return pd.NaT


def source_file_name(file_path):
    """Return only the workbook filename."""
    return Path(file_path).name