"""
Standard AWB detail adapter.

Handles AY, SQ, SZ and 8M workbook layouts.
"""

import pandas as pd

from config import (
    HEADER_SCAN_LIMIT,
    NUMERIC_COLUMNS,
    TARGET_COLUMNS,
)

from utils import (
    clean_text,
    extract_airline,
    extract_direction,
    extract_fortnight,
    extract_month,
    is_invalid_line_item,
    map_column_name,
    safe_date,
    safe_numeric,
    source_file_name,
)


DETAIL_SHEET_NAMES = {
    "ANNEXURE",
    "ANX",
    "SUMMARY",
    "DATA_INSERT",
}


def is_detail_sheet(sheet_name):
    """Return True when a sheet is likely to contain AWB-level data."""
    name = clean_text(sheet_name).upper()

    if name in DETAIL_SHEET_NAMES:
        return True

    if name.startswith("ANNEX "):
        return True

    return False


def find_header_row(file_path, sheet_name):
    """
    Dynamically locate the header row by scanning the first rows.

    A valid AWB detail header should map at least four target
    columns and must contain an AWB column.
    """
    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=HEADER_SCAN_LIMIT,
    )

    best_row = None
    best_score = 0

    for row_index, row in preview.iterrows():
        mapped_columns = []

        for value in row.tolist():
            mapped = map_column_name(value)

            if mapped:
                mapped_columns.append(mapped)

        score = len(set(mapped_columns))

        if "AWB No" in mapped_columns and score > best_score:
            best_row = row_index
            best_score = score

    if best_row is None or best_score < 4:
        return None

    return best_row


def read_detail_sheet(file_path, sheet_name):
    """Read and normalise one AWB-level detail sheet."""
    header_row = find_header_row(file_path, sheet_name)

    if header_row is None:
        return pd.DataFrame(columns=TARGET_COLUMNS)

    raw_df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header_row,
    )

    detail_df = pd.DataFrame(index=raw_df.index)

    used_targets = set()

    # Preserve exact target columns first.
    for column in raw_df.columns:
        if column in TARGET_COLUMNS and column not in used_targets:
            detail_df[column] = raw_df[column]
            used_targets.add(column)

    # Map airline-specific aliases only for missing targets.
    for column in raw_df.columns:
        target_column = map_column_name(column)

        if (
            target_column
            and target_column not in used_targets
        ):
            detail_df[target_column] = raw_df[column]
            used_targets.add(target_column)

    # Create missing target columns.
    for column in TARGET_COLUMNS:
        if column not in detail_df.columns:
            detail_df[column] = ""

    # Enforce exact output order.
    detail_df = detail_df[TARGET_COLUMNS]

    # Derive metadata.
    direction = extract_direction(file_path)
    airline = extract_airline(file_path)
    month = extract_month(file_path)
    fortnight = extract_fortnight(file_path, sheet_name)

    detail_df["Direction"] = direction
    detail_df["Airline"] = airline
    detail_df["Month"] = month
    detail_df["Fortnight"] = fortnight
    detail_df["Source File"] = source_file_name(file_path)

    # Remove blank and total rows.
    detail_df = detail_df[
        ~detail_df.apply(is_invalid_line_item, axis=1)
    ].copy()

    # Clean AWB fields.
    detail_df["AWB No"] = detail_df["AWB No"].apply(clean_text)
    detail_df["Sfx"] = detail_df["Sfx"].apply(clean_text)

    # AWB number is mandatory for an AWB-level line.
    detail_df = detail_df[
        detail_df["AWB No"].str.strip() != ""
    ].copy()

    # Convert numeric fields.
    for column in NUMERIC_COLUMNS:
        detail_df[column] = detail_df[column].apply(safe_numeric)

    # Convert flight date.
    detail_df["Flight Date"] = detail_df["Flight Date"].apply(safe_date)

    # Clean textual fields.
    text_columns = [
        "Flight No",
        "AWB No",
        "Sfx",
        "Origin",
        "Dest",
        "Billing SHC",
    ]

    for column in text_columns:
        detail_df[column] = detail_df[column].apply(clean_text)

    return detail_df.reset_index(drop=True)


def parse_standard_workbook(file_path):
    """
    Parse all supported AWB-level sheets from one workbook.
    """
    excel_file = pd.ExcelFile(file_path)

    parsed_frames = []

    for sheet_name in excel_file.sheet_names:
        if not is_detail_sheet(sheet_name):
            continue

        detail_df = read_detail_sheet(
            file_path=file_path,
            sheet_name=sheet_name,
        )

        if not detail_df.empty:
            parsed_frames.append(detail_df)

    if not parsed_frames:
        return pd.DataFrame(columns=TARGET_COLUMNS)

    combined_df = pd.concat(
        parsed_frames,
        ignore_index=True,
    )

    return combined_df[TARGET_COLUMNS]