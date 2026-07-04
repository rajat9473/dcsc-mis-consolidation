"""
BZ Transfer/COSYS workbook adapter.

BZ uses a different workbook structure from the standard airline
annexure files. AWB-level fields are read from Export TP_SC.
"""

import pandas as pd

from config import NUMERIC_COLUMNS, TARGET_COLUMNS

from utils import (
    clean_text,
    extract_airline,
    extract_direction,
    extract_fortnight,
    extract_month,
    safe_date,
    safe_numeric,
    source_file_name,
)


BZ_SHEET_NAME = "Export TP_SC"


BZ_COLUMN_MAPPING = {
    "FLIGHT No.": "Flight No",
    "FLIGHT NO.": "Flight No",
    "FLIGHT NO": "Flight No",
    "FLIGHT DATE": "Flight Date",
    "AWB No": "AWB No",
    "AWB NO": "AWB No",
    "SFX": "Sfx",
    "ORG": "Origin",
    "DES": "Dest",
    "PCS": "Pcs",
    "Gross wgt": "Gross Wt",
    "GROSS WGT": "Gross Wt",
    "Received_Chg_Wgt": "Chg Wt",
    "RECEIVED_CHG_WGT": "Chg Wt",
    "Billing SHC": "Billing SHC",
    "BILLING SHC": "Billing SHC",
}


def normalise_bz_header(value):
    """Normalise a BZ header for case-insensitive matching."""
    return clean_text(value).upper()


BZ_LOOKUP = {
    normalise_bz_header(source): target
    for source, target in BZ_COLUMN_MAPPING.items()
}


def parse_bz_workbook(file_path):
    """
    Parse the BZ Transfer/COSYS workbook.

    The Export TP_SC sheet provides one clean row per AWB and
    contains the flight, AWB, weight and billing SHC fields needed
    by the consolidated MIS.
    """
    excel_file = pd.ExcelFile(file_path)

    matching_sheet = None

    for sheet_name in excel_file.sheet_names:
        if clean_text(sheet_name).upper() == BZ_SHEET_NAME.upper():
            matching_sheet = sheet_name
            break

    if matching_sheet is None:
        return pd.DataFrame(columns=TARGET_COLUMNS)

    raw_df = pd.read_excel(
        file_path,
        sheet_name=matching_sheet,
        header=0,
    )

    detail_df = pd.DataFrame(index=raw_df.index)

    used_targets = set()

    for column in raw_df.columns:
        normalised_column = normalise_bz_header(column)
        target_column = BZ_LOOKUP.get(normalised_column)

        if (
            target_column
            and target_column not in used_targets
        ):
            detail_df[target_column] = raw_df[column]
            used_targets.add(target_column)

    for column in TARGET_COLUMNS:
        if column not in detail_df.columns:
            detail_df[column] = ""

    detail_df = detail_df[TARGET_COLUMNS]

    detail_df["Direction"] = extract_direction(file_path)
    detail_df["Airline"] = extract_airline(file_path)
    detail_df["Month"] = extract_month(file_path)
    detail_df["Fortnight"] = extract_fortnight(file_path)
    detail_df["Source File"] = source_file_name(file_path)

    detail_df["AWB No"] = detail_df["AWB No"].apply(clean_text)
    detail_df["Sfx"] = detail_df["Sfx"].apply(clean_text)

    detail_df = detail_df[
        detail_df["AWB No"].str.strip() != ""
    ].copy()

    detail_df = detail_df[
        ~detail_df["AWB No"].str.upper().isin(
            ["TOTAL", "GRAND TOTAL", "SUBTOTAL"]
        )
    ].copy()

    for column in NUMERIC_COLUMNS:
        detail_df[column] = detail_df[column].apply(safe_numeric)

    detail_df["Flight Date"] = detail_df["Flight Date"].apply(safe_date)

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


def read_bz_summary(file_path):
    """
    Read BZ service totals from the SUMMRY sheet.

    Returns a dictionary containing service element totals and
    the workbook grand total.
    """
    excel_file = pd.ExcelFile(file_path)

    summary_sheet = None

    for sheet_name in excel_file.sheet_names:
        if clean_text(sheet_name).upper() in {"SUMMRY", "SUMMARY"}:
            summary_sheet = sheet_name
            break

    if summary_sheet is None:
        return {
            "service_totals": {},
            "grand_total": 0.0,
        }

    raw_df = pd.read_excel(
        file_path,
        sheet_name=summary_sheet,
        header=None,
    )

    service_totals = {}
    grand_total = 0.0

    for _, row in raw_df.iterrows():
        values = row.tolist()

        if len(values) < 4:
            continue

        service_name = clean_text(values[1]).upper()
        amount = safe_numeric(values[3])

        if not service_name:
            continue

        if service_name == "TOTAL":
            grand_total = amount
            continue

        if amount:
            service_totals[service_name] = amount

    return {
        "service_totals": service_totals,
        "grand_total": grand_total,
    }