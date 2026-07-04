"""
Core workbook discovery and parsing orchestration.

The parser recursively scans Export and Import directories,
classifies source/reference workbooks, selects the appropriate
airline adapter and combines AWB-level rows into one common schema.
"""

import logging
from pathlib import Path

import pandas as pd

from adapters.bz_adapter import parse_bz_workbook
from adapters.standard import parse_standard_workbook
from config import LOG_FILE_NAME, TARGET_COLUMNS, UNIQUE_KEY
from utils import extract_month


LOGGER = logging.getLogger("dcsc_mis")


def setup_logging(base_dir):
    """Configure console and file logging."""
    log_dir = Path(base_dir) / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / LOG_FILE_NAME

    LOGGER.setLevel(logging.INFO)

    if LOGGER.handlers:
        LOGGER.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )

    file_handler = logging.FileHandler(
        log_file,
        mode="w",
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    LOGGER.addHandler(file_handler)
    LOGGER.addHandler(console_handler)

    return LOGGER


def discover_workbooks(base_dir):
    """
    Recursively discover every .xlsx file below Export and Import.

    Excel temporary lock files beginning with ~$ are ignored.
    """
    base_path = Path(base_dir)

    discovered_files = []

    for direction_folder in ["Export", "Import"]:
        root = base_path / direction_folder

        if not root.exists():
            LOGGER.warning(
                "Direction folder not found: %s",
                root,
            )
            continue

        for file_path in root.rglob("*.xlsx"):
            if file_path.name.startswith("~$"):
                continue

            discovered_files.append(file_path)

    return sorted(
        discovered_files,
        key=lambda path: str(path).lower(),
    )


def get_airline_from_path(file_path):
    """Read airline code from the parent folder."""
    return Path(file_path).parent.name.upper().strip()


def is_reference_workbook(file_path):
    """
    Detect a master/reference workbook that should be discovered
    but must not be counted again as a monthly billing source.

    SQ's rich Export Invoice Format workbook contains Data_Insert
    and Handling_SC master structures but has no billing month in
    its filename. Its AWBs overlap the fortnight source files.
    """
    if extract_month(file_path):
        return False

    try:
        excel_file = pd.ExcelFile(file_path)

        sheet_names = {
            str(sheet_name).strip().upper()
            for sheet_name in excel_file.sheet_names
        }

        master_markers = {
            "DATA_INSERT",
            "HANDLING_SC",
        }

        if master_markers.issubset(sheet_names):
            return True

    except Exception:
        return False

    return False


def parse_workbook(file_path):
    """
    Select the correct adapter based on airline code.
    """
    airline = get_airline_from_path(file_path)

    if airline == "BZ":
        LOGGER.info(
            "Using BZ adapter: %s",
            file_path.name,
        )

        return parse_bz_workbook(file_path)

    LOGGER.info(
        "Using standard adapter: %s",
        file_path.name,
    )

    return parse_standard_workbook(file_path)


def find_duplicate_keys(dataframe):
    """
    Return all rows involved in duplicate assignment keys.
    """
    if dataframe.empty:
        return pd.DataFrame(columns=TARGET_COLUMNS)

    duplicate_mask = dataframe.duplicated(
        subset=UNIQUE_KEY,
        keep=False,
    )

    return dataframe.loc[duplicate_mask].copy()


def parse_all_workbooks(base_dir):
    """
    Discover, classify and parse all source workbooks.

    A single malformed workbook is logged and does not stop the
    complete consolidation run.
    """
    setup_logging(base_dir)

    files = discover_workbooks(base_dir)

    LOGGER.info(
        "Discovered %s workbook(s).",
        len(files),
    )

    parsed_frames = []
    failed_files = []
    empty_files = []
    skipped_reference_files = []

    for file_path in files:
        try:
            LOGGER.info(
                "Inspecting: %s",
                file_path,
            )

            if is_reference_workbook(file_path):
                LOGGER.info(
                    "Skipped reference/master workbook: %s",
                    file_path.name,
                )

                skipped_reference_files.append(
                    str(file_path)
                )

                continue

            LOGGER.info(
                "Parsing source workbook: %s",
                file_path.name,
            )

            dataframe = parse_workbook(file_path)

            if dataframe.empty:
                LOGGER.warning(
                    "No AWB detail rows found: %s",
                    file_path.name,
                )

                empty_files.append(str(file_path))

                continue

            parsed_frames.append(dataframe)

            LOGGER.info(
                "Parsed %s AWB row(s): %s",
                len(dataframe),
                file_path.name,
            )

        except Exception as exc:
            LOGGER.exception(
                "Failed to parse %s: %s",
                file_path.name,
                exc,
            )

            failed_files.append(
                {
                    "file": str(file_path),
                    "error": str(exc),
                }
            )

    if parsed_frames:
        consolidated_df = pd.concat(
            parsed_frames,
            ignore_index=True,
        )

        consolidated_df = consolidated_df[
            TARGET_COLUMNS
        ]

    else:
        consolidated_df = pd.DataFrame(
            columns=TARGET_COLUMNS
        )

    duplicate_df = find_duplicate_keys(
        consolidated_df
    )

    if not duplicate_df.empty:
        LOGGER.warning(
            "Found %s row(s) involved in duplicate uniqueness keys.",
            len(duplicate_df),
        )

    else:
        LOGGER.info(
            "No duplicate uniqueness keys found."
        )

    LOGGER.info(
        "Source workbooks parsed: %s",
        len(parsed_frames),
    )

    LOGGER.info(
        "Reference workbooks skipped: %s",
        len(skipped_reference_files),
    )

    LOGGER.info(
        "Failed workbooks: %s",
        len(failed_files),
    )

    LOGGER.info(
        "Total consolidated AWB rows: %s",
        len(consolidated_df),
    )

    return {
        "data": consolidated_df,
        "files": files,
        "failed_files": failed_files,
        "empty_files": empty_files,
        "skipped_reference_files": skipped_reference_files,
        "duplicates": duplicate_df,
    }