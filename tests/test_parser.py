"""Basic tests for DCSC MIS consolidation."""

import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[1]

# Make project root importable when pytest is executed.
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


from enricher import enrich_charges
from parser import (
    discover_workbooks,
    parse_all_workbooks,
)
from reconciler import reconcile_all
from utils import (
    extract_airline,
    extract_direction,
    extract_fortnight,
    extract_month,
)


def test_metadata_extraction():
    file_path = (
        BASE_DIR
        / "Export"
        / "AY"
        / "AY EXPORT FEB'2026 (01-15).xlsx"
    )

    assert extract_direction(file_path) == "Export"
    assert extract_airline(file_path) == "AY"
    assert extract_month(file_path) == "FEB-2026"
    assert extract_fortnight(file_path) == "01-15"


def test_recursive_workbook_discovery():
    files = discover_workbooks(BASE_DIR)

    assert len(files) == 16

    assert all(
        file_path.suffix.lower() == ".xlsx"
        for file_path in files
    )


def test_consolidated_row_count():
    result = parse_all_workbooks(BASE_DIR)

    assert len(result["data"]) == 1244
    assert len(result["failed_files"]) == 0
    assert len(result["duplicates"]) == 0

    assert len(
        result["skipped_reference_files"]
    ) == 1


def test_direction_row_counts():
    result = parse_all_workbooks(BASE_DIR)

    dataframe = result["data"]

    export_rows = (
        dataframe["Direction"]
        .eq("Export")
        .sum()
    )

    import_rows = (
        dataframe["Direction"]
        .eq("Import")
        .sum()
    )

    assert export_rows == 861
    assert import_rows == 383


def test_reconciliation_matches():
    result = parse_all_workbooks(BASE_DIR)

    enriched_df = enrich_charges(
        consolidated_df=result["data"],
        source_files=result["files"],
        skipped_reference_files=result[
            "skipped_reference_files"
        ],
    )

    reconciliation_df = reconcile_all(
        source_files=result["files"],
        consolidated_df=enriched_df,
        skipped_reference_files=result[
            "skipped_reference_files"
        ],
    )

    assert len(reconciliation_df) == 16

    assert (
        reconciliation_df["Status"]
        .eq("MATCH")
        .all()
    )