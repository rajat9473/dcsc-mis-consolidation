"""
DCSC MIS Consolidation entry point.

Run:
    python main.py
"""

from pathlib import Path

from enricher import enrich_charges
from parser import parse_all_workbooks
from reconciler import reconcile_all
from writer import write_output_workbook


def main():
    """Run the complete MIS consolidation pipeline."""
    base_dir = Path(__file__).resolve().parent

    print("=" * 60)
    print("DCSC MIS CONSOLIDATION")
    print("=" * 60)

    parse_result = parse_all_workbooks(
        base_dir
    )

    consolidated_df = enrich_charges(
        consolidated_df=parse_result["data"],
        source_files=parse_result["files"],
        skipped_reference_files=parse_result[
            "skipped_reference_files"
        ],
    )

    reconciliation_df = reconcile_all(
        source_files=parse_result["files"],
        consolidated_df=consolidated_df,
        skipped_reference_files=parse_result[
            "skipped_reference_files"
        ],
    )

    output_result = write_output_workbook(
        base_dir=base_dir,
        consolidated_df=consolidated_df,
        reconciliation_df=reconciliation_df,
    )

    match_count = (
        reconciliation_df["Status"]
        .eq("MATCH")
        .sum()
    )

    mismatch_count = (
        reconciliation_df["Status"]
        .eq("MISMATCH")
        .sum()
    )

    total_amount = consolidated_df[
        "Total Amt (Incl GST)"
    ].sum()

    print()
    print("=" * 60)
    print("RUN SUMMARY")
    print("=" * 60)

    print(
        f"Discovered workbooks : "
        f"{len(parse_result['files'])}"
    )

    print(
        f"Reference skipped    : "
        f"{len(parse_result['skipped_reference_files'])}"
    )

    print(
        f"Failed workbooks     : "
        f"{len(parse_result['failed_files'])}"
    )

    print(
        f"Duplicate key rows   : "
        f"{len(parse_result['duplicates'])}"
    )

    print(
        f"Export AWB rows      : "
        f"{output_result['export_rows']}"
    )

    print(
        f"Import AWB rows      : "
        f"{output_result['import_rows']}"
    )

    print(
        f"Month summary rows   : "
        f"{output_result['summary_rows']}"
    )

    print(
        f"Reconciliation MATCH : "
        f"{match_count}"
    )

    print(
        f"Reconciliation FAIL  : "
        f"{mismatch_count}"
    )

    print(
        f"Consolidated amount  : "
        f"{total_amount:,.4f}"
    )

    print()
    print(
        "Output workbook:"
    )

    print(
        output_result["output_path"]
    )

    print("=" * 60)


if __name__ == "__main__":
    main()