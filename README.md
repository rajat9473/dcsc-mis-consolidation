# DCSC MIS Consolidation

A Python-based Excel consolidation tool for normalising heterogeneous airline billing workbooks into a single DCSC MIS workbook.

## Overview

DCSC receives monthly airline billing data in multiple Excel formats. Different airlines use different sheet names, column headers, column orders, and billing structures.

This solution recursively scans the `Export/` and `Import/` folders, identifies AWB-level detail data, normalises it into a common schema, derives metadata, consolidates the records, and reconciles the generated totals against source Service Challan totals.

## Tech Stack

- Python 3
- pandas
- openpyxl
- pytest

## Features

- Recursively discovers `.xlsx` files under `Export/` and `Import/`.
- Derives Direction from the top-level folder.
- Derives Airline from the airline folder.
- Extracts Month and Fortnight from filenames.
- Supports fortnight metadata stored in workbook sheet names.
- Dynamically detects header rows below title and carrier rows.
- Normalises airline-specific column headers into a common schema.
- Handles AY, SQ, SZ, and 8M using the standard adapter.
- Handles the BZ Transfer/COSYS layout using a dedicated adapter.
- Excludes blank, trailing TOTAL, and subtotal rows.
- Detects duplicate AWB records using the assignment uniqueness key.
- Classifies the SQ rich master workbook as a reference workbook to prevent duplicate billing records.
- Enriches AWB records with GST-inclusive charge amounts.
- Reconciles consolidated totals against Service Challan / SC totals.
- Generates Export, Import, Month_Summary, and Reconciliation sheets.
- Writes execution logs for audit and debugging.
- Includes basic automated tests using pytest.

## Project Structure

```text
DCSC_MIS_Sample_Pack/
├── Export/
├── Import/
├── adapters/
│   ├── __init__.py
│   ├── standard.py
│   └── bz_adapter.py
├── tests/
│   └── test_parser.py
├── output/
│   └── MIS_Consolidated_Output.xlsx
├── logs/
├── config.py
├── utils.py
├── parser.py
├── enricher.py
├── reconciler.py
├── writer.py
├── main.py
├── requirements.txt
├── README.md
├── EXTENSION_NOTE.md
└── MIS_Consolidated_TEMPLATE.xlsx
```

## Setup

Clone the repository and move into the project directory.

Create a Python virtual environment:

```bash
python -m venv venv
```

Activate the virtual environment on Windows PowerShell:

```powershell
.\venv\Scripts\Activate.ps1
```

Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Run the Consolidation

Run the following command from the project root:

```bash
python main.py
```

The program will recursively scan the input folders, parse the source workbooks, enrich charge data, run reconciliation checks, and generate the consolidated MIS workbook.

## Generated Output

The final workbook is generated at:

```text
output/MIS_Consolidated_Output.xlsx
```

The execution log is written to:

```text
logs/dcsc_mis_run.log
```

## Output Workbook Sheets

### Export

Contains consolidated AWB-level Export records.

### Import

Contains consolidated AWB-level Import records.

### Month_Summary

Contains totals grouped by:

- Direction
- Airline
- Month

The summary includes AWB line count, pieces, gross weight, chargeable weight, and GST-inclusive charge totals.

### Reconciliation

Contains source-level reconciliation information including:

- Direction
- Airline
- Month
- Fortnight
- Source File
- Parsed Total
- Source SC Total
- Difference
- Status

A source is marked `MATCH` when the absolute difference is within the configured reconciliation tolerance.

## Common Schema

The consolidated AWB-level data uses the following target columns:

1. Direction
2. Airline
3. Month
4. Fortnight
5. Source File
6. Flight No
7. Flight Date
8. AWB No
9. Sfx
10. Origin
11. Dest
12. Pcs
13. Gross Wt
14. Chg Wt
15. Billing SHC
16. Handling Amt (Incl GST)
17. X-Ray Amt (Incl GST)
18. Demurrage Amt (Incl GST)
19. Total Amt (Incl GST)

## Uniqueness Key

Duplicate detection uses the following key:

```text
(Direction, Airline, Month, Fortnight, AWB No, Sfx)
```

Direction is included because the same airline may appear in both Import and Export data.

## Parsing Architecture

The solution separates workbook processing into multiple layers.

`parser.py` handles recursive workbook discovery, source classification, adapter selection, and consolidation.

`adapters/standard.py` handles structurally similar airline workbooks and normalises different header names.

`adapters/bz_adapter.py` handles the BZ Transfer/COSYS workbook structure.

`enricher.py` handles GST-inclusive AWB-level charge enrichment.

`reconciler.py` compares consolidated totals with authoritative Service Challan totals.

`writer.py` generates the final consolidated Excel workbook.

This separation keeps the consolidation core extensible and allows additional airline formats to be added with minimal changes.

## Charge Enrichment Assumption

Where a source workbook provides a direct AWB-level GST-inclusive service amount, the amount is mapped directly to the corresponding AWB.

Some source formats expose an authoritative aggregate Service Challan total without a complete AWB-level breakdown for every service component.

For the remaining unallocated amount, the solution allocates the residual proportionally using chargeable weight.

This assumption is explicitly isolated in `enricher.py`.

The final AWB totals are reconciled against the authoritative source Service Challan total, and reconciliation differences are surfaced in the `Reconciliation` sheet.

## Reconciliation Support

The reconciliation layer supports:

- Standard Service Challan sheets
- SC sheets
- TOTAL SC sheets
- 8M fortnight-specific SC tabs
- BZ SUMMRY grand total

The supplied sample pack produces 16 reconciliation checks.

## Tests

Run the automated tests using:

```bash
python -m pytest -q
```

The tests validate:

- Metadata extraction
- Recursive workbook discovery
- Consolidated row count
- Export and Import row counts
- Source reconciliation status

## Sample Pack Results

For the supplied synthetic sample pack, the program produces:

```text
Discovered workbooks : 16
Reference skipped    : 1
Failed workbooks     : 0
Duplicate key rows   : 0
Export AWB rows      : 861
Import AWB rows      : 383
Month summary rows   : 10
Reconciliation MATCH : 16
Reconciliation FAIL  : 0
Consolidated amount  : 10,741,760.4394
```

All 16 source reconciliation checks match within the configured tolerance.

## Extending to New Airlines

Airlines with sheet-name or column-name differences can be supported through configuration-driven header aliases and sheet detection rules.

Structurally different airline workbooks can be implemented using dedicated adapters following the same pattern as the BZ adapter.

For more details, see `EXTENSION_NOTE.md`.

## Author

Rajat Chitransh