# DCSC MIS Consolidation — Coding Assignment

## Background

DCSC operates an air-cargo terminal and raises monthly invoices on each airline for the
cargo it handles. Every airline's billing arrives as its **own** Excel workbook with its
**own** layout, sheet names, and column order. Some airlines split a month into two
**fortnight** files (01–15 and 16–end); others send one file per month; a few keep
fortnight data as separate **tabs inside one workbook**; and the transfer/COSYS airlines
use an entirely different sheet structure.

Today these are consolidated by hand. We want a tool that ingests all of them and produces
**one consolidated MIS**, separated by **direction (Import / Export)** and by **month**.

> All data in this pack is **synthetic dummy data** generated for the test. No real client
> data is included — flight numbers, AWB numbers, rates, shippers, agents and amounts are
> all fictional and do not match any actual values.

---

## Your task

Write a program (Python preferred; state your choice) that:

1. **Scans** the `Export/` and `Import/` folders recursively and discovers every `.xlsx`.
2. **Parses** each workbook, locating the AWB-level detail sheet regardless of its name
   (`ANNEXURE`, `ANX`, `Summary`, `Export TP Data`, in-file `ANNEX 01-15` tabs, etc.) and
   normalising the differing column headers to a common schema.
3. **Derives** the metadata that isn't always in the rows: `Direction` (from the top folder),
   `Airline` (from the folder / filename code), `Month` and `Fortnight` (from the filename,
   e.g. `FEB'2026 (01-15)`), and `Source File`.
4. **Consolidates** everything into a single output workbook matching
   **`MIS_Consolidated_TEMPLATE.xlsx`** — one `Export` sheet, one `Import` sheet
   (both at AWB-line level), and a `Month_Summary` sheet with totals by
   Airline × Month × Direction.
5. **Reconciles**: the summed charge columns per source file should foot back to that file's
   `Service Challan` / `SC` / `TOTAL SC` total. Flag any file where it doesn't.

---

## Input files in this pack

| Direction | Airline | Format quirk to handle |
|-----------|---------|------------------------|
| Export | **AY** | Separate fortnight files; two months (Jan + Feb) → tests month-wise merge |
| Export | **8M** | One workbook, fortnight kept as **separate tabs** (`ANNEX 01-15`, `ANNEX 16-28`) |
| Export | **SZ** | Single monthly file, minimal sheets; two months (Jan + Feb) |
| Export | **SQ** | Fortnight files **plus** a rich master (`Export Invoice Format` → `Data_Insert`) |
| Export | **BZ** | Transfer/COSYS layout — different sheet + column names entirely |
| Import | **AY, SQ** | Fortnight files, import lifecycle (arrival → segregation → DO → gate-out) |
| Import | **SZ** | Single monthly import file |

Note that **not every airline flies both directions** (8M and BZ are export-only here), and
the same airline appears in **both** Import and Export (AY, SQ, SZ) — so keys must include
`Direction`, not just airline.

## Common schema (target columns)

`Direction, Airline, Month, Fortnight, Source File, Flight No, Flight Date, AWB No, Sfx,`
`Origin, Dest, Pcs, Gross Wt, Chg Wt, Billing SHC, Handling Amt, X-Ray Amt, Demurrage Amt,`
`Total Amt (Incl GST)`

Uniqueness key: `(Direction, Airline, Month, Fortnight, AWB No, Sfx)`.

---

## Edge cases we're deliberately testing

- Header rows are **not always row 1** (title/date/carrier rows sit above them).
- Column **order and names differ** per airline for the same concept
  (`CHG WGT` vs `CHG WT.` vs `Received_Chg_Wgt`; `AWB No` vs `MAWB No` vs `AWB NO`).
- Import uses **MAWB + HAWB** and an arrival→gate-out timeline instead of carting/RCS.
- Some sheets carry trailing blank/`TOTAL` rows that must be excluded from line items.
- Fortnight can come from a **filename** (AY, SQ) or from a **tab name** (8M).
- Amounts are GST-inclusive at 18%; demurrage is only billed when days > 0.

---

## Deliverables

1. Source code + a short `README` explaining how to run it.
2. The generated consolidated workbook.
3. A brief note (½ page) on how you'd extend this to the remaining ~11 airline formats
   (config-driven mapping, per-airline adapters, etc.) and how you'd handle a new airline.

## Evaluation

- **Correctness** of parsing/normalisation across all five formats and both directions.
- **Robustness** to the edge cases above (does it crash on the BZ layout? mis-key AY import?).
- **Reconciliation** logic and how clearly mismatches are surfaced.
- **Code quality**: readability, config-driven vs hard-coded, ease of adding a new airline.
- **Bonus**: month-wise vs fortnight-wise toggle; a simple run log; basic tests.

## Suggested time

Half a day to a day. Focus on a clean, extensible core over covering every last sheet.
