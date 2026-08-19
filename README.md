# Diehl VIN Excel VBA

This repository is the source of truth for the VBA used by the Diehl VIN workbook.

## Files

- `vba/Module3.bas` — main VIN manager logic, dashboard refresh/filter logic, movable dashboard anchors, quick lookup, and self-repair.
- `vba/Dashboard_Sheet.bas` — code-behind for the `Dashboard` worksheet.
- `vba/Quick_Lookup_Sheet.bas` — code-behind for the `QUICK LOOKUP` worksheet.
- `vba/ThisWorkbook.bas` — workbook open/startup events.

## Workflow

1. Edit the VBA files in this repo.
2. Copy/import the matching code into the Excel VBA project.
3. Compile with `Debug > Compile VBAProject`.
4. Run `ResetExcelState` if Excel gets stuck with events/screen updating disabled.

## Current workbook assumptions

- Master sheet: `Vehicle Master`
- Dashboard sheet: `Dashboard`
- Quick lookup sheet: `QUICK LOOKUP`
- Header row: `2`
- First data row: `3`
- Open / Closed / All selector: `Dashboard!C5`

The dashboard uses named anchors so sections can be cut and moved without changing hard-coded VBA coordinates.
