# Diehl VIN Excel VBA

This repository is the source of truth for the VBA used by the Diehl VIN workbook.

## Files

- `vba/Module3.bas` — main VIN manager logic, dashboard refresh/filter logic, movable dashboard anchors, quick lookup, and self-repair.
- `vba/Dashboard_Sheet.bas` — code-behind for the `Dashboard` worksheet.
- `vba/Quick_Lookup_Sheet.bas` — code-behind for the `QUICK LOOKUP` worksheet.
- `vba/ThisWorkbook.bas` — workbook open/startup events.

## Installing the VBA into Excel

1. Open the `.xlsm` workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Put the contents of `vba/Module3.bas` into the standard VBA module named `Module3`.
4. Put the contents of `vba/Dashboard_Sheet.bas` into the code window for the `Dashboard` worksheet.
5. Put the contents of `vba/Quick_Lookup_Sheet.bas` into the code window for the `QUICK LOOKUP` worksheet.
6. Put the contents of `vba/ThisWorkbook.bas` into `ThisWorkbook`.
7. In the VBA editor choose `Debug > Compile VBAProject`.
8. Save the workbook as a macro-enabled `.xlsm` file.

## What to run after installing the code

If all four code files have just been installed, run this first from the VBA editor or Macro dialog:

`ResetExcelState`

This makes sure Excel events, screen updating, calculation, and alerts are in a normal state.

Then run:

`ShowAllRecords`

This clears the current Vehicle Master filters and returns the master sheet to the All Records view.

At the current repository revision, these are the two main procedures that are ready to run. The repository is being used to rebuild the full dashboard system incrementally, so do not assume procedures from older chat versions exist unless they are present in the current `vba/Module3.bas` file.

## Normal workflow

1. Make/edit VBA changes in this repository.
2. Copy/import the updated file into the matching Excel VBA project location.
3. Run `Debug > Compile VBAProject`.
4. Run `ResetExcelState` if Excel gets stuck with events or screen updating disabled.
5. Test the changed feature in a copy of the workbook before using it on the live file.

## Current workbook assumptions

- Master sheet: `Vehicle Master`
- Dashboard sheet: `Dashboard`
- Quick lookup sheet: `QUICK LOOKUP`
- Header row: `2`
- First data row: `3`
- Open / Closed / All selector target in the planned full dashboard: `Dashboard!C5`

## Important

The current GitHub code is the source of truth. Older VBA pasted in chat may contain functions that are not yet present in the repository. Before running a macro name from an older message, check that it exists in the current repository version.

The full planned dashboard includes dynamic sections for Order / Delivery Stage, Customer, Model, Year, Brand, Body Vendor, Post Delivery, Program of Concession, Open / Closed / All filtering, highlighting, click-to-filter behavior, and Quick Lookup. Those features should be added and tested in GitHub before being copied into the live workbook.
