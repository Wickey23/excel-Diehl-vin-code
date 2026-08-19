# Diehl VIN Excel VBA

This repository is the source of truth for the Diehl VIN workbook VBA.

## Current architecture

The current code is the restored **self-healing dashboard** version.

If a controlled Dashboard section is deleted or damaged, `RefreshDashboardFast` calls `EnsureDashboardStructure`, detects the missing section, and rebuilds the Dashboard automatically.

The Dashboard includes:

- Quick Filters
- Open / Closed / All selector
- Program of Concession
- Year
- Brand
- Body Vendor
- Order / Delivery Stage
- Customer
- Post Delivery
- Model
- dynamic counts and newly added values
- blank / missing-value rows
- click-to-filter behavior
- Quick Lookup sheet
- self-repair / rebuild

The left-side Program, Year, Brand, and Body Vendor sections are rebuilt as a dynamic vertical stack so one list growing does not overwrite the next section.

## Cleaner dashboard layout

Visible `FILTER ACTION` columns have been removed from the Dashboard. The click actions are stored in hidden helper columns `AZ:BD`, so clicking a dashboard row still filters Vehicle Master without displaying technical filter strings.

Visible tables are compact 3-column blocks with a spacer column between them:

- Quick Filters: `B:D`
- Status / Order Delivery Stage: `F:H`
- Customer: `J:L`
- Post Delivery: `N:P`
- Model: `R:T`
- Program / Year / Brand / Body Vendor dynamic stack: `B:D`, beginning at row 20

## VBA files

- `vba/Module3.bas` — complete main VIN manager, filtering, self-healing Dashboard, dynamic sections, hidden click actions, and Quick Lookup logic.
- `vba/Dashboard_Sheet.bas` — Dashboard worksheet events and click handlers.
- `vba/Quick_Lookup_Sheet.bas` — Quick Lookup worksheet events.
- `vba/ThisWorkbook.bas` — workbook startup/reset behavior.

## Install into Excel

1. Open the `.xlsm` workbook.
2. Press `Alt + F11`.
3. Open `Modules > Module3`, delete its old contents, and paste the entire current `vba/Module3.bas` file.
4. Open the code window for the `Dashboard` worksheet and replace it with `vba/Dashboard_Sheet.bas`.
5. Open the code window for the `QUICK LOOKUP` worksheet and replace it with `vba/Quick_Lookup_Sheet.bas`.
6. Open `ThisWorkbook` and replace it with `vba/ThisWorkbook.bas`.
7. Choose `Debug > Compile VBAProject`.
8. Do not continue until the project compiles without an error.

## First run after installing

Run these macros in this order:

1. `ResetExcelState`
2. `RebuildDashboardLayout`
3. `SetupQuickLookupSheet`
4. `RefreshVINSystem`

After the initial setup, normal Dashboard use should not require manually running those macros.

## Self-healing behavior

When the Dashboard worksheet activates, it runs `RefreshDashboardFast`.

`RefreshDashboardFast` checks the required Dashboard sections. If a section title or controlled layout has been deleted, it runs `RebuildDashboardLayout` and recreates the Dashboard before refreshing the counts and dynamic tables.

To manually force a repair at any time, run:

`RebuildDashboardLayout`

Then run:

`RefreshVINSystem`

## Workbook assumptions

- Master sheet: `Vehicle Master`
- Dashboard sheet: `Dashboard`
- Quick Lookup sheet: `QUICK LOOKUP`
- Vehicle Master header row: `2`
- Vehicle Master first data row: `3`
- Open / Closed / All selector: `Dashboard!C5`

The code looks up Vehicle Master fields by header text rather than relying on fixed Vehicle Master column letters, so rearranging Vehicle Master columns should not break the filters as long as the header names remain unchanged.

## Important

Use the GitHub files as the master copy. Make changes in GitHub first, then copy the updated VBA into Excel and compile before testing the live workbook.
