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
- user-added dashboard sections for any Vehicle Master column

The left-side Program, Year, Brand, and Body Vendor sections are rebuilt as a dynamic vertical stack so one list growing does not overwrite the next section.

## Cleaner dashboard layout

Visible `FILTER ACTION` columns have been removed from the Dashboard. The click actions are stored in hidden helper columns, so clicking a dashboard row still filters Vehicle Master without displaying technical filter strings.

Visible tables are compact 3-column blocks with a spacer column between them:

- Quick Filters: `B:D`
- Status / Order Delivery Stage: `F:H`
- Customer: `J:L`
- Post Delivery: `N:P`
- Model: `R:T`
- Program / Year / Brand / Body Vendor dynamic stack: `B:D`, beginning at row 20
- User-added dashboard sections: `V:X`, stacked automatically

Old visible `FILTER ACTION` leftovers from prior versions are cleared automatically when the Dashboard activates.

## Add any new Vehicle Master column to the Dashboard

A generic dashboard-column system is included.

1. Add the new column to `Vehicle Master` and give it a header in row 2.
2. Go to the Dashboard.
3. In the green **ADD DASHBOARD COLUMN** area, type the exact Vehicle Master header into the yellow input cell (`W6`).
4. Press Enter.
5. VBA verifies that the header exists and automatically creates a new 3-column dashboard section in `V:X`.
6. The new section automatically shows unique values, counts, a blank/missing row, and click-to-filter behavior.
7. The section also follows the Dashboard Open / Closed / All selector.

The saved list of user-added dashboard columns is kept in hidden helper cells `BE2:BE50`. Click actions for these sections are stored in hidden column `BF`.

If the same header is entered twice, VBA will not create a duplicate section. If the header does not exist in Vehicle Master, VBA displays an error instead of building a broken table.

## VBA files

- `vba/Module3.bas` — complete main VIN manager, filtering, self-healing Dashboard, built-in dynamic sections, hidden click actions, and Quick Lookup logic.
- `vba/Dynamic_Dashboard_Columns.bas` — generic system for adding any Vehicle Master column to the Dashboard by typing its header name.
- `vba/Dashboard_Sheet.bas` — Dashboard worksheet events and click handlers.
- `vba/Quick_Lookup_Sheet.bas` — Quick Lookup worksheet events.
- `vba/ThisWorkbook.bas` — workbook startup/reset behavior.

## Install into Excel

1. Open the `.xlsm` workbook.
2. Press `Alt + F11`.
3. Open `Modules > Module3`, delete its old contents, and paste the entire current `vba/Module3.bas` file.
4. Insert another standard VBA module and paste the entire `vba/Dynamic_Dashboard_Columns.bas` file into it.
5. Open the code window for the `Dashboard` worksheet and replace it with `vba/Dashboard_Sheet.bas`.
6. Open the code window for the `QUICK LOOKUP` worksheet and replace it with `vba/Quick_Lookup_Sheet.bas`.
7. Open `ThisWorkbook` and replace it with `vba/ThisWorkbook.bas`.
8. Choose `Debug > Compile VBAProject`.
9. Do not continue until the project compiles without an error.

## First run after installing

Run these macros in this order:

1. `ResetExcelState`
2. `RebuildDashboardLayout`
3. `SetupQuickLookupSheet`
4. `SetupDynamicColumnArea`
5. `RefreshVINSystem`
6. `RefreshCustomDashboardSections`

After the initial setup, normal Dashboard use should not require manually running those macros.

## Self-healing behavior

When the Dashboard worksheet activates, it removes obsolete layout artifacts, runs `RefreshDashboardFast`, and then refreshes any user-added dashboard sections.

`RefreshDashboardFast` checks the required built-in Dashboard sections. If a required section title or controlled layout has been deleted, it runs `RebuildDashboardLayout` and recreates the built-in Dashboard before refreshing the counts and dynamic tables.

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
- Add-dashboard-column input: `Dashboard!W6`

The code looks up Vehicle Master fields by header text rather than relying on fixed Vehicle Master column letters, so rearranging Vehicle Master columns should not break the filters as long as the header names remain unchanged.

## Important

Use the GitHub files as the master copy. Make changes in GitHub first, then copy the updated VBA into Excel and compile before testing the live workbook.
