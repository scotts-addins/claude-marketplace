# Inputs sheet architecture

The Inputs sheet is the brain of the workbook. Every formula on every other sheet references it. Get this right and the user can re-point the whole workbook at a new period or org by changing one cell.

## Why a dedicated sheet (not just a row at the top)

The user is going to want to use this workbook again next month, and the month after, and probably for years. A scattered "inputs at the top of each sheet" approach means they have to remember to update every sheet. A dedicated Inputs sheet means they update one place and everything refreshes. Worth the small extra effort up front.

## Standard layout

Sheet name: `Inputs` (always — keep it consistent so users learn where to look).

```
        A                            B
1   Inputs
2   
3   Org ID                          112353
4   
5   Period start                    4/1/2026
6   Period end                      4/30/2026
7   
8   Prior year period start         4/1/2025
9   Prior year period end           4/30/2025
10  
11  YTD start                       1/1/2026
12  YTD end                         4/30/2026
```

Adapt as needed:
- For Balance Sheet, only `As of` date (single cell) and prior year `As of` are needed instead of period start/end.
- For Budget vs. Actual, add `Budget version` (text, e.g., `"Working"` or `"Approved"`) and any tracking-category cells the user wants to filter by.
- For multi-period reports (e.g., a 12-month rolling P&L), include a starting month cell and let the report sheet derive each month from it via `EDATE`.

## Named ranges and the Org ID cell reference

Define named ranges for the date and parameter inputs, but reference the Org ID directly as `Inputs!$B$3` in formulas — do **not** create an `OrgID` named range:

| Reference style | Cells |
|-----------------|-------|
| Direct cell reference (`Inputs!$B$3`) | Org ID |
| Named range | `PeriodStart` → `Inputs!$B$5` |
| Named range | `PeriodEnd` → `Inputs!$B$6` |
| Named range | `PYStart` → `Inputs!$B$8` |
| Named range | `PYEnd` → `Inputs!$B$9` |
| Named range | `YTDStart` → `Inputs!$B$11` |
| Named range | `YTDEnd` → `Inputs!$B$12` |

So formulas on the report sheets read like:
```
=SCOTT.GL(Inputs!$B$3, "4000", PeriodStart, PeriodEnd)
=SCOTT.RANGE(Inputs!$B$3, "4000", "4999", PeriodStart, PeriodEnd)
=SCOTT.BUDGETR(Inputs!$B$3, BudgetVersion, "4000", "4999", PeriodStart, PeriodEnd)
```

In `xlsxwriter`, named ranges are added via `workbook.define_name("PeriodStart", "=Inputs!$B$5")`. The cell reference for Org ID needs no setup — it's just `Inputs!$B$3` typed into the formula string.

## Formatting

- The label column (column A) is bold and right-aligned.
- The value column (column B) is left-aligned, with date cells formatted as dates (`m/d/yyyy`) and the Org ID as a plain number (no thousands separator — it's an identifier, not an amount).
- Set column A width to ~30, column B to ~18.
- Add a thin border under each labeled section to visually group related inputs.

## Things to never do

- **Never hardcode dates or org IDs into report-sheet formulas.** The whole point of the Inputs sheet is broken if some formulas reference Inputs and others don't.
- **Never put dates as text strings.** They have to be Excel date serial values (entered with `openpyxl.utils.datetime` or as `datetime` objects). SCOTT formulas reject text dates.
- **Never put a spilling formula on the Inputs sheet.** This sheet is for static parameters only. If you need to derive things, use a separate "Calcs" sheet.
