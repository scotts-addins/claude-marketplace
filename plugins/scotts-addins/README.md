# scotts-addins

A Claude plugin for users of [Scott's Add-ins](https://scottsaddins.com).

## Skills bundled in this plugin

### scott-workbook-builder
Builds finished Excel workbooks (`.xlsx` files) populated with `=SCOTT.*` formulas, driven by a dedicated Inputs sheet. Reports include Profit & Loss, Balance Sheet, Budget vs. Actual, 12-month P&L, transaction registers by GL account, sales-quantity searches, and custom layouts on request.

Triggers on phrases like:
- "Build me a P&L using Scott's Add-ins"
- "Create a balance sheet workbook with Scott's formulas"
- "Make a budget vs. actual file using Scott's"
- "I'd like to create an Excel workbook using Scott's add-in"

### scott-formulas
A reference and conversational helper for `=SCOTT.*` formula syntax. Triggers when you ask formula questions without requesting a file:
- "How do I write `SCOTT.GL`?"
- "What arguments does `SCOTT.BUDGET` take?"
- "What's the difference between `SCOTT.XTRACK` and `SCOTT.QCLASS`?"

The two skills work together. `scott-workbook-builder` consults `scott-formulas/references/formula-reference.md` when building, so this plugin must be installed as a unit — don't try to install the two skills separately.

## How it works

When you ask Claude (in Cowork or Claude Code) to build a workbook:

1. Claude asks what report type, what accounting system (Xero or QuickBooks), what org ID, what period.
2. You upload your chart of accounts (one-time, generated via `=SCOTT.CHART`).
3. Claude builds the workbook using `xlsxwriter`, applies a post-processor that rewrites formulas into the canonical `_xldudf_SCOTT_*` form Excel expects, and saves it.
4. You open the workbook in Excel, the add-in evaluates the formulas, and your numbers populate.

## First-time setup of a built workbook

Every workbook Claude builds includes step-by-step instructions in the chat handoff. The short version:

1. Open the `.xlsx` file Claude built.
2. Cells will initially show `#VALUE` — a known Microsoft bug. Load Scott's Add-ins from the Excel task bar.
3. Connect the workbook to your organization(s) via the add-in.
4. Click the Recalc button under the add-in's Function tab.
5. If still `#VALUE`, save, restart Excel, and reopen — that fully clears the cache.

After that one-time dance, the workbook is fully live: change the Inputs sheet and recalc to re-point the report.

## Versioning

See `plugin.json` for the current version. Updates roll out via the marketplace.
