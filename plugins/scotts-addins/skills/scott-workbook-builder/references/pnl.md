# P&L (Profit & Loss) report

A standard P&L with budget comparison columns. Built off the Inputs sheet — Org ID lives in `Inputs!$B$3` (direct cell reference, no named range), date inputs use the named ranges `PeriodStart`, `PeriodEnd`, plus `BudgetVersion` (defaults to `"Working"`).

## Default columns

Every P&L the skill produces uses these six columns:

| Column | Header | Formula |
|--------|--------|---------|
| A | Code | text |
| B | Description | `=SCOTT.DESC(Inputs!$B$3, $A{row})` |
| C | Actual | `=SCOTT.GL(Inputs!$B$3, $A{row}, PeriodStart, PeriodEnd)` |
| D | Budget | `=SCOTT.BUDGET(Inputs!$B$3, BudgetVersion, $A{row}, PeriodStart, PeriodEnd)` |
| E | Variance $ | income lines: `=C-D`; expense lines: `=D-C` (positive = favorable) |
| F | Variance % | `=IFERROR((C-D)/ABS(D),"")` |

The variance-sign convention: **positive variance always means favorable**. Revenue uses `Actual − Budget`; direct costs and operating expenses use `Budget − Actual`. Variance % is signed (positive = actual exceeds budget). Tell the user this in the handoff and stamp it as a sub-subtitle on the report itself so a casual reader doesn't misread an overrun.

If the user explicitly says "no budget comparison," drop columns D, E, F and revert to a single Actual column. Don't omit them on a hunch — the budget-vs-actual layout is the skill's default.

## Drive the account list from the chart, not from ranges

List every individual revenue, expense, COGS, and overhead account from the user's chart of accounts, in the order the chart returns them, with one `SCOTT.GL` row per account. Do **not** use `SCOTT.RANGE` against guessed code ranges as the primary approach. The chart of accounts comes from the user via the workflow step in `SKILL.md` (the user supplies a workbook generated with `=SCOTT.CHART`). See `references/chart-of-accounts.md` for parsing details and how to map account types onto P&L sections.

Range-based totals are a fallback for the case where the chart genuinely isn't available. When you fall back, say so in the handoff so the user knows the report is approximate.

## Sheet layout

Sheet name: `P&L`. Sections in this order:

```
REVENUE                       <- REVENUE + SALES types
  Total Revenue

DIRECT COSTS                  <- DIRECTCOSTS type
  Total Direct Costs

GROSS PROFIT                  <- = Total Revenue − Total Direct Costs

OPERATING EXPENSES            <- EXPENSE + DEPRECIATN + OVERHEADS types
  Total Operating Expenses

OPERATING INCOME              <- = Gross Profit − Total Operating Expenses

OTHER INCOME                  <- OTHERINCOME type (placed after Operating Income)
  Total Other Income

NET INCOME                    <- = Operating Income + Total Other Income
```

The chart-of-accounts type → section mapping is in `references/chart-of-accounts.md`. **Note two important conventions:**

1. **OVERHEADS-type accounts go into Operating Expenses**, not their own section. Treat them as ordinary operating costs alongside EXPENSE and DEPRECIATN accounts.

2. **Other Income is placed AFTER Operating Income**, not at the top with Revenue. This matches standard GAAP-style presentation — operating performance comes first, then non-operating items roll into Net Income.

## Formula choices

**Revenue and expense lines** — use `SCOTT.GL` (raw values from Xero). Don't reflexively reach for `SCOTT.NGL` to make revenue read as positive — leave the sign as Xero returns it unless the user specifically asks for the flip.

**Account ranges** — use `SCOTT.RANGE(Inputs!$B$3, startCode, endCode, PeriodStart, PeriodEnd)` for section totals when the user hasn't given specific account codes. Common range conventions:
- Revenue: `4000`–`4999`
- COGS: `5000`–`5999`
- Operating expenses: `6000`–`9999`

These are conventions, not laws. Many charts of accounts differ. **Call this out in your handoff message** so the user can adjust if their chart uses different ranges.

**Specific accounts** — when the user gives you account codes (e.g., "build a P&L with revenue accounts 4001, 4002, 4003"), put each code in its own row in column A and use `SCOTT.GL(Inputs!$B$3, $A7, PeriodStart, PeriodEnd)`, with `SCOTT.DESC(Inputs!$B$3, $A7)` in column B for the description. This is the pattern in the `scott-formulas` skill's "Common patterns" section.

**Subtotals and totals** — plain `=SUM()` of the rows above. Use bold + a top border to visually distinguish. These are *not* SCOTT.* formulas — write them with `worksheet.write_formula()`, not `write_dynamic_array_formula()`.

**Gross Profit and Net Income** — explicit subtractions of the prior totals. Don't try to be clever with `SCOTT.XNI` here unless the user asked for it specifically — explicit math is auditable.

## Title row

Cell A1 should be the report name — either what the user gave you, or fall back to "Profit & Loss". Plain text, no formula.

Cell A2 is the period subtitle. Compute the string in Python at build time using the actual period dates and write it as text — for example:

```python
subtitle = f'For the period {period_start.strftime("%b")} {period_start.day}, {period_start.year} through {period_end.strftime("%b")} {period_end.day}, {period_end.year}'
ws.merge_range('A2:E2', subtitle, fmts['subtitle'])
```

Don't use a `="..." & TEXT(PeriodStart, "mmm d, yyyy") & "..."` formula here. Excel will encode `&` as `&amp;` in the saved XML, and on open it shows up as garbage in the formula bar. Static text built in Python is cleaner and the user can always edit it directly.

## Section headers

`REVENUE`, `COST OF GOODS SOLD`, `OPERATING EXPENSES` — bold, slightly larger font (12pt vs. 11pt body), no fill color. Keep it clean and printable.

## Variance / commentary columns

If the user asked for variance vs. PY or vs. budget, add columns to the right:

- **Variance $**: `=C7 - D7` (current minus PY)
- **Variance %**: `=IFERROR((C7 - D7) / ABS(D7), "")`

Format variance % as a percentage. Use `IFERROR` so divide-by-zero shows blank instead of `#DIV/0!`.

## Placeholder handoff

When you've used range-based section totals (4000–4999 etc.) instead of specific account codes, your handoff message should say something like:

> I built the P&L using account ranges (4000–4999 for revenue, 5000–5999 for COGS, 6000+ for OpEx). If your chart of accounts uses different ranges, let me know and I'll adjust. You can also drop specific account codes into column A if you want individual lines instead of section totals.
