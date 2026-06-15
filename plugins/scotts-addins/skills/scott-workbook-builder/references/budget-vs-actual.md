# Budget vs. Actual report

A P&L-style report comparing actuals (`SCOTT.GL`) against budget (`SCOTT.BUDGET` / `SCOTT.BUDGETR`) for the same period, with variance columns.

## Drive the account list from the chart, not from ranges

Like the P&L, every revenue and expense account on this report should come from the user's chart of accounts (workbook generated with `=SCOTT.CHART`), one row per account. Do **not** use `SCOTT.BUDGETR` against guessed code ranges as the primary approach — that hides which accounts the budget is missing for, and a user who's checking a budget vs. actual specifically wants per-account precision. See `references/chart-of-accounts.md` for parsing and section mapping.

## Inputs sheet additions

In addition to the standard P&L inputs (Org ID at `Inputs!$B$3`, named ranges `PeriodStart` and `PeriodEnd`), Budget vs. Actual needs:

```
Budget version          Working
```

The version is a text string matching what's in Xero's budget setup (commonly `"Working"`, `"Approved"`, `"Forecast"`, or a custom name). Define a named range `BudgetVersion` for it.

If the user didn't say which version, ask — guessing wrong here means every budget number comes back blank.

## Sheet layout

Sheet name: `Budget vs Actual`.

```
       A         B                  C            D            E              F
1                              [Company / Report Title]
2                              Budget vs. Actual — [PeriodStart] – [PeriodEnd]
3                              Budget version: [BudgetVersion]
4
5   Code   Description        Actual       Budget       Variance $    Variance %
6
7   REVENUE
8   4000   Revenue            [SCOTT.GL]   [SCOTT.BUDGET]  =C8-D8     =IFERROR((C8-D8)/ABS(D8),"")
9   ...
10         Total Revenue      =SUM         =SUM           =C10-D10    =IFERROR(...)
11
12  EXPENSES
13  6000   Salaries           [SCOTT.GL]   [SCOTT.BUDGET]  =D13-C13   ← favorable when actual < budget
14  ...
15         Total Expenses     =SUM         =SUM           =D15-C15
16
17         NET INCOME         [Rev − Exp]  [Rev − Exp]    =C17-D17   =IFERROR(...)
```

## Variance sign convention

The crucial design choice: for revenue lines, "variance favorable" means actual > budget (more revenue is good). For expense lines, "variance favorable" means actual < budget (less spending is good). If you compute `Actual − Budget` everywhere, expense overruns show as positive numbers, which can read as "good" to a casual scanner.

Two acceptable conventions:

1. **Always compute Actual − Budget**, and add a separate "F/U" column (Favorable/Unfavorable) that reads `="F"` when revenue is positive variance or expense is negative variance, and `="U"` otherwise.

2. **Flip the sign on the expense variance**: revenue uses `=Actual − Budget`, expenses use `=Budget − Actual`. So positive variance always means "favorable." Less mathematically pure but reads more naturally for non-finance users.

Default to option 2 unless the user is clearly a finance professional who'll find it confusing. Call out the convention you used in the handoff message.

## Budget formula details

The standard single-account form is:

```
=SCOTT.BUDGET(Inputs!$B$3, BudgetVersion, accountCode, PeriodStart, PeriodEnd)
```

The range form for section totals is:

```
=SCOTT.BUDGETR(Inputs!$B$3, BudgetVersion, "4000", "4999", PeriodStart, PeriodEnd)
```

Note that `BudgetVersion` is the **second** argument in both — easy to get wrong. For multi-account specific lists, use `SCOTT.BUDGETMULTI`. See `scott-formulas/references/formula-reference.md` for exact parameter order on every variant.

## What if the budget hasn't been entered for a period

`SCOTT.BUDGET` returns 0 (not an error) if no budget exists for an account/period combination. This means the workbook will build successfully but the budget column will read all zeros. Mention this possibility in your handoff if it seems likely (e.g., user is asking for next month's budget).

## Tracking-category-filtered version

If the user wants a Budget vs. Actual filtered to a specific cost center, the actual side becomes `SCOTT.XTRACK` instead of `SCOTT.GL`. Budget filtering by tracking category isn't supported in `SCOTT.BUDGET` directly — if the user needs this, ask them whether they want budget unfiltered (showing the org-wide budget) or whether they have a separate budget version per tracking category that they can name explicitly.
