# Balance Sheet report

A point-in-time statement of assets, liabilities, and equity. Built off the Inputs sheet — Org ID lives in `Inputs!$B$3` (direct cell reference, no named range), plus an `AsOf` named range (single date) and optionally `PYAsOf` for prior-year comparison.

## Inputs sheet adjustments for BS

A balance sheet doesn't need period start/end — just an "as of" date. When building a workbook with only a Balance Sheet, the Inputs sheet should be:

```
        A                            B
3   Org ID                          112353
5   As of                           4/30/2026
8   Prior year as of                4/30/2025
```

Named ranges: `AsOf`, `PYAsOf` (Org ID is a direct `Inputs!$B$3` cell reference, not a named range).

When the workbook also includes a P&L, keep the P&L inputs and add the BS inputs below them — don't try to share `PeriodEnd` with `AsOf`. They're conceptually different and the user might want them at different dates.

## Balance dates with SCOTT.GL

`SCOTT.GL` takes a start and end date and returns the *change* in the account over that range. For balance sheet accounts, the user wants the *cumulative balance* at a point in time. The convention is to pass a very early start date (e.g., `1/1/1900` or the org's first fiscal year) so the function returns the running balance through `AsOf`.

A reusable approach: define a named range `BSStart` on the Inputs sheet pointing at a sentinel date like `1/1/1900`. Then BS formulas read:

```
=SCOTT.GL(Inputs!$B$3, "1000", BSStart, AsOf)
```

This is a quirk of how the GL function works, not a limitation — once explained, it's fine. **Call it out in your handoff message** so the user knows what `BSStart` is for and doesn't accidentally change it.

## Sheet layout

Sheet name: `Balance Sheet`. Sections (in order):

```
ASSETS
  Cash and Cash Equivalents       <- BANK accounts
    Total Cash and Cash Equivalents
  Current Assets                  <- CURRENT, INVENTORY, PREPAYMENT
    Total Current Assets
  Fixed Assets                    <- FIXED accounts
    Total Fixed Assets
  Long Term Assets                <- NONCURRENT accounts
    Total Long Term Assets
TOTAL ASSETS

LIABILITIES
  Current Liabilities             <- CURRLIAB
    Total Current Liabilities
  Long Term Liabilities           <- TERMLIAB, LIABILITY (NONCURRLIAB)
    Total Long Term Liabilities
TOTAL LIABILITIES

EQUITY
  [equity accounts]               <- EQUITY accounts
  Current Year Earnings           <- =SCOTT.XNI(Inputs!$B$3, FYStart, AsOf)
  Total Equity

TOTAL LIABILITIES + EQUITY

Check: Assets − (Liab + Equity)   <- should be 0
```

The chart-of-accounts type → section mapping is in `references/chart-of-accounts.md`. **Fixed Assets and Long Term Assets are separate subsections** — FIXED-type accounts go to Fixed Assets, NONCURRENT-type accounts go to Long Term Assets. Don't combine them into a single "Non-current Assets" section even though they're both technically non-current.

The balance check row should always be `0`. Format as currency with red conditional fill if non-zero. This catches both data entry errors and formula errors during construction.

Row 34 is a balance check — it should always be `0`. Format as currency with red conditional fill if non-zero. This catches both data entry errors and formula errors during construction.

## Drive the account list from the chart, not from ranges

Do **not** use `SCOTT.RANGE` against guessed code ranges (e.g., "1000–1499 for cash"). Account numbering varies wildly by org — what's a fixed asset for one org is a deferred revenue for another, and one missing 100-block of codes will silently exclude live accounts from the report.

Instead, list every individual asset, liability, and equity account from the user's chart of accounts, in the order they appear there, with one `SCOTT.GL` row per account. The chart of accounts comes from the user via the workflow step in `SKILL.md` (the user supplies a workbook generated with `=SCOTT.CHART`). See `references/chart-of-accounts.md` for parsing details and how to map account types onto BS sections.

If for some reason the chart isn't available (offline mode, user can't authenticate), fall back to ranges only as a last resort, and call it out prominently in the handoff so the user knows the report is approximate.

## Sign conventions on the BS

Don't reflexively flip signs. `SCOTT.GL` returns whatever Xero stores — assets typically positive, liabilities and equity typically negative (credit balances). If the user wants liabilities and equity to display as positive, ask them — don't silently substitute `SCOTT.NGL`. Default to `SCOTT.GL` for everything and let the user request a sign flip explicitly.

## BANK accounts with credit balances — section convention

Each chart account is classified into a BS section by its chart type, exclusively. A BANK-type account always goes to "Cash and Cash Equivalents," even when its balance is a credit (overdrawn cash account, or a credit card with an unpaid balance). The credit balance shows up as a negative number in the Cash section.

Xero's native Balance Sheet makes a different choice — it reclassifies BANK accounts with credit balances to Current Liabilities for display, with the sign flipped to positive. Both presentations are mathematically equivalent (Total Assets and Total Liabilities + Equity tie the same way), but the section composition will diverge from Xero's BS for any period where a BANK account carries a credit balance.

This is a deliberate design choice. Reclassifying at runtime would require either (a) reading both period balances and routing rows conditionally — complex and fragile when the comparison column for prior year disagrees with the current — or (b) pre-fetching balances at build time, which breaks the live-formula architecture. The simpler, more debuggable behavior is to honor the chart-of-accounts type as the single source of truth for section assignment.

**Always include this footnote at the bottom of every Balance Sheet** so users comparing to their accounting system's native BS understand the difference. Sample text:

> Note on bank accounts with credit balances: Each account is classified by its chart-of-accounts type. BANK-type accounts (cash, credit cards) are presented in Cash and Cash Equivalents regardless of balance sign. When a BANK account has a credit balance — typically a credit card with an unpaid balance, or an overdrawn cash account — it appears here with a negative number. Some accounting systems reclassify such accounts to Current Liabilities for display. Both presentations are mathematically equivalent (Total Assets and Total Liabilities + Equity tie out the same way), but the section composition will differ from a Xero/QuickBooks-native Balance Sheet for any period in which a BANK account carries a credit balance.

## Title

Cell A1 falls back to "Balance Sheet" if the user didn't name it. Cell A2 should be the static subtitle, computed in Python at build time:

```python
subtitle = f'As of {as_of.strftime("%b")} {as_of.day}, {as_of.year}'
ws.merge_range('A2:D2', subtitle, fmts['subtitle'])
```

Don't use `="As of " & TEXT(AsOf, "mmm d, yyyy")` — the `&` gets XML-encoded and shows up as garbage in the formula bar on open.

## Equity tie-out

Be aware that the equity section needs to include current-year retained earnings to balance. Xero handles this automatically in its own balance sheet by computing net income for the year-to-date through the `AsOf` date. If the workbook isn't balancing, the most likely culprit is a missing CY net income line. Either pull it explicitly:

```
=SCOTT.XNI(Inputs!$B$3, YTDStartForBS, AsOf)
```

…on its own equity row (with `YTDStartForBS` as a named range pointing at the year-start date), or note in the handoff that the user may need to add this manually depending on whether their Xero closes month-by-month.
