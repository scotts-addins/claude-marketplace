# Formatting standards

The output should look like a finished deliverable, not a prototype. These standards cover the basics — apply them to every workbook this skill produces.

## Fonts and sizes

- **Default body**: Calibri 11pt, black.
- **Section headers** (REVENUE, ASSETS, etc.): Calibri 12pt, bold.
- **Title** (row 1): Calibri 16pt, bold.
- **Subtitle** (row 2 — period reference): Calibri 11pt, italic, gray (`#666666`).

Never mix fonts within a single workbook.

## Number formats

- **Currency cells**: `#,##0_);(#,##0);"–"_)` — comma separators, parens for negatives, en-dash for zero. No currency symbol unless the user asks (most accountants prefer the column header or a single note to indicate currency).
- **Percentage cells**: `0.0%` — one decimal place.
- **Dates**: `m/d/yyyy` (US) — keep it consistent across the workbook.
- **Org ID and account codes**: plain integer or text, no thousands separator.

For currency cells where negatives are normal (variance columns), consider `#,##0_);[Red](#,##0);"–"_)` to color negatives red.

## Column widths

Set explicitly — don't rely on auto-fit, which makes the file look raw.

- Account code column (column A): width 10
- Description column (column B): width 30
- Amount columns (C, D, E, etc.): width 16
- Variance % columns: width 12

## Row heights

Default row height (15) is fine for body rows. For section headers, bump to 20 for breathing room. For title row, 28.

## Borders and shading

Use sparingly. The aesthetic is "clean accountants' worksheet," not "Excel circa 2003."

- Thin bottom border under the column header row (row 4 typically).
- Thin top border on subtotal rows.
- Thicker (medium) top border on grand totals (Total Revenue, Net Income).
- No fill colors anywhere except the title row, which can be a very pale gray (`#F2F2F2`) to set it apart.

Never use:
- Diagonal borders
- Dotted/dashed borders
- Bright fill colors
- More than two border weights in a single sheet

## Frozen panes

Freeze the rows above the data (rows 1–4 typically) so headers stay visible when scrolling. In `xlsxwriter`: `ws.freeze_panes(row, col)` — note this takes 0-indexed row/col, so to freeze the first 4 rows use `ws.freeze_panes(4, 0)`.

## Print setup

Even if the user never prints, decent print settings show care:
- Page orientation: landscape for P&L and Budget vs. Actual (multiple amount columns); portrait for single-period Balance Sheet.
- Fit to width: 1 page wide.
- Margins: narrow.
- Print area: explicit (don't let it default to "everything used").

In `xlsxwriter`:
```python
ws.set_landscape()             # or ws.set_portrait()
ws.fit_to_pages(1, 0)          # 1 wide, unlimited tall
ws.set_margins(left=0.5, right=0.5)
```

## Tab colors

Color-code sheet tabs by purpose:
- Inputs: `#5B9BD5` (blue) — denotes "set things here"
- P&L / Balance Sheet / Budget vs Actual: no tab color (default)
- Reference sheets like a Chart of Accounts pull: `#A5A5A5` (gray)

## Worksheet protection

Don't protect any sheet. The user is going to want to drop in their own account codes, adjust ranges, and add columns. Protection gets in the way.

## What "looks finished" means in practice

If the user opens the file and immediately wants to format something, the skill failed at this step. The bar is: open the file, see the data, optionally print, done. If something looks off (currency without a format, dates rendering as 5-digit numbers, columns clipped), go back and fix it before handing off.
