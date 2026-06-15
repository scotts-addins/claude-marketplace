"""Canonical xlsxwriter pattern for building Scott's Add-ins workbooks.

This is a small, runnable example that demonstrates every important
choice this skill makes. Copy it as a starting point — don't try to
build a workbook from scratch each time.

KEY RULES (also in SKILL.md, repeated here because they bite):

1. Use xlsxwriter, not openpyxl. xlsxwriter has native dynamic-array
   formula support; openpyxl does not, and post-processing the file
   is fragile.

2. For ANY cell whose formula contains "SCOTT.", call
   ws.write_dynamic_array_formula(cell, formula, fmt). This writes the
   cm="1" cellMetadata reference and registers xl/metadata.xml — which
   is what stops Excel from injecting "@" in front of the formula on
   open.

3. After wb.close(), call xldudf_prefix_scott_calls(out_path) to rewrite
   every SCOTT.<NAME> in the saved formula text to _xldudf_SCOTT_<NAME>.
   This is REQUIRED — without it the cells show #VALUE or #NAME on
   first open and only resolve when the user manually edits the cell.
   See the function's docstring for the full explanation.

4. For ordinary formulas (=SUM, =IFERROR, =C10-C16), use
   ws.write_formula(cell, formula, fmt). They don't need DA marking and
   are not touched by the _xldudf_ post-processor.

5. Org ID is a DIRECT CELL REFERENCE (Inputs!$B$3), not a named range.
   Other inputs (PeriodStart, PeriodEnd, BudgetVersion) are named ranges.

6. Build subtitles as static Python strings — never as Excel formulas
   that use & for concatenation. Excel encodes & as &amp; in the saved
   XML and it shows up as garbage in the formula bar on open.

7. Don't reflexively flip signs. Use SCOTT.GL, not SCOTT.NGL, unless
   the user specifically asks for the negation.
"""
import os
import re
import shutil
import zipfile
from datetime import datetime
import xlsxwriter


# Friendly-to-canonical function name mapping for SCOTT.* calls.
#
# Most functions store as `_xldudf_SCOTT_<NAME>` where <NAME> matches
# what the user types (SCOTT.GL -> _xldudf_SCOTT_GL). A few don't, and
# their canonical (Excel-stored) name has been confirmed empirically by
# inspecting workbooks that Scott has saved after typing the formula
# natively in the formula bar.
#
# Add new entries here as we discover them. The key is the friendly name
# (what the user/skill writes), the value is the canonical name Excel
# uses internally. Anything not in this map is assumed to store
# unchanged.
FUNC_NAME_MAP = {
    # XTRACK family — leading X is stripped
    'XTRACK':  'TRACK',
    'XTRACKM': 'TRACKM',
    'XTRACKR': 'TRACKR',
    # Range aggregator — registered internally under a different name
    'RANGE':   'COMRANGE',
}


def xldudf_prefix_scott_calls(xlsx_path):
    """Post-process the saved xlsx so every SCOTT.<NAME> in formula text
    is rewritten to its canonical `_xldudf_SCOTT_<canonical_name>` form.

    Why this is required
    --------------------
    Excel itself stores JavaScript custom-function calls with an
    `_xldudf_` prefix and with dots replaced by underscores in the saved
    sheet XML. When the user types `=SCOTT.GL(...)` into the formula
    bar, Excel silently translates that to `_xldudf_SCOTT_GL(...)`
    before storing the cell, and translates back to the friendly
    `SCOTT.GL` form for display.

    xlsxwriter has no awareness of this convention — it writes the
    formula text literally as `SCOTT.GL`. On file open, Excel doesn't
    recognize the dot-form as a UDF call (the dot reads as a structured-
    reference operator), so it caches `#NAME` or `#VALUE` for every
    SCOTT cell. The only way to clear it without this fix is for the
    user to edit each cell by hand, which triggers the formula-bar
    re-parse that does the translation.

    Friendly-vs-canonical name mismatch
    -----------------------------------
    Most functions store unchanged: `SCOTT.GL` -> `_xldudf_SCOTT_GL`.
    A few don't, and have to be mapped explicitly via FUNC_NAME_MAP
    above. Currently known exceptions:

      - XTRACK / XTRACKM / XTRACKR  -> the leading X is stripped
      - RANGE                       -> stored as COMRANGE

    Any time a SCOTT.* call hits #NAME on first open despite the
    `_xldudf_` prefix being applied, suspect this kind of mismatch.
    Have the user type the formula manually, save, and inspect the
    canonical name in the saved XML; then add the mapping here.
    """
    pattern = re.compile(r'\bSCOTT\.([A-Z][A-Z0-9_]*)')

    def repl(m):
        friendly = m.group(1)
        canonical = FUNC_NAME_MAP.get(friendly, friendly)
        return f'_xldudf_SCOTT_{canonical}'

    tmp = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        names = zin.namelist()
        blobs = {n: zin.read(n) for n in names}
    for n in list(blobs.keys()):
        if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
            text = blobs[n].decode()
            blobs[n] = pattern.sub(repl, text).encode()
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, blobs[n])
    shutil.move(tmp, xlsx_path)

# Number formats
CURRENCY_FMT = '#,##0_);(#,##0);"\u2013"_)'
CURRENCY_RED_FMT = '#,##0_);[Red](#,##0);"\u2013"_)'
PCT_FMT = '0.0%'
DATE_FMT = 'm/d/yyyy'


def make_formats(wb):
    """Standard format set used across every report."""
    return {
        'title': wb.add_format({'font_name': 'Calibri', 'font_size': 16, 'bold': True, 'align': 'center', 'bg_color': '#F2F2F2'}),
        'subtitle': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'italic': True, 'font_color': '#666666', 'align': 'center'}),
        'section': wb.add_format({'font_name': 'Calibri', 'font_size': 12, 'bold': True}),
        'header': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'align': 'center', 'bottom': 1}),
        'label_right': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'align': 'right'}),
        'body': wb.add_format({'font_name': 'Calibri', 'font_size': 11}),
        'body_currency': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': CURRENCY_FMT}),
        'body_currency_red': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': CURRENCY_RED_FMT}),
        'body_pct': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': PCT_FMT}),
        'total_currency': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'num_format': CURRENCY_FMT, 'top': 1}),
        'total_pct': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'num_format': PCT_FMT, 'top': 1}),
        'grand_currency': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'num_format': CURRENCY_FMT, 'top': 2}),
        'grand_pct': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True, 'num_format': PCT_FMT, 'top': 2}),
        'bold_label': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'bold': True}),
        'date': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': DATE_FMT}),
        'integer': wb.add_format({'font_name': 'Calibri', 'font_size': 11, 'num_format': '0'}),
        'note': wb.add_format({'font_name': 'Calibri', 'font_size': 10, 'italic': True, 'font_color': '#666666'}),
    }


def write_inputs_sheet(wb, fmts, rows):
    """Build the Inputs sheet.

    rows is a list of (label, value, name_or_None, fmt_key_or_None).
    - label is the description in column A.
    - value goes in column B.
    - name_or_None: if a string, defines a workbook-level named range
      pointing at Inputs!$B$<row>. Use None for Org ID — it's a direct
      cell reference, not a named range.
    - fmt_key_or_None: a key into the fmts dict, e.g. 'date', 'integer'.
    - An empty label produces a blank spacer row.
    """
    ws = wb.add_worksheet('Inputs')
    ws.set_tab_color('#5B9BD5')
    ws.set_column('A:A', 30)
    ws.set_column('B:B', 18)
    ws.write('A1', 'Inputs', fmts['title'])
    for i, (label, value, name, fmt_key) in enumerate(rows, start=3):
        if not label:
            continue
        ws.write(i - 1, 0, label, fmts['label_right'])
        cell_fmt = fmts.get(fmt_key) if fmt_key else fmts['body']
        if isinstance(value, datetime):
            ws.write_datetime(i - 1, 1, value, cell_fmt)
        else:
            ws.write(i - 1, 1, value, cell_fmt)
        if name:
            wb.define_name(name, f"=Inputs!$B${i}")
    ws.freeze_panes(1, 0)
    return ws


def write_scott(ws, cell, formula, fmt):
    """Write any formula containing SCOTT.* — uses dynamic-array form
    so Excel doesn't inject @ on open. ALWAYS use this for SCOTT.*."""
    ws.write_dynamic_array_formula(cell, formula, fmt)


# ---------- Example: minimal monthly P&L ----------

def build_example_pnl(out_path):
    wb = xlsxwriter.Workbook(out_path)
    fmts = make_formats(wb)

    # Inputs sheet — note Org ID has name=None (no named range).
    # The 112353 value below is illustrative only. In real builds, Org ID
    # MUST come from an explicit user answer (see SKILL.md workflow step 3).
    # Never hardcode an Org ID — always pass it in from the user's response.
    inputs = [
        ('Org ID', 112353, None, 'integer'),                           # Inputs!$B$3 — direct cell ref
        ('', None, None, None),
        ('Period start', datetime(2026, 4, 1), 'PeriodStart', 'date'), # Inputs!$B$5
        ('Period end',   datetime(2026, 4, 30), 'PeriodEnd', 'date'),  # Inputs!$B$6
    ]
    write_inputs_sheet(wb, fmts, inputs)

    ws = wb.add_worksheet('P&L')
    ws.set_column('A:A', 10)
    ws.set_column('B:B', 32)
    ws.set_column('C:C', 16)

    # Title — merged cell with static text
    ws.merge_range('A1:C1', 'Profit & Loss', fmts['title'])
    ws.set_row(0, 28)

    # Subtitle — built in Python from the actual dates, NOT an Excel & formula
    period_start = datetime(2026, 4, 1)
    period_end = datetime(2026, 4, 30)
    subtitle = f'For the period {period_start.strftime("%b")} {period_start.day}, {period_start.year} through {period_end.strftime("%b")} {period_end.day}, {period_end.year}'
    ws.merge_range('A2:C2', subtitle, fmts['subtitle'])

    # Column headers
    for i, h in enumerate(['Code', 'Description', 'Current period']):
        ws.write(3, i, h, fmts['header'])

    # REVENUE section
    ws.write('A6', 'REVENUE', fmts['section'])
    ws.write('A7', '4000-4999')
    ws.write('B7', 'All revenue (range total)')

    # SCOTT.* formula — write_dynamic_array_formula, NOT write_formula.
    # Org ID is the literal cell reference Inputs!$B$3.
    # PeriodStart, PeriodEnd are named ranges defined above.
    write_scott(ws, 'C7',
        '=SCOTT.RANGE(Inputs!$B$3,"4000","4999",PeriodStart,PeriodEnd)',
        fmts['body_currency'])

    # Subtotal — ordinary =SUM, ordinary write_formula
    ws.write('B9', 'Total Revenue', fmts['bold_label'])
    ws.write_formula('C9', '=SUM(C7:C8)', fmts['total_currency'])

    # EXPENSES section
    ws.write('A11', 'EXPENSES', fmts['section'])
    ws.write('A12', '6000-9999')
    ws.write('B12', 'All operating expenses (range total)')
    write_scott(ws, 'C12',
        '=SCOTT.RANGE(Inputs!$B$3,"6000","9999",PeriodStart,PeriodEnd)',
        fmts['body_currency'])

    ws.write('B14', 'Total Expenses', fmts['bold_label'])
    ws.write_formula('C14', '=SUM(C12:C13)', fmts['total_currency'])

    # NET INCOME — ordinary subtraction, not a SCOTT.* call
    ws.write('B16', 'NET INCOME', fmts['bold_label'])
    ws.write_formula('C16', '=C9-C14', fmts['grand_currency'])

    ws.freeze_panes(4, 0)
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.set_margins(left=0.5, right=0.5)

    wb.close()

    # REQUIRED: rewrite SCOTT.<NAME> -> _xldudf_SCOTT_<NAME> in the saved
    # XML. Without this, every SCOTT cell shows #VALUE or #NAME on first
    # open and only resolves when the user edits the cell by hand.
    xldudf_prefix_scott_calls(out_path)

    return out_path


if __name__ == '__main__':
    out_dir = os.environ.get('OUT', '.')
    p = build_example_pnl(os.path.join(out_dir, 'example_pnl.xlsx'))
    print(f"Built {p}")
