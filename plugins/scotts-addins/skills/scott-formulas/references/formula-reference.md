# Scott's Add-ins Formula Reference

Complete reference for all `=SCOTT.*` functions. Parameters in `[brackets]` are optional.

## Conventions used in this reference

- **Spilling arrays** — many of the data-fetch functions (`SCOTT.CHART`, `SCOTT.LIST`, `SCOTT.XINVOICE`, `SCOTT.XBILL`, `SCOTT.XGLTRX`, `SCOTT.XPAYMENT`, the various `_HEADER` variants) return multi-column / multi-row results that spill from the anchor cell. Always leave room below and to the right.
- **Dates in returned arrays** — date columns inside spilling arrays come back as **ISO-format strings** (`"2026-04-01"`), not native Excel dates. To filter or sort by date in dependent formulas, wrap the column in `DATEVALUE()`.
- **Status values** — invoices and bills use Xero's status set: `"AUTHORISED"`, `"Awaiting Payment"`, `"Paid"`, `"VOIDED"`, `"DRAFT"`. Note the British spelling on `AUTHORISED`.
- **Variadic parameter shape** — `SCOTT.GLMULTI`, `SCOTT.BUDGETMULTI`, `SCOTT.GLMULTITRACK`, `SCOTT.XGLTRX`, `SCOTT.XINVOICE`, `SCOTT.XBILL`, `SCOTT.XPAYMENT` all accept variable numbers of trailing arguments (account codes or contact names). The variadic args always come **last**, after the date and filter parameters. Use `TEXTSPLIT(cell, ", ")` or `FILTER(table[col], criterion)` to drive these dynamically.
- **Spill output is not safely consumable via spill refs (`A2#`).** When you want to do downstream arithmetic on a SCOTT spill — `SUMPRODUCT`, `SUMIFS`, `FILTER`, etc. — reference the cells via **static ranges** (`Invoices!E2:E10000`), not via the spill-anchor syntax (`Invoices!A2#`). Custom-function spills render in the cells but don't behave like a true dynamic array for downstream array functions; `CHOOSECOLS(A2#, 5)` and friends silently return empty or junk. Static ranges generous enough to cover the largest plausible result (e.g. `:10000`) are the safe pattern — `SUMPRODUCT`/`FILTER` ignore the empty cells. **Spilling array headers** — `SCOTT.XINVOICE`, `SCOTT.XBILL`, `SCOTT.XGLTRX`, `SCOTT.XPAYMENT` spill **data only** (no header row from the function itself); write column headers manually in the row above the anchor if you want them visible.

---

## Table of Contents

1. [Account Info Functions](#1-account-info) — DESC, CHART, LIST
2. [GL Balance Functions](#2-gl-balance) — GL, NGL, GLMULTI, RANGE
3. [Budget Functions](#3-budget) — BUDGET, BUDGETMULTI, BUDGETR
4. [Tracking Category Functions (Xero)](#4-tracking) — XTRACK, XTRACKR, XTRACKM, GLMULTITRACK
5. [Class Functions (QuickBooks)](#5-class-qb) — QCLASS
6. [Net Income](#6-net-income) — XNI
7. [Invoice / AR Functions](#7-invoices) — XINVOICE, XINVOICE_HEADER, XINVOICE_SUM
8. [Bill / AP Functions](#8-bills) — XBILL, XBILL_HEADER, XBILL_SUM
9. [Transaction & Payment Detail](#9-transactions) — XGLTRX, XPAYMENT

---

## 1. Account Info Functions {#1-account-info}

### SCOTT.DESC
Returns the GL account description/name for a given account code.

```
=SCOTT.DESC(org_id, account_code)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | account_code | Xero GL account code | `"4301"` |

**Example:** `=SCOTT.DESC(112353, "6100")` → returns `"Advertising"`

---

### SCOTT.CHART
Returns the full Chart of Accounts as a spilling array.

```
=SCOTT.CHART(org_id, [include_archived])
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | include_archived | TRUE to include archived accounts | `TRUE` |

**Returns** — spilling array, 3 columns:

| Column | Header | Description |
|--------|--------|-------------|
| 1 | Code | Account code (string) |
| 2 | Name | Account name |
| 3 | Type | Xero account type (BANK, CURRENT, REVENUE, EXPENSE, etc.) |

**Note**: Ensure enough empty cells below and to the right of the anchor.

**Example:** `=SCOTT.CHART(112353, TRUE)`

---

### SCOTT.LIST
Returns a list of the specified type as a spilling array. Works for both Xero and QuickBooks orgs, with some `type` values shared and others system-specific.

```
=SCOTT.LIST(org_id, type, [include_archived])
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | type | See table below — values vary by accounting system | `"contacts"` |
| 3 | include_archived | TRUE to include archived | `TRUE` |

**Available `type` values:**

| Type value | Xero | QuickBooks | Returns |
|------------|------|-----------|---------|
| `"account"` | ✓ | ✓ | List of GL accounts (similar to but simpler than `SCOTT.CHART`) |
| `"items"` | ✓ | ✓ | List of inventory / service items |
| `"contacts"` | ✓ | — | All contacts (Xero treats customers + vendors as one list) |
| `"track"` | ✓ | — | Tracking categories with options enumerated |
| `"tracking"` | ✓ | — | **Same as `"track"`** — accepted alias |
| `"category"` | ✓ | — | **Same as `"track"`** — accepted alias |
| `"class"` | — | ✓ | QuickBooks classes |
| `"department"` | — | ✓ | QuickBooks departments |
| `"location"` | — | ✓ | QuickBooks locations |
| `"customer"` | — | ✓ | QuickBooks customers |
| `"vendor"` | — | ✓ | QuickBooks vendors |

**Aliases — keep all of them.** `"track"`, `"tracking"`, and `"category"` all return the same Xero tracking-categories list. Different users reach for different words; the function accepts all three and the reference preserves them all so anyone looking up their preferred term finds it documented.

**Use `"track"` for Xero tracking-category discovery; use `"class"` for the QuickBooks equivalent** when building tracking-style reports — those are the dropdown sources for `SCOTT.XTRACK` (Xero) and `SCOTT.QCLASS` (QB) respectively. (`"track"` is the canonical form to use in skill-generated formulas, but the function will accept `"tracking"` or `"category"` equally if the user supplies one of those.)

**Returns** — spilling array; columns depend on the `type` parameter:

For `type="contacts"`:

| Column | Header |
|--------|--------|
| 1 | Name |
| 2 | First Name |
| 3 | Last Name |
| 4 | Email Address |
| 5 | Contact Groups (comma-separated) |

(Additional tracking-related columns may be appended depending on the org's tracking setup.)

For `type="track"`:

| Column | Header |
|--------|--------|
| 1 | Category |
| 2 | Option |

One row per tracking option, with the category name repeated for each of its options. Use this to dynamically discover the tracking categories and options available in the org — useful for building departmental P&Ls without hardcoding department names.

**Example:** `=SCOTT.LIST(112353, "contacts")` → spilling list of all contacts
**Example:** `=SCOTT.LIST(112353, "track")` → list of tracking categories and options

---

## 2. GL Balance Functions {#2-gl-balance}

### SCOTT.GL
Returns the GL account balance for a single account over a date range.

```
=SCOTT.GL(org_id, account_code, begin_date, end_date)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | account_code | Xero GL account code | `"6100"` |
| 3 | begin_date | Start date of period | `DATE(2025,1,1)` |
| 4 | end_date | End date of period | `DATE(2025,1,31)` |

**Returns**: Single numeric value (the balance for that period).

**Example:** `=SCOTT.GL(112353, "6100", DATE(2025,1,1), DATE(2025,3,31))`

---

### SCOTT.NGL
Returns the **negated** GL balance (multiplied by -1).

```
=SCOTT.NGL(org_id, account_code, begin_date, end_date)
```

Parameters are identical to `SCOTT.GL`.

**Example:** `=SCOTT.NGL(112353, "4000", DATE(2025,1,1), DATE(2025,12,31))` → positive revenue figure

---

### SCOTT.GLMULTI
Returns the combined GL balance for **multiple specific accounts**.

```
=SCOTT.GLMULTI(org_id, begin_date, end_date, account1, account2, ...)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | begin_date | Start date | `DATE(2025,1,1)` |
| 3 | end_date | End date | `DATE(2025,12,31)` |
| 4+ | account_codes | One or more GL account codes | `"4301"`, `"4302"`, `"4303"` |

**Returns**: Single numeric value (sum of all listed accounts).

**Example (hardcoded accounts):**
```
=SCOTT.GLMULTI(112353, DATE(2025,1,1), DATE(2025,12,31), "4301", "4302", "4303")
```

**Example (dynamic with TEXTSPLIT — accounts as comma-separated text in A1):**
```
=SCOTT.GLMULTI(112353, DATE(2025,1,1), DATE(2025,12,31), TEXTSPLIT(A1, ", "))
```

**Example (dynamic with FILTER — accounts from a table):**
```
=SCOTT.GLMULTI(112353, DATE(2025,1,1), DATE(2025,12,31),
    FILTER(tblAccounts[Code], tblAccounts[Group]="Revenue", 0))
```

---

### SCOTT.RANGE
Returns the combined GL balance for a **contiguous range** of account codes.

```
=SCOTT.RANGE(org_id, start_code, end_code, begin_date, end_date, [include_archived])
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | start_code | First account code in range | `"4000"` |
| 3 | end_code | Last account code in range | `"4999"` |
| 4 | begin_date | Start date | `DATE(2025,1,1)` |
| 5 | end_date | End date | `DATE(2025,12,31)` |
| 6 | include_archived | TRUE to include archived accounts | `TRUE` |

**Example:** `=SCOTT.RANGE(112353, "4000", "4999", DATE(2025,1,1), DATE(2025,12,31))`
→ Sum of all accounts from 4000 through 4999 (total revenue)

---

## 3. Budget Functions {#3-budget}

### SCOTT.BUDGET
Returns the budget amount for a **single account**.

```
=SCOTT.BUDGET(org_id, budget_name, account_code, begin_date, end_date)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | budget_name | Name of the budget in Xero | `"FY2025 Budget"` |
| 3 | account_code | GL account code | `"6100"` |
| 4 | begin_date | Start date | `DATE(2025,1,1)` |
| 5 | end_date | End date | `DATE(2025,1,31)` |

**Example:** `=SCOTT.BUDGET(112353, "FY2025 Budget", "6100", DATE(2025,1,1), DATE(2025,1,31))`

---

### SCOTT.BUDGETMULTI
Returns the combined budget for **multiple specific accounts**.

```
=SCOTT.BUDGETMULTI(org_id, budget_name, begin_date, end_date, account1, account2, ...)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | budget_name | Budget name in Xero | `"FY2025 Budget"` |
| 3 | begin_date | Start date | `DATE(2025,1,1)` |
| 4 | end_date | End date | `DATE(2025,12,31)` |
| 5+ | account_codes | GL account codes | `"6100"`, `"6200"` |

**Example:** `=SCOTT.BUDGETMULTI(112353, "FY2025 Budget", DATE(2025,1,1), DATE(2025,12,31), "6100", "6200", "6300")`

---

### SCOTT.BUDGETR
Returns the budget for a **range of accounts**.

```
=SCOTT.BUDGETR(org_id, budget_name, start_code, end_code, begin_date, end_date)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | budget_name | Budget name in Xero | `"FY2025 Budget"` |
| 3 | start_code | First account in range | `"6000"` |
| 4 | end_code | Last account in range | `"6999"` |
| 5 | begin_date | Start date | `DATE(2025,1,1)` |
| 6 | end_date | End date | `DATE(2025,12,31)` |

**Example:** `=SCOTT.BUDGETR(112353, "FY2025 Budget", "6000", "6999", DATE(2025,1,1), DATE(2025,12,31))`

---

## 4. Tracking Category Functions (Xero) {#4-tracking}

Tracking categories in Xero are dimensions like Department, Cost Center, Location, or Project. They have a two-level structure: each **category** (e.g., "Department") has multiple **options** (e.g., "Sales", "Engineering").

The functions in this section are **Xero-only**. For QuickBooks orgs, use [SCOTT.QCLASS](#5-class-qb) instead.

### SCOTT.XTRACK
Returns GL balance for a **single account** filtered by **one tracking category**.

```
=SCOTT.XTRACK(org_id, account_code, tracking_category, tracking_option, begin_date, end_date)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | account_code | GL account code | `"6100"` |
| 3 | tracking_category | Tracking category name | `"Department"` |
| 4 | tracking_option | Tracking option value | `"Engineering"` |
| 5 | begin_date | Start date | `DATE(2025,1,1)` |
| 6 | end_date | End date | `DATE(2025,12,31)` |

**Tip**: To find transactions with NO tracking assigned, leave tracking_option empty:
`=SCOTT.XTRACK(112353, "6100", , "", DATE(2025,1,1), DATE(2025,12,31))`

**Example:** `=SCOTT.XTRACK(112353, "6100", "Department", "Sales", DATE(2025,1,1), DATE(2025,3,31))`

---

### SCOTT.XTRACKR
Returns GL balance for a **range of accounts** filtered by **one tracking category**.

```
=SCOTT.XTRACKR(org_id, start_code, end_code, tracking_category, tracking_option, begin_date, end_date, [include_archived])
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | start_code | First account code | `"6000"` |
| 3 | end_code | Last account code | `"6999"` |
| 4 | tracking_category | Tracking category name | `"Department"` |
| 5 | tracking_option | Option value | `"Engineering"` |
| 6 | begin_date | Start date | `DATE(2025,1,1)` |
| 7 | end_date | End date | `DATE(2025,12,31)` |
| 8 | include_archived | Include archived accounts | `TRUE` |

**Example:** `=SCOTT.XTRACKR(112353, "6000", "6999", "Department", "Engineering", DATE(2025,1,1), DATE(2025,12,31))`

---

### SCOTT.XTRACKM
Returns GL balance for a **single account** filtered by **two tracking categories**.

```
=SCOTT.XTRACKM(org_id, account_code, track_cat_1, track_opt_1, track_cat_2, track_opt_2, begin_date, end_date)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | account_code | GL account code | `"6100"` |
| 3 | track_cat_1 | First tracking category | `"Department"` |
| 4 | track_opt_1 | First tracking option | `"Engineering"` |
| 5 | track_cat_2 | Second tracking category | `"Project"` |
| 6 | track_opt_2 | Second tracking option | `"Alpha"` |
| 7 | begin_date | Start date | `DATE(2025,1,1)` |
| 8 | end_date | End date | `DATE(2025,12,31)` |

**Example:** `=SCOTT.XTRACKM(112353, "6100", "Department", "Engineering", "Project", "Alpha", DATE(2025,1,1), DATE(2025,12,31))`

---

### SCOTT.GLMULTITRACK
Returns the combined GL balance for **multiple accounts** filtered by **one tracking category**.

```
=SCOTT.GLMULTITRACK(org_id, tracking_category, tracking_option, begin_date, end_date, account1, account2, ...)
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | tracking_category | Tracking category name | `"Department"` |
| 3 | tracking_option | Option value | `"Sales"` |
| 4 | begin_date | Start date | `DATE(2025,1,1)` |
| 5 | end_date | End date | `DATE(2025,12,31)` |
| 6+ | account_codes | GL account codes | `"6100"`, `"6200"` |

**Example:** `=SCOTT.GLMULTITRACK(112353, "Department", "Sales", DATE(2025,1,1), DATE(2025,12,31), "6100", "6200", "6300")`

---

## 5. Class Functions (QuickBooks) {#5-class-qb}

QuickBooks uses **classes** as its dimensional tracking concept, equivalent in spirit to Xero's tracking categories but flatter — classes are a single-level hierarchy (with optional sub-classes), not a category-plus-option two-level structure.

The functions in [section 4](#4-tracking) (XTRACK, XTRACKR, XTRACKM, GLMULTITRACK) are **Xero-only** and won't return meaningful results for a QuickBooks org. Use the function below instead.

### SCOTT.QCLASS
Returns GL balance for a **single account** filtered by a **QuickBooks class**.

```
=SCOTT.QCLASS(org_id, account_code, class_name, begin_date, end_date, [include_sub_classes])
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | account_code | GL account code | `"6100"` |
| 3 | class_name | QuickBooks class name | `"Retail"` |
| 4 | begin_date | Start date | `DATE(2025,1,1)` |
| 5 | end_date | End date | `DATE(2025,12,31)` |
| 6 | include_sub_classes | (Optional) `TRUE` to roll up activity from any sub-classes under this class; default behavior includes only the class itself | `TRUE` |

**Returns**: Single numeric value (the balance for the account in that class).

**Example (single class):**
```
=SCOTT.QCLASS(112353, "6100", "Retail", DATE(2025,1,1), DATE(2025,12,31))
```

**Example (class with sub-classes rolled up):**
```
=SCOTT.QCLASS(112353, "6100", "Retail", DATE(2025,1,1), DATE(2025,12,31), TRUE)
```

**System note**: This is a **QuickBooks-only** function. For Xero orgs, use `SCOTT.XTRACK` (section 4) instead — those use a (category, option) pair instead of a single class name.

---

## 6. Net Income {#6-net-income}

### SCOTT.XNI
Returns net income, optionally filtered by up to two tracking categories.

```
=SCOTT.XNI(org_id, begin_date, end_date, [track_cat_1], [track_opt_1], [track_cat_2], [track_opt_2])
```

| # | Parameter | Description | Example |
|---|-----------|-------------|---------|
| 1 | org_id | Scott's Org ID | `112353` |
| 2 | begin_date | Start date | `DATE(2025,1,1)` |
| 3 | end_date | End date | `DATE(2025,12,31)` |
| 4 | track_cat_1 | (Optional) First tracking category | `"Department"` |
| 5 | track_opt_1 | (Optional) First tracking option | `"Sales"` |
| 6 | track_cat_2 | (Optional) Second tracking category | |
| 7 | track_opt_2 | (Optional) Second tracking option | |

**Example (total):** `=SCOTT.XNI(112353, DATE(2025,1,1), DATE(2025,12,31))`
**Example (by dept):** `=SCOTT.XNI(112353, DATE(2025,1,1), DATE(2025,12,31), "Department", "Engineering")`

---

## 7. Invoice / AR Functions {#7-invoices}

### SCOTT.XINVOICE
Returns **line-item detail** of invoices as a spilling array.

```
=SCOTT.XINVOICE(org_id, begin_date, end_date, [begin_due_date], [end_due_date], [item], [gl_account], [track_cat], [track_opt], [track_cat_2], [track_opt_2], [contact1], [contact2], ...)
```

| # | Parameter | Description |
|---|-----------|-------------|
| 1 | org_id | Scott's Org ID |
| 2 | begin_date | Invoice date range start |
| 3 | end_date | Invoice date range end |
| 4 | begin_due_date | (Optional) Due date range start |
| 5 | end_due_date | (Optional) Due date range end |
| 6 | item | (Optional) Filter by item name |
| 7 | gl_account | (Optional) Filter by GL account |
| 8 | track_cat | (Optional) Tracking category |
| 9 | track_opt | (Optional) Tracking option |
| 10 | track_cat_2 | (Optional) Second tracking category |
| 11 | track_opt_2 | (Optional) Second tracking option |
| 12+ | contacts | (Optional) Filter to specific contact(s) |

**Returns** — spilling array, 20 columns:

| Column | Header |
|--------|--------|
| 1 | Contact |
| 2 | Item |
| 3 | Description |
| 4 | Unit Amount |
| 5 | Quantity |
| 6 | Total Tax |
| 7 | Total |
| 8 | Invoice Date *(ISO string, e.g. "2026-04-01")* |
| 9 | Due Date *(ISO string)* |
| 10 | Invoice Number |
| 11 | Type |
| 12 | Status (`AUTHORISED`, `Paid`, `Awaiting Payment`, `VOIDED`, `DRAFT`) |
| 13 | Reference |
| 14 | GL Account |
| 15 | Tracking Category 1 |
| 16 | Tracking Option 1 |
| 17 | Tracking Category 2 |
| 18 | Tracking Option 2 |
| 19 | Currency Code |
| 20 | Currency Rate |

**Example (all invoices in Q1):**
```
=SCOTT.XINVOICE(112353, DATE(2025,1,1), DATE(2025,3,31))
```

**Example (filtered to a contact):**
```
=SCOTT.XINVOICE(112353, DATE(2025,1,1), DATE(2025,3,31),,,,,,,,, "Acme Corp")
```

---

### SCOTT.XINVOICE_HEADER
Returns **invoice-level headers** (not line items) as a spilling array.

```
=SCOTT.XINVOICE_HEADER(org_id, begin_date, end_date, [status], [begin_due_date], [end_due_date], [begin_expected_pay_date], [end_expected_pay_date], [contact_group], [contact1], [contact2], ...)
```

| # | Parameter | Description |
|---|-----------|-------------|
| 1 | org_id | Scott's Org ID |
| 2 | begin_date | Invoice date range start |
| 3 | end_date | Invoice date range end |
| 4 | status | (Optional) e.g. `"AUTHORISED"`, `"PAID"`, `"VOIDED"` |
| 5 | begin_due_date | (Optional) Due date start |
| 6 | end_due_date | (Optional) Due date end |
| 7 | begin_expected_pay | (Optional) Expected payment date start |
| 8 | end_expected_pay | (Optional) Expected payment date end |
| 9 | contact_group | (Optional) Contact group name |
| 10+ | contacts | (Optional) Specific contact names |

**Returns** — spilling array, 9 columns:

| Column | Header |
|--------|--------|
| 1 | Invoice Number |
| 2 | Contact |
| 3 | Total |
| 4 | Amount Due |
| 5 | Invoice Date *(ISO string)* |
| 6 | Due Date *(ISO string)* |
| 7 | Status |
| 8 | Expected Payment Date *(ISO string)* |
| 9 | SubTotal |

**Example:** `=SCOTT.XINVOICE_HEADER(112353, DATE(2025,1,1), DATE(2025,3,31), "AUTHORISED")`

---

### SCOTT.XINVOICE_SUM
Returns the **sum of Amount Due** for invoices matching the filter criteria.

```
=SCOTT.XINVOICE_SUM(org_id, begin_date, end_date, [status], [begin_due_date], [end_due_date], [contact_group], [contact1], [contact2], ...)
```

| # | Parameter | Description |
|---|-----------|-------------|
| 1 | org_id | Scott's Org ID |
| 2 | begin_date | Invoice date range start |
| 3 | end_date | Invoice date range end |
| 4 | status | (Optional) e.g. `"AUTHORISED"` |
| 5 | begin_due_date | (Optional) Due date start |
| 6 | end_due_date | (Optional) Due date end |
| 7 | contact_group | (Optional) Contact group |
| 8+ | contacts | (Optional) Specific contacts |

**Returns**: Single numeric value.

**Example (total outstanding AR):**
```
=SCOTT.XINVOICE_SUM(112353, DATE(2020,1,1), DATE(2025,12,31), "AUTHORISED")
```

**Example (aging bucket — invoices dated in a specific window):**
```
=SCOTT.XINVOICE_SUM(112353, DATE(2025,2,1), DATE(2025,2,28), "AUTHORISED")
```

---

## 8. Bill / AP Functions {#8-bills}

### SCOTT.XBILL
Returns **line-item detail** of bills as a spilling array. Parameters mirror `SCOTT.XINVOICE`.

```
=SCOTT.XBILL(org_id, begin_date, end_date, [begin_due_date], [end_due_date], [item], [gl_account], [track_cat], [track_opt], [track_cat_2], [track_opt_2], [contact1], [contact2], ...)
```

**Returns** — spilling array, 19 columns (same shape as `XINVOICE` minus the Invoice Number column):

| Column | Header |
|--------|--------|
| 1 | Contact |
| 2 | Item |
| 3 | Description |
| 4 | Unit Amount |
| 5 | Quantity |
| 6 | Total Tax |
| 7 | Total |
| 8 | Invoice Date *(ISO string)* |
| 9 | Due Date *(ISO string)* |
| 10 | Reference |
| 11 | Type |
| 12 | Status |
| 13 | GL Account |
| 14 | Tracking Category 1 |
| 15 | Tracking Option 1 |
| 16 | Tracking Category 2 |
| 17 | Tracking Option 2 |
| 18 | Currency Code |
| 19 | Currency Rate |

**Example:** `=SCOTT.XBILL(112353, DATE(2025,1,1), DATE(2025,3,31))`

---

### SCOTT.XBILL_HEADER
Returns **bill-level headers** as a spilling array. Parameters mirror `SCOTT.XINVOICE_HEADER`.

```
=SCOTT.XBILL_HEADER(org_id, begin_date, end_date, [status], [begin_due_date], [end_due_date], [begin_planned_pay_date], [end_planned_pay_date], [contact_group], [contact1], [contact2], ...)
```

**Returns** — spilling array, 9 columns:

| Column | Header |
|--------|--------|
| 1 | Reference |
| 2 | Contact |
| 3 | Total |
| 4 | Amount Due |
| 5 | Invoice Date *(ISO string)* |
| 6 | Due Date *(ISO string)* |
| 7 | Status |
| 8 | Planned Payment Date *(ISO string)* |
| 9 | SubTotal |

**Example:** `=SCOTT.XBILL_HEADER(112353, DATE(2025,1,1), DATE(2025,3,31), "AUTHORISED")`

---

### SCOTT.XBILL_SUM
Returns the **sum of Amount Due** for bills matching the filter criteria.

```
=SCOTT.XBILL_SUM(org_id, begin_date, end_date, [status], [begin_due_date], [end_due_date], [contact_group], [contact1], [contact2], ...)
```

**Returns**: Single numeric value.

**Example (total outstanding AP):**
```
=SCOTT.XBILL_SUM(112353, DATE(2020,1,1), DATE(2025,12,31), "AUTHORISED")
```

---

## 9. Transaction & Payment Detail {#9-transactions}

### SCOTT.XGLTRX
Returns **GL transaction line detail** as a spilling array.

```
=SCOTT.XGLTRX(org_id, begin_date, end_date, [track_cat], [track_opt], [track_cat_2], [track_opt_2], [account1], [account2], ...)
```

| # | Parameter | Description |
|---|-----------|-------------|
| 1 | org_id | Scott's Org ID |
| 2 | begin_date | Transaction date start |
| 3 | end_date | Transaction date end |
| 4 | track_cat | (Optional) Tracking category |
| 5 | track_opt | (Optional) Tracking option |
| 6 | track_cat_2 | (Optional) Second tracking category |
| 7 | track_opt_2 | (Optional) Second tracking option |
| 8+ | accounts | (Optional) Filter to specific GL accounts |

**Returns** — spilling array, 14 columns:

| Column | Header |
|--------|--------|
| 1 | Date *(ISO string)* |
| 2 | Source |
| 3 | Account Code |
| 4 | Account Name |
| 5 | Debit |
| 6 | Credit |
| 7 | Description |
| 8 | Reference |
| 9 | Tracking Category 1 |
| 10 | Tracking Option 1 |
| 11 | Tracking Category 2 |
| 12 | Tracking Option 2 |
| 13 | Contact |
| 14 | Journal Number |

**Example (all transactions for an account):**
```
=SCOTT.XGLTRX(112353, DATE(2025,1,1), DATE(2025,3,31),,,,,  "6100", "6200")
```

**Example (transactions by department):**
```
=SCOTT.XGLTRX(112353, DATE(2025,1,1), DATE(2025,3,31), "Department", "Engineering",,, "6100")
```

---

### SCOTT.XPAYMENT
Returns **payment detail** as a spilling array.

```
=SCOTT.XPAYMENT(org_id, begin_date, end_date, [status], [gl_account], [contact1], [contact2], ...)
```

| # | Parameter | Description |
|---|-----------|-------------|
| 1 | org_id | Scott's Org ID |
| 2 | begin_date | Payment date start |
| 3 | end_date | Payment date end |
| 4 | status | (Optional) Payment status |
| 5 | gl_account | (Optional) Bank/GL account |
| 6+ | contacts | (Optional) Filter to specific contacts |

**Returns** — spilling array, 10 columns:

| Column | Header |
|--------|--------|
| 1 | Contact |
| 2 | Reference |
| 3 | Amount |
| 4 | Payment Date *(ISO string)* |
| 5 | Invoice Number |
| 6 | Type (`Invoice Payment`, `Bill Payment`) |
| 7 | Status |
| 8 | GL Account |
| 9 | Currency Code |
| 10 | Currency Rate |

**Example:** `=SCOTT.XPAYMENT(112353, DATE(2025,1,1), DATE(2025,3,31))`

---

## Quick Reference: Default Settings

| Setting | Value |
|---------|-------|
| Scott's Org ID | `112353` |
| Fiscal year start | January 1 |
| Date format in formulas | `DATE(YYYY,M,D)` or cell reference |
