# HR OT Sheet — Module Specification

**Module Name:** `hr_ot_sheet`
**Version:** 1.1.0
**Odoo Version:** 17
**Author:** Chandika Rathnayake
**License:** AGPL-3
**Category:** Human Resources / Payroll

---

## 1. Purpose

This module enables HR teams to bulk-import overtime (normal and holiday) and late deduction hours from an Excel spreadsheet and apply them as payslip inputs for a given payroll period. It bridges an external attendance/OT register with Odoo's payroll engine.

---

## 2. Dependencies

| Dependency | Purpose |
|---|---|
| `hr` | Employee and contract records |
| `hr_payroll` | Payslip, payslip inputs, salary structures |
| `openpyxl` (Python) | Reading `.xlsx` files — declared in `external_dependencies` |

### Employee Rate Fields

`x_studio_ot_rate` and `x_studio_employee_rate` are defined directly by this module on `hr.employee`. No external Studio setup is required.

| Field | Type | Used For |
|---|---|---|
| `x_studio_ot_rate` | Float | Rate per OT hour |
| `x_studio_employee_rate` | Float | Rate per hour for late deductions |

---

## 3. Module Structure

```
hr_ot_sheet/
├── __init__.py
├── __manifest__.py
├── models/
│   ├── __init__.py
│   ├── hr_employee.py       # x_studio_ot_rate, x_studio_employee_rate
│   └── ot_sheet.py          # HrOtSheet, HrOtSheetLine
├── wizards/
│   └── __init__.py          # (empty — wizard removed)
├── views/
│   └── hr_ot_sheet_views.xml
├── data/
│   └── ot_input_codes.xml   # Payslip input types + sequence
└── security/
    └── ir.model.access.csv
```

---

## 4. Data Models

### 4.1 `hr.ot.sheet` — OT Sheet Header

Represents one OT processing batch for a given month/year.

| Field | Type | Description |
|---|---|---|
| `name` | Char | Auto-generated reference — format `{seq}/{mm}/{yyyy}` |
| `month` | Selection (1–12) | Payroll month |
| `year` | Integer | Payroll year (default: current year) |
| `state` | Selection | `draft` (default) or `done` |
| `line_ids` | One2many → `hr.ot.sheet.line` | Employee OT lines |
| `import_file` | Binary | Uploaded Excel file |
| `import_filename` | Char | File name of the uploaded Excel |

**Sequence:** `hr.ot.sheet` — 3-digit padded, e.g. `001/03/2026`.

---

### 4.2 `hr.ot.sheet.line` — OT Line

One row per employee within an OT Sheet.

| Field | Type | Description |
|---|---|---|
| `sheet_id` | Many2one → `hr.ot.sheet` | Parent sheet (cascade delete) |
| `employee_id` | Many2one → `hr.employee` | Employee (required) |
| `ot_normal_hrs` | Float | Normal OT hours |
| `ot_holiday_hrs` | Float | Holiday OT hours |
| `late_ded_hrs` | Float | Late deduction hours |
| `ot_rate` | Float (related) | `employee.x_studio_ot_rate` (read-only) |
| `employee_rate` | Float (related) | `employee.x_studio_employee_rate` (read-only) |
| `ot_normal` | Monetary (computed) | `ot_normal_hrs × ot_rate` |
| `ot_holiday` | Monetary (computed) | `ot_holiday_hrs × ot_rate` |
| `late_deduction` | Monetary (computed) | `late_ded_hrs × employee_rate` |
| `company_currency` | Many2one (related) | Employee company currency |
| `description` | Char (computed) | Human-readable summary of this line |
| `applied` | Boolean | `True` once pushed to a payslip |

**Computed description format:**
`2.5 OT hours @ 250.00 | 1.0 Holiday OT hours @ 250.00 | 0.5 Late hours @ 200.00`
Components with 0 hours are omitted.

---

## 5. Payslip Input Types

Defined in `data/ot_input_codes.xml` and created at module install:

| XML ID | Name | Code | Effect on Payslip |
|---|---|---|---|
| `input_type_ot_normal` | OT Normal | `OT_NORMAL` | Positive addition |
| `input_type_ot_holiday` | OT Holiday | `OT_HOLIDAY` | Positive addition |
| `input_type_late_deduction` | Late Deduction | `LATE_DEDUCTION` | Negative deduction |

If an input type is missing at runtime, `action_apply_to_payslips` creates it automatically.

---

## 6. User Interface

### 6.1 Menu Structure

```
Human Resources
└── OT Sheet (root menu)
    └── OT Sheets  → list/form of hr.ot.sheet
```

### 6.2 OT Sheet Form View

**Header buttons (visible in `draft` state only):**
- **Import Excel** — parses the uploaded file and populates lines.
- **Apply to Payslips** — pushes all lines to payroll; sets state to `done`.

**Status bar:** `Draft` → `Done`

**Lines tab:** Editable inline tree with all line fields. Read-only columns: `ot_rate`, `employee_rate`, computed amounts, `description`, `applied`.

**Import tab (draft only):** File upload widget.

---

## 7. Excel File Format

The import reads `.xlsx` files. **Row 1 is always treated as a header row and skipped for data.**

### Header-based column detection

The importer reads row 1 to detect column positions by name. Column order in the file does not matter as long as the headers match one of the recognised aliases below.

| Canonical Name | Accepted Header Aliases |
|---|---|
| `emp_code` | `emp_code`, `employee code`, `barcode`, `emp code`, `code` |
| `emp_name` | `emp_name`, `employee name`, `name`, `emp name` |
| `ot_normal_hrs` | `ot_normal_hrs`, `normal ot`, `ot normal`, `normal ot hours`, `ot normal hours`, `normal overtime` |
| `ot_holiday_hrs` | `ot_holiday_hrs`, `holiday ot`, `ot holiday`, `holiday ot hours`, `ot holiday hours`, `holiday overtime` |
| `late_ded_hrs` | `late_ded_hrs`, `late deduction`, `late hours`, `late ded`, `late deduction hours` |
| `description` | `description`, `desc`, `notes`, `remarks` |

If a header cell does not match any alias, the importer falls back to the legacy positional order: A=emp_code, B=emp_name, C=ot_normal_hrs, D=ot_holiday_hrs, E=late_ded_hrs, F=description.

**Employee lookup:** First attempts to match by `hr.employee.barcode` (emp_code). If not found, falls back to a case-insensitive name match on emp_name.

Rows with unresolvable employees or non-numeric hour values are skipped; a notification lists all skipped rows after the import completes.

---

## 8. Business Workflows

### 8.1 OT Sheet Lifecycle

```
[Create Sheet] → Draft
      ↓
[Upload Excel on Import tab] → Click "Import Excel" → Lines populated
      ↓
[Review & Edit Lines]
      ↓
[Apply to Payslips] → Done
```

### 8.2 Apply to Payslips Logic (`action_apply_to_payslips`)

For each `hr.ot.sheet.line`:

1. Determine the date range from the sheet's `month` and `year`.
2. Search for an existing `hr.payslip` for the employee covering that period.
3. **If payslip exists and is `done` or `paid`:** skip the line and record a warning.
4. If no payslip exists:
   - Fetch the employee's active contract. If absent, skip with a warning.
   - Determine the salary structure via `contract.structure_type_id.default_struct_id` (falls back to any available structure). If none found, skip with a warning.
   - Create a new `hr.payslip` in `draft` state.
5. Upsert payslip inputs for each non-zero component:
   - `OT_NORMAL` → `ot_normal` amount
   - `OT_HOLIDAY` → `ot_holiday` amount
   - `LATE_DEDUCTION` → `late_deduction` amount (stored as negative)
6. Mark the line `applied = True`.
7. Set sheet `state = 'done'`.
8. If any warnings were collected, display them as a sticky notification. Lines that were skipped remain `applied = False` and can be re-processed after resolving the issue.

---

## 9. Security

Access rules in `security/ir.model.access.csv`:

| Model | Group | Read | Write | Create | Delete |
|---|---|---|---|---|---|
| `hr.ot.sheet` | All internal users | ✓ | — | — | — |
| `hr.ot.sheet` | HR Payroll Manager | ✓ | ✓ | ✓ | ✓ |
| `hr.ot.sheet.line` | All internal users | ✓ | — | — | — |
| `hr.ot.sheet.line` | HR Payroll Manager | ✓ | ✓ | ✓ | ✓ |

Only members of `hr_payroll.group_hr_payroll_manager` can create, edit, or delete OT sheets and lines.
