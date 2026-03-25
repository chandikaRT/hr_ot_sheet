# Test Fixtures

## sample_ot.xlsx

Place a valid OT Excel file here named `sample_ot.xlsx` to enable test 4 (import test).

The file must follow the module's expected format — either positional or with recognised header names:

| emp_code | emp_name | ot_normal_hrs | ot_holiday_hrs | late_ded_hrs | description |
|---|---|---|---|---|---|
| E001 | John Doe | 8 | 4 | 0.5 | March OT |

Test 4 is automatically skipped if this file is absent.
