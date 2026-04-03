---
title: "Excel Data Validation and Interactive Forms"
description: "Create Excel templates with dropdown lists, number constraints, and input validation — all preserved when XLFill generates the output."
---

Sometimes your Excel output isn't the end of the story. Users need to fill in values, select from dropdowns, enter numbers within a range. XLFill's `jx:dataValidation` command adds Excel data validation rules that survive template expansion — every output row gets its own validation, and users see helpful dropdowns and error messages when they interact with the file.

## jx:dataValidation — dropdowns and constraints

### Dropdown list (static values)

The most common use case: a cell with a dropdown menu.

```
Cell C2 comment:
  jx:dataValidation(type="list" formula="Active,Inactive,On Leave" lastCell="C2")

Cell C2: ${e.Status}
```

When the user opens the output file, cell C2 (and every expanded row) shows a dropdown with three options: Active, Inactive, On Leave. The cell is pre-filled with the employee's current status from data, but users can change it.

### Dropdown from a cell range

For longer lists, reference a range on a hidden sheet:

```
Cell C2 comment:
  jx:dataValidation(type="list" formula="Lookups!$A$1:$A$50" lastCell="C2")
```

The "Lookups" sheet contains your list values. Users see a dropdown but can't easily see or modify the source list.

> **Pro tip:** Put your lookup data on a separate sheet and hide it with Excel's sheet visibility feature. The validation still works, but users can't accidentally edit the source data.

### Integer range

Constrain input to whole numbers within a range:

```
Cell D2 comment:
  jx:dataValidation(type="whole" operator="between" formula="1" formula2="100"
                     errorTitle="Invalid rating"
                     error="Please enter a number between 1 and 100"
                     lastCell="D2")
```

If a user types 150, Excel shows the error message. No VBA, no macros — just native Excel validation.

### Decimal range

```
Cell E2 comment:
  jx:dataValidation(type="decimal" operator="between" formula="0.0" formula2="1.0"
                     errorTitle="Invalid percentage"
                     error="Enter a value between 0 and 1"
                     lastCell="E2")
```

### Date range

```
Cell F2 comment:
  jx:dataValidation(type="date" operator="between"
                     formula="2025-01-01" formula2="2025-12-31"
                     errorTitle="Invalid date"
                     error="Date must be within 2025"
                     lastCell="F2")
```

### Custom formula

For complex validation, use an Excel formula:

```
Cell G2 comment:
  jx:dataValidation(type="custom" formula="AND(LEN(G2)>=3, LEN(G2)<=50)"
                     errorTitle="Invalid input"
                     error="Text must be 3-50 characters"
                     lastCell="G2")
```

## Validation attributes reference

| Attribute | Required | Description |
|-----------|----------|-------------|
| `type` | Yes | `list`, `whole`, `decimal`, `date`, `time`, `textLength`, `custom` |
| `formula` | Yes | Validation formula or comma-separated list values |
| `formula2` | No | Second formula (for `between` / `notBetween` operators) |
| `operator` | No | `between`, `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterOrEqual`, `lessOrEqual` |
| `errorTitle` | No | Title of the error dialog |
| `error` | No | Error message text |
| `promptTitle` | No | Title of the input prompt |
| `prompt` | No | Input prompt message (shown when cell is selected) |
| `showErrorMessage` | No | Show error dialog (default: `true`) |
| `showInputMessage` | No | Show input prompt (default: `true`) |
| `allowBlank` | No | Allow empty cells (default: `true`) |
| `lastCell` | Yes | Bottom-right cell of the validation area |

## Data validation in loops

This is where XLFill shines over manual Excel setup. When `jx:dataValidation` is inside a `jx:each` area, every expanded row gets its own validation rule:

```
Cell A1 comment:
  jx:area(lastCell="E2")
  jx:each(items="employees" var="e" lastCell="E2")

Cell A2: ${e.Name}
Cell B2: ${e.Department}
Cell C2: ${e.Status}
Cell D2: ${e.Rating}
Cell E2: (blank — user fills this in)

Cell C2 comment:
  jx:dataValidation(type="list" formula="Active,Inactive,On Leave" lastCell="C2")

Cell D2 comment:
  jx:dataValidation(type="whole" operator="between" formula="1" formula2="5"
                     promptTitle="Rating" prompt="Enter a rating from 1 to 5"
                     lastCell="D2")

Cell E2 comment:
  jx:dataValidation(type="list" formula="Approved,Rejected,Pending" lastCell="E2")
```

With 100 employees, you get 100 rows — each with a Status dropdown, a 1-5 Rating constraint, and an Approval dropdown. No manual copy-paste of validation rules.

> **Did you know?** This was one of the most requested features in JXLS (issue #46) that was never implemented. XLFill handles it natively.

## jx:protect — lock formulas, allow data entry

After you add validations, you probably want to protect the sheet so users can only edit the validated cells:

```
Cell A1 comment:
  jx:area(lastCell="E2")
  jx:each(items="employees" var="e" lastCell="E2")
  jx:protect(password="review2025" lastCell="E2")
```

By default, `jx:protect` locks all cells. To allow editing on specific cells, unlock them in the template:

1. In Excel, select the cells you want users to edit (C2, D2, E2)
2. Right-click → Format Cells → Protection → uncheck "Locked"
3. Save the template

Now the output sheet is protected: formulas and labels are locked, but the Status, Rating, and Approval columns are editable.

> **Gotcha:** Remember to unlock cells *in the template* before protecting. If you forget, every cell will be locked, and users won't be able to enter data anywhere.

## Building a data entry form

Combine validation, protection, and formatting for a complete data collection form:

### Template

```
Row 1: Headers (Name, Department, Status, Rating, Review Notes, Approval)
Row 2: ${e.Name}, ${e.Department}, ${e.Status}, ${e.Rating}, (blank), (blank)

Cell A1 comment:
  jx:area(lastCell="F2")
  jx:each(items="employees" var="e" lastCell="F2")
  jx:protect(password="hr2025" lastCell="F2")
  jx:freezePanes(lastCell="F1")

Cell C2 comment:
  jx:dataValidation(type="list" formula="Active,Inactive,On Leave,Terminated"
                     lastCell="C2")

Cell D2 comment:
  jx:dataValidation(type="whole" operator="between" formula="1" formula2="5"
                     promptTitle="Performance Rating"
                     prompt="1=Needs Improvement, 3=Meets Expectations, 5=Exceeds"
                     errorTitle="Invalid Rating"
                     error="Rating must be between 1 and 5"
                     lastCell="D2")

Cell E2 comment:
  jx:dataValidation(type="textLength" operator="lessThan" formula="500"
                     promptTitle="Review Notes"
                     prompt="Enter your review (max 500 characters)"
                     lastCell="E2")

Cell F2 comment:
  jx:dataValidation(type="list" formula="Approved,Rejected,Pending"
                     lastCell="F2")
```

### Go code

```go
data := map[string]any{
    "employees": fetchEmployeesForReview(),
}

xlfill.Fill("review_form_template.xlsx", "review_form_Q1_2025.xlsx", data)
```

The output is a locked spreadsheet where HR managers can only:
- Change the Status dropdown
- Enter a 1-5 rating
- Type review notes (max 500 chars)
- Select an approval status

Everything else — names, departments, headers — is read-only.

## Dynamic dropdown sources

For dropdowns that depend on data, pass the list values through the data map:

```go
data := map[string]any{
    "employees":   employees,
    "departments": []string{"Engineering", "Sales", "Marketing", "HR", "Operations"},
}
```

Then in the template, reference the data:

```
Cell B2 comment:
  jx:dataValidation(type="list" formula="Lookups!$A$1:$A$20" lastCell="B2")
```

Or use a `jx:each` on a separate "Lookups" sheet to populate the dropdown source dynamically.

## Combining validation with conditional formatting

Make validated cells visually informative:

```
Cell C2 comment:
  jx:dataValidation(type="list" formula="Active,Inactive,On Leave" lastCell="C2")
  jx:conditionalFormat(lastCell="C2" type="cellIs" operator="equal"
                        formula="\"Inactive\"" format="font:red,bold")
```

Inactive employees are highlighted in red. Users can still change the dropdown value, and the formatting updates automatically.

## Tips and tricks

- **Input messages are helpful.** Use `promptTitle` and `prompt` to guide users — the message appears when they select the cell, before they type anything.

- **Allow blank for optional fields.** Set `allowBlank="true"` (the default) for fields that don't require input.

- **Combine with `jx:autoColWidth`.** Dropdown cells might need wider columns to show the full list. Add `jx:autoColWidth` to the area.

- **Test your validations.** After generating the output, open it in Excel and try entering invalid data. Make sure the error messages are clear and helpful.

- **Cascade dropdowns with named ranges.** For dependent dropdowns (Country → City), use named ranges and INDIRECT formulas. Put the lookup data on a hidden sheet, define named ranges per country, and use `jx:dataValidation(type="list" formula="INDIRECT(B2)")`.

- **Use `jx:repeat` for blank entry rows.** If you need empty rows for users to add new data:
  ```
  Cell A10 comment:
    jx:repeat(count="20" lastCell="F10")
    jx:dataValidation(type="list" formula="Active,Inactive" lastCell="C10")
  ```
  Twenty blank rows, each with a dropdown in column C.

## What's next?

- Build nested hierarchical reports: **[Nested Reports &rarr;](/xlfill/guides/nested-reports/)**
- See the full data validation command reference: **[jx:dataValidation Reference &rarr;](/xlfill/commands/datavalidation/)**
- Protect sheets with passwords: **[jx:protect Reference &rarr;](/xlfill/commands/protect/)**
