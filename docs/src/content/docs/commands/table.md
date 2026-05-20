---
title: "jx:table"
description: Create structured Excel tables with auto-filter, banded rows, and named styles — from a template command.
---

`jx:table` creates a proper Excel structured table (ListObject) over the output range. You get auto-filter headers, banded row styling, total rows, and named table references — all without writing a line of styling code.

## Syntax

```
jx:table(lastCell="D1" name="EmployeeTable" style="TableStyleMedium9")
jx:table(lastCell="D1" name="SalesData" style="TableStyleLight1" showHeaderRow="true" showTotalRow="true")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the table area |
| `name` | Yes | Table name (must be unique in the workbook, no spaces) |
| `style` | No | Excel table style name (default: `TableStyleMedium9`) |
| `showHeaderRow` | No | Show the header row with filter dropdowns (default: `true`) |
| `showTotalRow` | No | Add a total row at the bottom (default: `false`) |
| `showFirstColumn` | No | Highlight the first column (default: `false`) |
| `showLastColumn` | No | Highlight the last column (default: `false`) |
| `showRowStripes` | No | Enable banded row shading (default: `true`) |
| `showColumnStripes` | No | Enable banded column shading (default: `false`) |

## When to use jx:table

Use it when your output needs to behave like a **real Excel table**, not just formatted cells:

- Auto-filter dropdowns on every column header
- Structured references in formulas (e.g., `=SUM(SalesData[Amount])`)
- Banded row styling that adapts as rows are added/removed
- Total rows with built-in aggregate functions
- Named table references for pivot tables and Power Query

## Example: Employee table with auto-filter

```
Cell A1 comment:
  jx:area(lastCell="D10")
  jx:each(items="employees" var="e" lastCell="D1")

Cell A1 also has:
  jx:table(lastCell="D1" name="Employees" style="TableStyleMedium2")
```

Cell values: `${e.Name}` | `${e.Department}` | `${e.Title}` | `${e.Salary}`

The output is a fully functional Excel table. Users can click any column header to filter or sort — no macros, no VBA.

## Example: Sales report with totals

```
Cell A1 comment:
  jx:table(lastCell="C1" name="MonthlySales" style="TableStyleMedium9" showTotalRow="true")
```

With `showTotalRow="true"`, Excel adds a total row at the bottom of the table. Users can click each cell in the total row to choose an aggregate function (Sum, Average, Count, etc.).

## Available styles

Excel ships with 60+ built-in table styles. Common ones:

| Style | Appearance |
|-------|------------|
| `TableStyleLight1` through `TableStyleLight21` | Subtle, professional |
| `TableStyleMedium1` through `TableStyleMedium28` | Balanced color and structure |
| `TableStyleDark1` through `TableStyleDark11` | Bold, high-contrast |

Pick a style in Excel's table design ribbon, note its name, use it in the `style` attribute.

## Deferred execution

`jx:table` uses deferred execution. The table definition is collected during processing and applied after all rows are written — ensuring the table range covers the correct number of output rows.

## Common pitfalls

- **Table names must be unique workbook-wide.** Two `jx:table` commands with the same `name` produce a file Excel refuses to open.
- **No spaces in `name`.** Excel rejects table names with spaces. Use `EmployeeTable` or `employee_table`, not `Employee Table`.
- **Tables can't span sheets.** A single `jx:table` covers one rectangular range on one sheet. Multi-sheet reports need a separate table per sheet.
- **Empty tables aren't allowed.** If your `jx:each` produces zero rows, the table command has no rows to wrap. Combine with `jx:if` to skip table creation when the data set is empty.

## What's next?

Add visual data indicators with conditional formatting:

**[jx:conditionalFormat &rarr;](/xlfill/commands/conditionalformat/)**
