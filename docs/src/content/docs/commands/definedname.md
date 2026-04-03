---
title: "jx:definedName"
description: Create named ranges in the output workbook — for formulas, pivot tables, and cross-sheet references.
---

`jx:definedName` creates an Excel defined name (named range) over the output area. Named ranges make formulas more readable, enable cross-sheet references, and serve as data sources for pivot tables and charts.

## Syntax

```
jx:definedName(lastCell="D1" name="SalesData")
jx:definedName(lastCell="D1" name="EmployeeList" scope="sheet")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the named range area |
| `name` | Yes | The defined name (must be a valid Excel name — no spaces, starts with letter or underscore) |
| `scope` | No | Scope: `workbook` (default) or `sheet` |

## When to use jx:definedName

Use it when your output needs **referenceable data ranges**:

- Formula references: `=SUM(SalesData)` instead of `=SUM(B2:B150)` — readable and maintainable
- Pivot table sources: point a pivot table at a named range; the range auto-sizes to your data
- Cross-sheet references: reference data from another sheet by name
- Chart data sources: charts referencing named ranges update correctly when rows change
- Data validation lists: dropdown lists that reference a named range

## Example: Named range for a data column

```
Cell A1 comment:
  jx:area(lastCell="D10")
  jx:each(items="employees" var="e" lastCell="D1")

Cell B1 comment:
  jx:definedName(lastCell="B1" name="Salaries")
```

Cell B1 value: `${e.Salary}`

After processing, the defined name "Salaries" covers B2:B{N} where N is the number of employees. Any formula referencing `=AVERAGE(Salaries)` or `=SUM(Salaries)` works correctly regardless of how many rows were generated.

## Example: Pivot table source

```
Cell A1 comment:
  jx:definedName(lastCell="D1" name="PivotSource")
  jx:each(items="transactions" var="t" lastCell="D1")
```

Create a pivot table in another sheet that references "PivotSource" as its data source. When the report is regenerated with more data, the named range expands and the pivot table picks it up.

## Deferred execution

`jx:definedName` uses deferred execution. The named range is created after all rows are written, so the range covers the correct number of output rows.

## What's next?

Compose templates from reusable fragments:

**[jx:include &rarr;](/xlfill/commands/include/)**
