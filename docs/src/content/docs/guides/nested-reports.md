---
title: "Nested Loops and Hierarchical Reports"
description: "Build multi-level reports — departments containing teams containing employees — with nested jx:each commands and automatic grouping."
---

Real reports are rarely flat. Departments contain teams. Teams contain employees. Orders contain line items. XLFill handles nested data naturally with nested `jx:each` commands — the inner loop repeats inside each iteration of the outer loop, just like nested `for` loops in Go.

## Two-level nesting: departments and employees

### Template layout

```
Row 1: Report Header (static)
Row 2: ${dept.Name} — Department header
Row 3: Name | Role | Salary — Column headers for employees
Row 4: ${e.Name} | ${e.Role} | ${e.Salary} — Employee data row

Cell A1 comment:
  jx:area(lastCell="C4")

Cell A2 comment:
  jx:each(items="departments" var="dept" lastCell="C4")

Cell A4 comment:
  jx:each(items="dept.Employees" var="e" lastCell="C4")
```

### Go code

```go
data := map[string]any{
    "departments": []map[string]any{
        {
            "Name": "Engineering",
            "Employees": []map[string]any{
                {"Name": "Alice", "Role": "Senior Engineer", "Salary": 95000},
                {"Name": "Bob", "Role": "Staff Engineer", "Salary": 120000},
                {"Name": "Carol", "Role": "Junior Engineer", "Salary": 65000},
            },
        },
        {
            "Name": "Sales",
            "Employees": []map[string]any{
                {"Name": "Dave", "Role": "Account Executive", "Salary": 82000},
                {"Name": "Eve", "Role": "Sales Director", "Salary": 110000},
            },
        },
    },
}

xlfill.Fill("template.xlsx", "org_report.xlsx", data)
```

### Output

```
Report Header
Engineering
  Name    | Role              | Salary
  Alice   | Senior Engineer   | 95,000
  Bob     | Staff Engineer    | 120,000
  Carol   | Junior Engineer   | 65,000
Sales
  Name    | Role              | Salary
  Dave    | Account Executive | 82,000
  Eve     | Sales Director    | 110,000
```

The outer `jx:each` repeats the department header, column headers, *and* the inner employee loop for each department. The inner `jx:each` accesses `dept.Employees` — the current department's employee list.

> **Pro tip:** The inner loop variable reference (`dept.Employees`) uses the outer loop's variable name (`dept`). This is how XLFill knows to iterate over each department's employees, not a global list.

## Three-level nesting: company → departments → employees

```
Row 1: ${company.Name} — Company header
Row 2: ${div.Name} — Division header
Row 3: ${dept.Name} — Department header
Row 4: ${e.Name} | ${e.Role} | ${e.Salary}

Cell A1 comment:
  jx:area(lastCell="C4")
  jx:each(items="companies" var="company" lastCell="C4")

Cell A2 comment:
  jx:each(items="company.Divisions" var="div" lastCell="C4")

Cell A3 comment:
  jx:each(items="div.Departments" var="dept" lastCell="C4")

Cell A4 comment:
  jx:each(items="dept.Employees" var="e" lastCell="C4")
```

Each level references the parent's variable: `company.Divisions`, `div.Departments`, `dept.Employees`.

> **Gotcha:** Keep nesting to 3 levels maximum. Beyond that, templates become hard to reason about, and the output can be difficult for users to navigate. If you need deeper nesting, consider splitting into multiple sheets or using groupBy to flatten the hierarchy.

## jx:group for collapsible sections

Excel's row grouping lets users collapse and expand sections. Add `jx:group` to make department sections collapsible:

```
Cell A2 comment:
  jx:each(items="departments" var="dept" lastCell="C4")
  jx:group(lastCell="C4" collapsed="false")
```

In the output, each department's rows have an outline indicator. Users can click the `+` / `-` buttons to collapse or expand individual departments.

### Collapsed by default

```
jx:group(lastCell="C4" collapsed="true")
```

Sections start collapsed — useful for large reports where users only want to drill into specific departments.

## groupBy for automatic categorization

Instead of pre-nesting your data, use `groupBy` on a flat list to automatically create groups:

```
Cell A2 comment:
  jx:each(items="employees" var="e" groupBy="Department" lastCell="C2")
```

Go code with a flat list:

```go
data := map[string]any{
    "employees": []map[string]any{
        {"Name": "Alice", "Department": "Engineering", "Salary": 95000},
        {"Name": "Bob", "Department": "Sales", "Salary": 82000},
        {"Name": "Carol", "Department": "Engineering", "Salary": 65000},
        {"Name": "Dave", "Department": "Sales", "Salary": 110000},
        {"Name": "Eve", "Department": "Engineering", "Salary": 120000},
    },
}
```

XLFill groups the employees by department before iterating. The output order is Engineering (Alice, Carol, Eve) then Sales (Bob, Dave).

> **Pro tip:** `groupBy` is simpler than nested `jx:each` when your data is already flat (e.g., a SQL query result). No need to restructure your data into a hierarchy — XLFill does it for you.

### groupBy with orderBy

Sort both the groups and the items within each group:

```
jx:each(items="employees" var="e" groupBy="Department" orderBy="Salary desc" lastCell="C2")
```

Employees within each department are sorted by salary, highest first.

## Aggregation functions for subtotals

XLFill's built-in aggregation functions work perfectly with nested reports:

### sumBy for department totals

```
Department row: ${dept.Name} — Total: ${sumBy(dept.Employees, "Salary")}
```

### avgBy for averages

```
Average Salary: ${avgBy(dept.Employees, "Salary")}
```

### countBy for counts

```
Headcount: ${countBy(dept.Employees, "Name")}
```

### Complete example with subtotals

```
Row 2: ${dept.Name} | Headcount: ${countBy(dept.Employees, "Name")} | Total: ${sumBy(dept.Employees, "Salary")}
Row 3: Name | Role | Salary
Row 4: ${e.Name} | ${e.Role} | ${e.Salary}

Cell A2 comment:
  jx:each(items="departments" var="dept" lastCell="C4")
  jx:group(lastCell="C4")

Cell A4 comment:
  jx:each(items="dept.Employees" var="e" lastCell="C4")
```

Each department section has:
- A header row with the department name, headcount, and salary total
- Individual employee rows
- Collapsible grouping

All calculated from the data — no Excel formulas needed for the aggregations.

> **Did you know?** `sumBy`, `avgBy`, `countBy`, `minBy`, and `maxBy` are built-in expression functions. They work on any slice of maps or structs. No custom function registration needed.

## jx:if inside loops for conditional rows

Show or hide rows based on conditions:

```
Cell A4 comment:
  jx:each(items="dept.Employees" var="e" lastCell="C5")

Cell A5 comment:
  jx:if(condition="e.Salary > 100000" lastCell="C5")

Row 4: ${e.Name} | ${e.Role} | ${e.Salary}
Row 5: ★ High earner — flagged for review
```

Row 5 only appears for employees earning over 100,000. You can use `jx:if` to add conditional notes, warnings, or additional detail rows.

### Conditional entire sections

```
Cell A2 comment:
  jx:each(items="departments" var="dept" lastCell="C4")
  jx:if(condition="countBy(dept.Employees, 'Name') > 0" lastCell="C4")
```

Skip departments with no employees. The entire section — header, column headers, employee rows — is omitted.

## Nested loops with merged cells

For hierarchical reports with merged department headers:

```
Cell A2 comment:
  jx:each(items="departments" var="dept" lastCell="C4")
  jx:mergeCells(lastCell="C2")

Cell A2: ${dept.Name}  (this cell spans A2:C2 in the template)
```

The department name is merged across all columns, creating a clear visual section header.

## Pattern: flat data with groupBy vs. nested data

| Approach | Data structure | Template complexity | Best for |
|----------|---------------|---------------------|----------|
| `groupBy` | Flat list | Simple (one `jx:each`) | SQL results, API responses |
| Nested `jx:each` | Pre-nested hierarchy | Moderate (multiple `jx:each`) | Complex layouts, different formatting per level |

If your data comes from a SQL query or flat JSON, use `groupBy`. If you already have nested structs (e.g., a tree structure), use nested `jx:each`.

## Tips and tricks

- **Keep nesting shallow.** 2 levels is common, 3 is the practical maximum. If you need more depth, consider using multiple sheets or flattening with groupBy.

- **Use groupBy instead of nested each when possible.** It's simpler — one command instead of two, and no need to restructure your data.

- **Aggregation functions eliminate formula complexity.** Instead of `=SUBTOTAL(9, C4:C15)` with expanding ranges, use `${sumBy(dept.Employees, "Salary")}`. No range adjustment needed.

- **Combine grouping with freeze panes.** For long reports, add `jx:freezePanes` to keep the header visible while scrolling:
  ```
  jx:freezePanes(lastCell="C1")
  ```

- **Test with edge cases.** What happens when a department has zero employees? When there's only one department? Run these through `ValidateData` to catch issues early.

- **Use `jx:pageBreak` between sections for printing:**
  ```
  Cell A2 comment:
    jx:each(items="departments" var="dept" lastCell="C4")
    jx:pageBreak(lastCell="C4")
  ```
  Each department starts on a new page when printed.

## What's next?

- Generate invoices with nested line items: **[Invoice Generation &rarr;](/xlfill/use-cases/invoices/)**
- Add charts to your nested reports: **[Charts and Dashboards &rarr;](/xlfill/guides/charts-and-dashboards/)**
- Test templates with nested data in CI: **[Testing Templates in CI/CD &rarr;](/xlfill/guides/testing-templates/)**
