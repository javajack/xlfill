---
title: Built-in Functions
description: All 16 built-in functions available in XLFill expressions — text, formatting, aggregation, and i18n.
---

XLFill includes 16 built-in functions that you can use directly in `${...}` expressions. No registration needed — they're available in every template.

## Text functions

### upper(s)

Converts a string to uppercase.

```
${upper(e.Name)}          → "ALICE JOHNSON"
${upper(e.Department)}    → "ENGINEERING"
```

### lower(s)

Converts a string to lowercase.

```
${lower(e.Email)}         → "alice@example.com"
${lower(e.Status)}        → "active"
```

### title(s)

Converts a string to title case (first letter of each word capitalized).

```
${title(e.Name)}          → "Alice Johnson"
${title("hello world")}   → "Hello World"
```

### join(sep, items)

Joins a slice of values into a single string with a separator.

```
${join(", ", e.Skills)}           → "Go, Python, SQL"
${join(" | ", e.Departments)}     → "Engineering | Research"
${join("\n", e.Notes)}            → multi-line text
```

Useful for combining list data into a single cell value.

## Formatting functions

### formatNumber(value, format)

Formats a number using Go's `fmt.Sprintf`-style patterns or common presets.

```
${formatNumber(e.Salary, "#,##0.00")}     → "75,000.00"
${formatNumber(e.Rating, "0.0")}          → "4.5"
${formatNumber(e.Percentage, "0.00%")}    → "85.50%"
```

### formatDate(value, layout)

Formats a date/time value using Go's time layout syntax.

```
${formatDate(e.HireDate, "2006-01-02")}           → "2024-03-15"
${formatDate(e.HireDate, "January 2, 2006")}      → "March 15, 2024"
${formatDate(e.HireDate, "02/01/2006")}            → "15/03/2024"
${formatDate(e.CreatedAt, "2006-01-02 15:04:05")} → "2024-03-15 09:30:00"
```

Uses Go's reference time (`Mon Jan 2 15:04:05 MST 2006`) — see [Go time package](https://pkg.go.dev/time#pkg-constants) for layout options.

## Null-safe functions

### coalesce(values...)

Returns the first non-nil, non-empty value from the arguments.

```
${coalesce(e.Nickname, e.FirstName, "Unknown")}   → first non-empty value
${coalesce(e.Phone, e.Mobile, "N/A")}             → fallback chain
```

Like SQL's `COALESCE` — essential for handling optional fields without template errors.

### ifEmpty(value, fallback)

Returns `fallback` if `value` is nil, empty string, or zero value; otherwise returns `value`.

```
${ifEmpty(e.Department, "Unassigned")}    → "Engineering" or "Unassigned"
${ifEmpty(e.Notes, "-")}                  → notes text or "-"
```

Simpler than `coalesce` when you have exactly one value and one fallback.

## Aggregation functions

These functions operate on slices and are particularly useful with `jx:updateCell` for summary rows, or in any cell that needs to aggregate collection data.

### sumBy(items, field)

Sums a numeric field across a slice of items.

```
${sumBy(employees, "Salary")}     → 450000
${sumBy(orders, "Total")}        → 12750.50
```

### avgBy(items, field)

Calculates the average of a numeric field across a slice.

```
${avgBy(employees, "Salary")}     → 75000
${avgBy(reviews, "Score")}       → 4.2
```

### countBy(items, field)

Counts non-nil values of a field across a slice.

```
${countBy(employees, "Email")}    → 15 (employees with email addresses)
${countBy(orders, "ShippedDate")} → 8 (orders that have shipped)
```

### minBy(items, field)

Returns the minimum value of a numeric field across a slice.

```
${minBy(employees, "Salary")}     → 45000
${minBy(scores, "Value")}        → 2.1
```

### maxBy(items, field)

Returns the maximum value of a numeric field across a slice.

```
${maxBy(employees, "Salary")}     → 120000
${maxBy(scores, "Value")}        → 9.8
```

## Link and annotation functions

### hyperlink(url, display)

Creates a clickable hyperlink in the cell.

```
${hyperlink("https://example.com", "Visit Site")}
${hyperlink(e.ProfileURL, e.Name)}
```

The cell shows the display text and links to the URL. Works in Excel, Google Sheets, and LibreOffice.

### comment(text)

Adds a cell comment (note) to the cell.

```
${comment("Reviewed by finance team")}
${comment(e.Notes)}
```

The comment appears as a hover tooltip on the cell. Useful for adding context without cluttering the visible data.

## Internationalization

### t(key)

Translates a key using the i18n translation map provided via `WithI18n`.

```
${t("header.name")}              → "Name" (en) or "Nombre" (es) or "Name" (de)
${t("status.active")}            → "Active" or "Activo" or "Aktiv"
```

**Setup in Go:**

```go
translations := map[string]string{
    "header.name":       "Name",
    "header.department": "Department",
    "header.salary":     "Salary",
    "status.active":     "Active",
    "status.inactive":   "Inactive",
}

xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithI18n(translations),
)
```

Build one template, generate reports in any language by swapping the translation map. No separate templates per language.

## Combining functions

Functions compose naturally in expressions:

```
${upper(ifEmpty(e.Department, "UNKNOWN"))}
${formatNumber(sumBy(items, "Total"), "#,##0.00")}
${hyperlink(e.URL, title(e.Name))}
${join(", ", e.Skills) + " (" + formatNumber(len(e.Skills), "0") + " skills)"}
```

## Custom functions

Need something not covered here? Register your own functions:

```go
xlfill.Fill("template.xlsx", "report.xlsx", data,
    xlfill.WithFunction("currency", func(args ...any) (any, error) {
        amount := args[0].(float64)
        code := args[1].(string)
        return fmt.Sprintf("%s %.2f", code, amount), nil
    }),
)
```

Use in templates: `${currency(e.Salary, "USD")}` produces `USD 75000.00`.

## What's next?

See how all commands work together:

**[Commands Overview &rarr;](/xlfill/guides/commands-overview/)**

Or check the full API reference:

**[API Reference &rarr;](/xlfill/reference/api/)**
