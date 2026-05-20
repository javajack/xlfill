---
title: Built-in Functions
description: All 16 built-in functions available in XLFill expressions — text, formatting, aggregation, and i18n.
---

XLFill includes 16 built-in functions that you can use directly in `${...}` expressions. No registration needed — they're available in every template.

:::note[User data shadows built-ins]
If your data map has a key with the same name as a built-in (e.g. `data["upper"] = ...`), your value wins and the built-in is hidden inside that template. Pick unique key names for data to avoid surprises.
:::

## How to read this guide

Every function below shows:
- **Syntax** — what the call looks like in a template cell
- **Examples** — typical inputs and outputs
- (Some entries) **Walkthrough** — a paragraph showing where you'd use it in a real report

If you're new to expressions, start with [the Expressions guide](/xlfill/guides/expressions/) — it covers the grammar (`${...}`, dot access, ternary, etc.) that hosts these functions.

Looking for the big picture (commands + functions + Go code together)? The [Beginner Tutorial](/xlfill/guides/beginner-tutorial/) walks through everything end-to-end.

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

**Walkthrough — collapse skill lists into one cell**

You have an employee record where `Skills` is `[]string{"Go", "Python", "SQL"}`. You don't want three rows per employee; you want all the skills in one cell. `join(", ", e.Skills)` gives you `"Go, Python, SQL"` in a single cell, comma-separated. Combine with `formatNumber(len(e.Skills), 0)` for a count, and you can write a header like `"Skills (3)"` alongside the list.

:::tip[Non-string items just work]
`join` calls `fmt.Sprintf("%v", item)` for each element, so `${join("/", e.Tags)}` works whether `Tags` is `[]string`, `[]int`, or `[]any`.
:::

## Formatting functions

### formatNumber(value, decimals)

Formats a number with comma thousands-separators and a given number of decimal places.

```
${formatNumber(e.Salary, 0)}      → "75,000"
${formatNumber(e.Salary, 2)}      → "75,000.00"
${formatNumber(e.Rating, 1)}      → "4.5"
${formatNumber(-1234567.89, 2)}   → "-1,234,567.89"
```

`decimals` is an integer count of digits after the decimal point.

:::caution[Not an Excel format string]
`formatNumber` takes a **decimals count**, not a `"#,##0.00"`-style pattern. If you need percentages or scientific notation, do the math in your expression and add the suffix yourself:

```
${formatNumber(e.Ratio * 100, 1) + "%"}     → "85.5%"
${formatNumber(e.Ratio * 100, 0) + "%"}     → "86%"
```

For native Excel formatting (so users can change the format later), apply the format to the cell in your **template** — XLFill preserves it.
:::

### formatDate(value, layout)

Formats a date/time value using Go's time layout syntax.

```
${formatDate(e.HireDate, "2006-01-02")}           → "2024-03-15"
${formatDate(e.HireDate, "January 2, 2006")}      → "March 15, 2024"
${formatDate(e.HireDate, "02/01/2006")}            → "15/03/2024"
${formatDate(e.CreatedAt, "2006-01-02 15:04:05")} → "2024-03-15 09:30:00"
```

Uses Go's reference time (`Mon Jan 2 15:04:05 MST 2006`) — see [Go time package](https://pkg.go.dev/time#pkg-constants) for layout options.

:::caution[Go time layouts confuse newcomers]
The layout `"2006-01-02"` is **the actual date** Go uses for its reference. It's not a format placeholder — `2006` literally means "year", `01` means "month", `02` means "day", `15` is hour (24h), `04` is minutes, `05` is seconds. Memorize that and Go time layouts stop feeling weird.

Common layouts:
- `"2006-01-02"` → ISO date (`2026-05-21`)
- `"02/01/2006"` → DD/MM/YYYY (`21/05/2026`)
- `"Jan 2, 2006"` → `"May 21, 2026"`
- `"3:04 PM"` → 12-hour time (`2:30 PM`)
- `"2006-01-02 15:04:05"` → ISO datetime

If you pass a string instead of a `time.Time`, `formatDate` tries to parse it. If parsing fails, the original string is returned unchanged — so you'll see your input string in the cell, not an error.
:::

**Walkthrough — date in the report header**

Your data map has `"reportDate": time.Now()`. In your template, cell A1 reads:

```
Generated on ${formatDate(reportDate, "January 2, 2006 at 3:04 PM")}
```

Output: `Generated on May 21, 2026 at 4:18 PM`. Same template, used daily, always shows the current date in a human-friendly format.

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

:::caution[`coalesce` and `ifEmpty` treat zero differently]
This catches people out:

- **`coalesce`** skips `nil` and empty strings, but **keeps numeric `0`**. So `coalesce(0, 99)` returns `0`.
- **`ifEmpty`** treats numeric `0` as empty. So `ifEmpty(0, 99)` returns `99`.

Use `coalesce` when `0` is a legitimate value (counts, IDs, prices that can be free). Use `ifEmpty` when `0` should fall through to the fallback (optional metrics, "no rating yet").
:::

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

**Walkthrough — totals row at the bottom of a loop**

You're rendering an employee table with `jx:each(items="employees" var="e" lastCell="C5")`. Right *below* the loop area (in row 6), you want a totals row.

| Cell | Value |
|------|-------|
| A6 | "Total" |
| B6 | `${formatNumber(sumBy(employees, "Salary"), 0)}` |
| C6 | `${formatNumber(avgBy(employees, "Salary"), 0)}` |

`sumBy` and `avgBy` work on the full data slice (not just the loop iteration's `e`), so they're safe to use *outside* the loop too. Place them where you want them in the template — the engine evaluates them once at that position.

:::caution[`countBy` counts matches, not non-nil]
Despite the name, `countBy(items, field, value)` counts items where `item[field] == value`. To count non-nil values, use `len(items)` minus the count of nils, or call a custom function. This is asymmetric with `sumBy/avgBy/minBy/maxBy` which all take only `(items, field)`.
:::

:::caution[Missing fields are skipped, not errored]
If an item lacks the field you're aggregating on, it's silently skipped. So `sumBy(orders, "Discount")` on an orders slice where half lack a `Discount` field still works — but the average computed by `avgBy` excludes those items. Be intentional about whether you want missing = 0 or missing = skipped.
:::

## Link and annotation functions

### hyperlink(url, display)

Creates a clickable hyperlink in the cell.

```
${hyperlink("https://example.com", "Visit Site")}
${hyperlink(e.ProfileURL, e.Name)}
${hyperlink("#Summary!A1", "Jump to summary")}
${hyperlink("mailto:" + e.Email, e.Name)}
```

The cell shows the display text and links to the URL. Works in Excel, Google Sheets, and LibreOffice.

**Walkthrough — clickable email links per row**

Each employee row should show the name as a clickable mailto link. In cell A2:

```
${hyperlink("mailto:" + e.Email, e.Name)}
```

When users click "Alice Johnson" in the output, Excel opens their mail client. Same idea works for ticket links (`hyperlink("https://tickets.example.com/" + t.ID, "Ticket #" + t.ID)`) or cross-sheet jumps (`hyperlink("#Sheet2!A1", "Go to details")`).

:::caution[Streaming mode drops hyperlinks]
With `WithStreaming(true)`, hyperlinks are not written — the display text lands in the cell but the link is gone. The fill produces a warning you can inspect via `Filler.Warnings()`. See [Performance Tuning](/xlfill/guides/performance-tuning/) for the full streaming limitations list.
:::

### comment(text)

Adds a cell comment (note) to the cell.

```
${comment("Reviewed by finance team")}
${comment(e.Notes)}
```

The comment appears as a hover tooltip on the cell. Useful for adding context without cluttering the visible data.

## Internationalization

### t(key)

Translates a key using the i18n translation map provided via `WithI18n`. If the key isn't found, the key itself is returned (so you can spot missing translations in the output).

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
${formatNumber(sumBy(items, "Total"), 2)}
${hyperlink(e.URL, title(e.Name))}
${join(", ", e.Skills) + " (" + formatNumber(len(e.Skills), 0) + " skills)"}
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
