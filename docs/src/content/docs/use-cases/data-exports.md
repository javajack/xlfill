---
title: "Data Export to Excel from Go APIs"
description: "Export database queries, API responses, and struct slices to Excel with auto-filter tables, formatted columns, and streaming for large datasets."
---

Not every Excel file needs to be a beautifully designed report. Sometimes you just need to get data out — a database query, an API response, a slice of structs — into an `.xlsx` file with decent formatting and an auto-filter table. XLFill makes this surprisingly fast, especially for large datasets.

## The quick path: struct slice to Excel

If you have a slice of Go structs, you're two function calls away from an Excel file:

```go
type Order struct {
    OrderID    string
    Customer   string
    Amount     float64
    Status     string
    CreatedAt  time.Time
}

orders := fetchOrders() // []Order

data := xlfill.StructSliceToData("orders", orders)
xlfill.Fill("export_template.xlsx", "orders_export.xlsx", data)
```

`StructSliceToData` converts your struct slice into the `map[string]any` format XLFill expects. Field names become template expression keys.

> **Pro tip:** `StructSliceToData` handles nested structs too. If `Order` has an `Address` field that's a struct, you can access `${o.Address.City}` in the template.

## SQLRowsToData for database queries

Skip the intermediate struct. Go straight from a database query to Excel:

```go
rows, err := db.Query(`
    SELECT order_id, customer_name, amount, status, created_at
    FROM orders
    WHERE created_at >= $1
    ORDER BY created_at DESC
`, startDate)
if err != nil {
    log.Fatal(err)
}

data, err := xlfill.SQLRowsToData("orders", rows)
if err != nil {
    log.Fatal(err)
}

xlfill.Fill("export_template.xlsx", "orders_export.xlsx", data)
```

Column names from the SQL query become field names in the data map. Your template uses `${o.order_id}`, `${o.customer_name}`, etc.

> **Did you know?** `SQLRowsToData` handles all standard Go SQL types — `string`, `int64`, `float64`, `time.Time`, `bool`, and nullable types (`sql.NullString`, etc.). Null values become `nil` in the data map.

## JSONToData for API responses

If your data comes from a REST API or JSON file:

```go
resp, _ := http.Get("https://api.example.com/orders")
body, _ := io.ReadAll(resp.Body)

data, err := xlfill.JSONToData(body)
// JSON: {"orders": [{"orderId": "ORD-001", ...}, ...]}

xlfill.Fill("export_template.xlsx", "orders_export.xlsx", data)
```

The JSON structure maps directly to the template. Arrays become slices for `jx:each`, objects become maps for dot-notation access.

## jx:table for auto-filter and banding

A data export without an auto-filter table is just a wall of text. Add `jx:table` to get Excel's built-in table features:

```
Cell A1 comment:
  jx:area(lastCell="E2")
  jx:each(items="orders" var="o" lastCell="E2")
  jx:table(name="OrdersTable" lastCell="E2" style="TableStyleMedium2")

A1: Order ID    B1: Customer    C1: Amount    D1: Status    E1: Date
A2: ${o.OrderID}  B2: ${o.Customer}  C2: ${o.Amount}  D2: ${o.Status}  E2: ${o.CreatedAt}
```

The output has:
- Auto-filter dropdowns on every column header
- Alternating row banding (striped rows)
- A named table (`OrdersTable`) that can be referenced in formulas
- Automatic total row option in Excel

### Table style options

Excel has dozens of built-in table styles:

| Style | Look |
|-------|------|
| `TableStyleLight1` - `TableStyleLight21` | Subtle, light backgrounds |
| `TableStyleMedium1` - `TableStyleMedium28` | Moderate contrast |
| `TableStyleDark1` - `TableStyleDark11` | Bold, dark headers |

Pick the one that matches your brand. `TableStyleMedium2` (blue header with banding) is a safe default.

## Streaming for large datasets

For exports with 10,000+ rows, enable streaming mode:

```go
data := xlfill.SQLRowsToData("orders", rows) // 100K orders

xlfill.Fill("export_template.xlsx", "large_export.xlsx", data,
    xlfill.WithStreaming(true),
)
```

Streaming mode writes rows to the output file as they're processed, instead of building the entire workbook in memory. The difference is dramatic:

| Rows | Sequential mode | Streaming mode |
|------|----------------|---------------|
| 10K | 1.2s, 85 MB | 0.4s, 30 MB |
| 100K | 12s, 800 MB | 3.5s, 120 MB |
| 500K | 60s, 4 GB | 15s, 250 MB |

> **Pro tip:** For datasets where you're not sure of the size upfront, use `WithAutoMode`:
> ```go
> xlfill.Fill("template.xlsx", "output.xlsx", data,
>     xlfill.WithAutoMode(map[string]any{"itemCount": len(orders)}),
> )
> ```
> XLFill analyzes the template and data hint, then picks sequential, streaming, or parallel mode automatically.

## HTTPHandler for download endpoints

Serve the export directly from an HTTP endpoint:

```go
http.Handle("/api/orders/export", xlfill.HTTPHandler("export_template.xlsx",
    func(r *http.Request) (map[string]any, error) {
        // Parse filters from query params
        status := r.URL.Query().Get("status")
        from := r.URL.Query().Get("from")

        rows, err := db.Query(
            "SELECT * FROM orders WHERE status = $1 AND created_at >= $2",
            status, from,
        )
        if err != nil {
            return nil, err
        }

        return xlfill.SQLRowsToData("orders", rows)
    },
    xlfill.WithStreaming(true),
))
```

Users click "Export to Excel" in your UI, the handler queries the database, fills the template, and streams the `.xlsx` directly to the browser. No temp files.

## Auto-fit column widths

Narrow columns with truncated text are the bane of data exports. Add `jx:autoColWidth`:

```
Cell A1 comment:
  jx:area(lastCell="E2")
  jx:each(items="orders" var="o" lastCell="E2")
  jx:autoColWidth(lastCell="E1")
```

Every column resizes to fit the widest value. Customer names, order IDs, dates — all fully visible.

## Formatting in templates

### Date formatting

Set the date format in the template cell (Format Cells → Date), or use `formatDate` in the expression:

```
${formatDate(o.CreatedAt, "2006-01-02")}
```

### Number formatting

Same approach — Excel number format or `formatNumber`:

```
${formatNumber(o.Amount, "#,##0.00")}
```

### Conditional formatting on status

Color-code the Status column:

```
Cell D2 comment:
  jx:conditionalFormat(lastCell="D2" type="cellIs" operator="equal"
                        formula="\"Cancelled\"" format="font:red")
```

Cancelled orders show in red. Active orders stay default. Users can filter the table to see only cancelled orders.

## Document properties for metadata

Add metadata to the generated file — useful for audit trails:

```go
xlfill.Fill("template.xlsx", "export.xlsx", data,
    xlfill.WithDocumentProperties(xlfill.DocProperties{
        Title:   "Orders Export",
        Author:  "Order Management System",
        Subject: fmt.Sprintf("Orders from %s to %s", from, to),
    }),
)
```

The metadata shows up in Excel's File → Properties and in Windows file properties.

## Complete data export example

```go
package main

import (
    "database/sql"
    "fmt"
    "log"
    "net/http"

    "github.com/javajack/xlfill"
    _ "github.com/lib/pq"
)

var db *sql.DB
var exportTemplate *xlfill.CompiledTemplate

func init() {
    var err error
    db, err = sql.Open("postgres", "postgres://localhost/myapp?sslmode=disable")
    if err != nil {
        log.Fatal(err)
    }

    exportTemplate, err = xlfill.Compile("templates/orders_export.xlsx",
        xlfill.WithStreaming(true),
        xlfill.WithDocumentProperties(xlfill.DocProperties{
            Author: "MyApp Export System",
        }),
    )
    if err != nil {
        log.Fatal(err)
    }
}

func exportHandler(w http.ResponseWriter, r *http.Request) {
    status := r.URL.Query().Get("status")
    if status == "" {
        status = "all"
    }

    var rows *sql.Rows
    var err error
    if status == "all" {
        rows, err = db.Query("SELECT order_id, customer, amount, status, created_at FROM orders ORDER BY created_at DESC")
    } else {
        rows, err = db.Query("SELECT order_id, customer, amount, status, created_at FROM orders WHERE status = $1 ORDER BY created_at DESC", status)
    }
    if err != nil {
        http.Error(w, "Query failed", http.StatusInternalServerError)
        return
    }

    data, err := xlfill.SQLRowsToData("orders", rows)
    if err != nil {
        http.Error(w, "Data conversion failed", http.StatusInternalServerError)
        return
    }

    w.Header().Set("Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition",
        fmt.Sprintf(`attachment; filename="orders_%s.xlsx"`, status))

    exportTemplate.FillWriter(data, w)
}

func main() {
    http.HandleFunc("/api/orders/export", exportHandler)
    log.Println("Listening on :8080")
    http.ListenAndServe(":8080", nil)
}
```

## When to use XLFill vs. plain CSV

| Criteria | CSV | XLFill Excel export |
|----------|-----|---------------------|
| Formatting | None | Full (colors, fonts, borders) |
| Auto-filter | No | Yes (jx:table) |
| Charts | No | Yes (jx:chart) |
| File size (100K rows) | ~15 MB | ~25 MB |
| Generation speed | Fastest | Fast (with streaming) |
| Opens in Excel | Yes (but ugly) | Yes (formatted) |
| Opens everywhere | Yes | Mostly (Excel, Sheets, LibreOffice) |

Use CSV for machine-to-machine data transfer. Use XLFill for anything a human will open.

## Tips and tricks

- **Use `WithAutoMode` for variable-size exports.** Let XLFill pick the optimal mode based on data size.

- **Add `jx:table` to every data export.** The auto-filter is the single most useful feature for users exploring data.

- **Use `formatDate` and `formatNumber` in templates.** Dates formatted as ISO strings (`2025-01-15`) are harder to read than `Jan 15, 2025`.

- **Set `Content-Disposition` with meaningful filenames.** `orders_active_2025-04-04.xlsx` is much better than `download.xlsx`.

- **Compile templates at startup for HTTP handlers.** Parse once, fill on every request.

- **Use `jx:autoColWidth` for readable exports.** It adds a fraction of a second but makes the output dramatically more usable.

## What's next?

- Build formatted reports with charts and grouping: **[Financial Reports &rarr;](/xlfill/use-cases/financial-reports/)**
- Serve exports from HTTP endpoints: **[Excel Export APIs &rarr;](/xlfill/guides/excel-export-api/)**
- Compare Go Excel libraries: **[Go Excel Libraries Compared &rarr;](/xlfill/guides/excel-libraries-compared/)**
