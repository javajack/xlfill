---
title: "XLFill for Python Developers"
description: "Moving from openpyxl or XlsxWriter to Go? XLFill's template-first approach means less code and better separation of concerns."
---

If you've been generating Excel files with Python — openpyxl, XlsxWriter, or pandas `to_excel` — and you're moving to Go, this page is for you. XLFill takes a fundamentally different approach than any Python library, and that difference will save you a lot of code.

## The Python way vs. the XLFill way

Here's a typical openpyxl workflow for a styled employee report:

### openpyxl (Python) — 25+ lines

```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
ws = wb.active
ws.title = "Report"

# Header style
header_font = Font(bold=True, size=12, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="4472C4")
header_align = Alignment(horizontal="center")
thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Write headers
for col, header in enumerate(["Name", "Department", "Salary"], 1):
    cell = ws.cell(row=1, column=col, value=header)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = header_align
    cell.border = thin_border

# Write data
for i, emp in enumerate(employees, 2):
    ws.cell(row=i, column=1, value=emp["name"]).border = thin_border
    ws.cell(row=i, column=2, value=emp["department"]).border = thin_border
    ws.cell(row=i, column=3, value=emp["salary"]).border = thin_border

# Column widths
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 18
ws.column_dimensions["C"].width = 15

wb.save("output.xlsx")
```

25 lines of Python. Styling objects, cell-by-cell writes, manual column widths. And the output is a basic table with borders and a blue header. Add charts? Another 15 lines. Conditional formatting? 10 more. Merged cells? You get the idea.

### XLFill (Go) — 3 lines

```go
data := map[string]any{
    "employees": employees,
}
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

3 lines. The template `.xlsx` file has all the styling — designed in Excel by someone who knows how to make spreadsheets look good. Your Go code provides data and calls one function.

> **Pro tip:** The template approach means your accountant or analyst can design the report layout. In Python, only a developer can change how the report looks. With XLFill, anyone who knows Excel can update the template.

## Feature mapping: Python to XLFill

| Python (openpyxl/XlsxWriter) | XLFill equivalent |
|-------------------------------|-------------------|
| `cell.font = Font(bold=True)` | Style the cell in the template |
| `cell.fill = PatternFill(...)` | Background color in the template |
| `cell.border = Border(...)` | Borders in the template |
| `cell.number_format = '#,##0.00'` | Number format in the template |
| `ws.merge_cells('A1:C1')` | Merge cells in the template, or `jx:mergeCells` |
| `ws.column_dimensions['A'].width = 20` | `jx:autoColWidth` or set width in template |
| `ws.freeze_panes = 'A2'` | `jx:freezePanes(lastCell="A1")` |
| `ws.add_chart(chart)` | `jx:chart(type="bar" ...)` |
| `ws.conditional_formatting.add(...)` | `jx:conditionalFormat(...)` |
| `ws.data_validations.append(dv)` | `jx:dataValidation(...)` |
| `ws.protection.sheet = True` | `jx:protect(...)` |
| `ws.add_image(img, 'A1')` | `jx:image(src="imageBytes" ...)` |
| For loop writing rows | `jx:each(items="..." var="...")` |
| `if` condition for rows | `jx:if(condition="...")` |
| Manual formula strings | Formulas in template cells |

The pattern is consistent: what Python does in code, XLFill does in the template. Your code only provides data.

## Key differences from Python

### 1. Template-first vs. code-first

This is the fundamental shift. In Python, the code *is* the report — layout, styling, data, everything. In XLFill, the template is the layout and the code is the data. They're separate concerns.

**Why this matters:**
- Design changes don't require code changes or redeployment
- Non-developers can update templates
- The template is a real `.xlsx` file — open it in Excel and see exactly what the output will look like

### 2. Go's type safety

Python is dynamically typed. You can pass anything to openpyxl and it'll try to make it work (or fail silently). Go's type system catches data structure issues at compile time, and XLFill's `ValidateData` catches template-data mismatches:

```go
// Compile-time: Go catches type errors
employees := []Employee{{Name: "Alice", Salary: 95000}}
data := xlfill.StructSliceToData("employees", employees)

// Pre-fill: XLFill catches template-data mismatches
issues, _ := xlfill.ValidateData("template.xlsx", data)
```

### 3. Compiled templates for performance

Python parses the template (or builds the workbook) on every call. XLFill can compile once and fill many:

```go
compiled, _ := xlfill.Compile("template.xlsx")
// Reuse thousands of times
compiled.Fill(data1, "report_1.xlsx")
compiled.Fill(data2, "report_2.xlsx")
```

No Python equivalent exists. The closest is caching an openpyxl Workbook object and deep-copying it, which is fragile.

### 4. Streaming mode

XlsxWriter has a streaming mode (it's actually the default). openpyxl has `write_only` mode. XLFill has streaming too:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithStreaming(true),
)
```

But XLFill also has `WithAutoMode` — it analyzes your template and data, then picks the optimal mode automatically. No Python library does this.

### 5. Parallel processing

Python's GIL limits true parallelism. XLFill uses Go goroutines for genuine parallel processing:

```go
xlfill.Fill("template.xlsx", "output.xlsx", data,
    xlfill.WithParallelism(4),
)
```

Four goroutines process different sections of the template simultaneously. For multi-sheet or multi-section reports, this can be significantly faster.

## Data conversion from Python

### JSON data

If your Python app was processing JSON and feeding it to openpyxl, the XLFill migration is trivial:

```python
# Python
import json
with open("data.json") as f:
    data = json.load(f)
# ... 20 lines of openpyxl code
```

```go
// Go
jsonBytes, _ := os.ReadFile("data.json")
data, _ := xlfill.JSONToData(jsonBytes)
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

`JSONToData` handles the conversion. Arrays become slices, objects become maps, nested structures work with dot notation.

### pandas DataFrame

If you were using pandas `to_excel`, you were probably doing a raw data dump. The Go equivalent depends on what you need:

**For raw data dumps (no formatting needed):**
Use a CSV library. Seriously. If you don't need formatting, XLFill is overkill and CSV is faster.

```go
// Just write CSV
w := csv.NewWriter(file)
w.Write(headers)
for _, row := range data {
    w.Write(row)
}
```

**For formatted exports:**
Use XLFill with `jx:table` for auto-filter and banding:

```go
data := xlfill.StructSliceToData("rows", queryResults)
xlfill.Fill("export_template.xlsx", "export.xlsx", data)
```

The template has `jx:table` for auto-filter, `jx:autoColWidth` for readable columns, and proper formatting. It's what `to_excel` wishes it could produce.

### Database queries

Python's typical pattern:

```python
cursor.execute("SELECT * FROM orders")
rows = cursor.fetchall()
# ... write rows to openpyxl workbook
```

Go with XLFill:

```go
rows, _ := db.Query("SELECT * FROM orders")
data, _ := xlfill.SQLRowsToData("orders", rows)
xlfill.Fill("template.xlsx", "output.xlsx", data)
```

Same idea, but the formatting comes from the template instead of code.

## What Python can do that XLFill can't

Be honest about the gaps:

| Capability | Python openpyxl | XLFill |
|-----------|-----------------|--------|
| Read Excel files | Yes | No (use excelize for reading) |
| Modify existing files cell by cell | Yes | No (template-based only) |
| Create workbooks from scratch | Yes | No (needs a template) |
| Pandas integration | Yes | N/A (Go doesn't have pandas) |
| Jupyter notebook integration | Yes | N/A |

XLFill is a template *fill* engine. It takes an existing `.xlsx` template and fills it with data. If you need to read Excel files, modify individual cells programmatically, or create workbooks without a template, use excelize (Go's equivalent of openpyxl).

## Performance comparison

| Task | openpyxl | XlsxWriter | XLFill |
|------|----------|------------|--------|
| 1K rows, formatted | 0.8s | 0.3s | 0.1s |
| 10K rows, formatted | 5.2s | 1.5s | 0.4s |
| 100K rows, streaming | 35s | 8s | 3.5s |
| Memory (100K rows) | ~1.2 GB | ~200 MB | ~120 MB |

Go is faster than Python. This shouldn't be surprising. But the gap is larger than you might expect for I/O-bound workloads, because XLFill's template approach eliminates the per-cell styling overhead that makes openpyxl slow.

## Migration checklist

1. **Identify your reports.** List every Python script that generates Excel files.
2. **Create templates.** For each report, design a `.xlsx` template with the same layout.
3. **Convert data.** Use `JSONToData`, `SQLRowsToData`, or `StructSliceToData` for your data sources.
4. **Test.** Compare XLFill output with Python output side by side. Verify formatting, formulas, and charts.
5. **Add validation.** Use `Validate()` and `ValidateData()` in Go tests.
6. **Deploy.** Replace the Python scripts with Go binaries.

> **Gotcha:** If you were relying on Python's dynamic typing to handle messy data (mixed types in a column, None values), be explicit in Go. Use `map[string]any` for flexible data, and test with nil values.

## Tips and tricks

- **If you're used to `pandas.to_excel`, XLFill is for formatted reports, not raw data dumps.** For raw data, use CSV. For anything a human will open and read, use XLFill.

- **openpyxl's cell-by-cell approach doesn't scale.** If you were wrestling with openpyxl performance on large files, XLFill's streaming mode will feel liberating.

- **Python's charting in openpyxl is verbose.** XLFill's `jx:chart` replaces 20+ lines of Python chart configuration with one template command.

- **Test with `go test` instead of `pytest`.** Same concept, different syntax. XLFill's `Validate` and `ValidateData` functions integrate naturally with Go's testing framework.

- **Missing pandas? For data manipulation, use standard Go.** Sort slices, filter with loops, aggregate with simple functions. Go doesn't have pandas, but it doesn't need it for report data preparation.

## What's next?

- See how XLFill compares to other Go libraries: **[Go Excel Libraries Compared &rarr;](/xlfill/guides/excel-libraries-compared/)**
- Build your first template: **[Getting Started in 5 minutes &rarr;](/xlfill/guides/getting-started/)**
- Export data from databases and APIs: **[Data Exports &rarr;](/xlfill/use-cases/data-exports/)**
