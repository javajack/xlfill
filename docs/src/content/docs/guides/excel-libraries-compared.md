---
title: "Go Excel Libraries Compared: Code-First vs Template-First"
description: "Excelize, tealeg/xlsx, go-xlsx-templater, and XLFill — which Go library should you use for Excel generation? A practical comparison."
---

You've got data in Go. You need it in Excel. But which library do you reach for? The Go ecosystem has several options, and they take fundamentally different approaches. Let's build the *same report* with each one and see what falls out.

## The report we're building

A simple employee list: Name, Department, Salary. Three columns, a styled header, data rows. Nothing exotic — the kind of thing every project needs eventually.

## Approach 1: excelize (code-first)

[excelize](https://github.com/qax-os/excelize) is the most popular Go library for reading and writing Excel files. It gives you full control over every cell, style, and formula. Here's the report:

```go
f := excelize.NewFile()
sheet := "Report"
f.NewSheet(sheet)

// Set column widths
f.SetColWidth(sheet, "A", "A", 20)
f.SetColWidth(sheet, "B", "B", 18)
f.SetColWidth(sheet, "C", "C", 15)

// Header style — 8 lines for some bold white text on a blue background
headerStyle, _ := f.NewStyle(&excelize.Style{
    Font:      &excelize.Font{Bold: true, Size: 12, Color: "FFFFFF"},
    Fill:      excelize.Fill{Type: "pattern", Color: []string{"4472C4"}, Pattern: 1},
    Alignment: &excelize.Alignment{Horizontal: "center"},
    Border: []excelize.Border{
        {Type: "left", Color: "000000", Style: 1},
        {Type: "top", Color: "000000", Style: 1},
        {Type: "right", Color: "000000", Style: 1},
        {Type: "bottom", Color: "000000", Style: 1},
    },
})

// Data row style — another 7 lines
dataStyle, _ := f.NewStyle(&excelize.Style{
    Border: []excelize.Border{
        {Type: "left", Color: "000000", Style: 1},
        {Type: "top", Color: "000000", Style: 1},
        {Type: "right", Color: "000000", Style: 1},
        {Type: "bottom", Color: "000000", Style: 1},
    },
})

// Write headers
f.SetCellValue(sheet, "A1", "Name")
f.SetCellValue(sheet, "B1", "Department")
f.SetCellValue(sheet, "C1", "Salary")
f.SetCellStyle(sheet, "A1", "C1", headerStyle)

// Write rows
for i, emp := range employees {
    row := i + 2
    f.SetCellValue(sheet, fmt.Sprintf("A%d", row), emp.Name)
    f.SetCellValue(sheet, fmt.Sprintf("B%d", row), emp.Department)
    f.SetCellValue(sheet, fmt.Sprintf("C%d", row), emp.Salary)
    f.SetCellStyle(sheet, fmt.Sprintf("A%d", row), fmt.Sprintf("C%d", row), dataStyle)
}

f.SaveAs("output.xlsx")
```

**~45 lines.** And the output looks... okay. No alternating row colors, no conditional formatting, no merged header. If the finance team asks you to change the header from blue to green, you're editing Go code, rebuilding, and redeploying.

## Approach 2: tealeg/xlsx (code-first, older)

[tealeg/xlsx](https://github.com/tealeg/xlsx) was one of the earliest Go Excel libraries. It's simpler than excelize but also less capable:

```go
file := xlsx.NewFile()
sheet, _ := file.AddSheet("Report")

// Header row
headerRow := sheet.AddRow()
for _, h := range []string{"Name", "Department", "Salary"} {
    cell := headerRow.AddCell()
    cell.SetValue(h)
    cell.GetStyle().Font.Bold = true
    cell.GetStyle().Fill.PatternType = "solid"
    cell.GetStyle().Fill.FgColor = "4472C4"
    cell.GetStyle().Font.Color = "FFFFFF"
}

// Data rows
for _, emp := range employees {
    row := sheet.AddRow()
    row.AddCell().SetValue(emp.Name)
    row.AddCell().SetValue(emp.Department)
    row.AddCell().SetFloat(emp.Salary)
}

file.Save("output.xlsx")
```

**~20 lines** — shorter, but no border support, limited styling, and the library is largely unmaintained. You'll hit walls fast with charts, conditional formatting, or anything beyond basic cell values.

## Approach 3: go-xlsx-templater (mustache templates)

[go-xlsx-templater](https://github.com/ivahaev/go-xlsx-templater) uses `{{mustache}}` syntax in Excel cells:

```go
doc := xlst.New()
doc.ReadTemplate("template.xlsx") // cells contain {{Name}}, {{Department}}, etc.
doc.Render(data)
doc.Save("output.xlsx")
```

**~4 lines of Go** — much better. But the template language is limited to simple value substitution. No loops (you can't repeat rows), no conditionals, no charts, no merged cells in loops. For anything beyond a flat key-value fill, you're stuck.

## Approach 4: XLFill (template-first, full-featured)

Design the template in Excel with all the formatting you want. Add `${e.Name}` in cells and a `jx:each(...)` command in a cell comment. Then:

```go
data := map[string]any{
    "employees": employees,
}

xlfill.Fill("template.xlsx", "output.xlsx", data)
```

**3 lines of Go.** The template carries all the styling — headers, colors, borders, number formats, merged cells, conditional formatting. Your code only provides data. When the finance team wants changes, they edit the `.xlsx` template themselves. You don't even hear about it.

## Side-by-side line count

| Library | Lines of Go | Template needed? | Styling in code? |
|---------|-------------|-------------------|-------------------|
| **excelize** | ~45 | No | Yes — every style is a struct literal |
| **tealeg/xlsx** | ~20 | No | Yes — limited style API |
| **go-xlsx-templater** | ~4 | Yes (mustache) | No — but no loops or commands |
| **XLFill** | ~3 | Yes (jx: commands) | No — template is the design |

## Decision matrix

| Criteria | excelize | tealeg/xlsx | go-xlsx-templater | XLFill |
|----------|----------|-------------|-------------------|--------|
| **Formatted reports** | Tedious | Limited | Basic | Excellent |
| **Raw data dumps** | Good | Acceptable | Overkill | Overkill |
| **Charts & sparklines** | Code-only | No | No | Template command |
| **Conditional formatting** | Code-only | No | No | Template command |
| **Data validation** | Code-only | No | No | Template command |
| **Non-devs can edit layout** | No | No | Partially | Yes |
| **Streaming large files** | Yes | No | No | Yes |
| **Active maintenance** | Yes | Stale | Stale | Yes |
| **JXLS migration** | N/A | N/A | N/A | Drop-in |
| **Batch generation** | Manual | Manual | Manual | Built-in |

## XLFill is a layer on top of excelize, not a replacement

This is an important point: XLFill uses excelize under the hood for reading and writing `.xlsx` files. It doesn't replace excelize — it adds a template engine on top.

If you need to do something excelize supports but XLFill doesn't expose as a template command, you can always post-process the output with excelize directly. XLFill fills the template; excelize handles the low-level I/O.

```go
// XLFill for the report, excelize for a one-off tweak
xlfill.Fill("template.xlsx", "output.xlsx", data)

f, _ := excelize.OpenFile("output.xlsx")
f.SetCellValue("Sheet1", "Z1", "Custom value")
f.Save()
```

> **Pro tip:** For 95% of use cases, you won't need this. XLFill's 20 commands and 18 built-in functions cover charts, tables, conditional formatting, data validation, grouping, images, and more — all from the template.

## When code-first is the better choice

Template-first isn't always the answer. Here's when raw excelize wins:

- **One-off data dumps** — if you're exporting a database table with no formatting requirements, excelize's `StreamWriter` is the simplest path. No template to maintain.
- **Programmatic spreadsheet editors** — if you're building a tool that *creates* spreadsheets dynamically (not filling reports), you need cell-level control.
- **Reading Excel files** — XLFill is a write/fill library. For reading `.xlsx`, use excelize directly.

> **Gotcha:** If you start with excelize for "just a quick data dump" and later get asked to add headers, branding, and conditional formatting, you'll wish you'd started with a template. It's easier to start with XLFill and keep the template simple than to retrofit styling code later.

## When template-first wins

- **Recurring reports** — monthly sales, quarterly P&L, weekly dashboards. Design once, fill forever.
- **Business-owned layouts** — when finance, operations, or marketing owns the design and needs to change it without developer involvement.
- **Complex formatting** — merged cells, conditional formatting, charts, data validation. These are painful to code and trivial to set up in Excel.
- **Multi-format output** — the same template with different data produces invoices, statements, certificates. Use `FillBatch` to generate thousands.
- **Team velocity** — 3 lines of Go vs. 45. Multiply that by every report in your system.

## Tips and tricks

- **Start with XLFill even for simple reports.** The overhead of a template file is tiny, and you'll thank yourself when requirements grow.
- **Use `Compile` for server-side reports.** Parse the template once, fill it thousands of times with zero file I/O per request.
- **Combine both libraries** when you need a one-off cell value that isn't template-driven. XLFill fills, excelize tweaks.
- **Profile with `SuggestMode`** if performance matters. XLFill will tell you whether streaming or parallel processing is optimal for your template.

## What's next?

- Already using JXLS in Java? See how to migrate: **[Migrating from JXLS &rarr;](/xlfill/guides/migrating-from-jxls/)**
- Coming from Python's openpyxl or XlsxWriter? **[XLFill for Python Developers &rarr;](/xlfill/guides/coming-from-python/)**
- Ready to try it? **[Getting Started in 5 minutes &rarr;](/xlfill/guides/getting-started/)**
