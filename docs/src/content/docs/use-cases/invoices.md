---
title: "Invoice Generation with XLFill"
description: "Generate professional invoices from Excel templates — line items, tax calculations, conditional formatting, and company branding."
---

Invoices are one of the most common Excel generation tasks — and one of the hardest to get right with code-first libraries. Company logo, styled headers, line items with varying row counts, subtotals, tax calculations, conditional formatting for overdue amounts. With XLFill, your accountant designs the invoice in Excel, and your Go code just provides the data.

## The complete invoice template

Here's a production-quality invoice template. Everything below is configured in the `.xlsx` file — zero styling code in Go.

### Template structure

```
Row 1-5: Company header (logo, name, address)
Row 6: Invoice metadata (number, date, due date)
Row 7: Customer details
Row 8: Column headers (Item, Description, Qty, Unit Price, Amount)
Row 9: ${item.Name} | ${item.Description} | ${item.Qty} | ${item.UnitPrice} | =C9*D9
Row 10: Subtotal row (formula)
Row 11: Tax row (parameterized formula)
Row 12: Total row (formula)
```

### Cell comments (commands)

```
Cell A1 comment:
  jx:area(lastCell="E12")

Cell A1 comment (for shared header):
  jx:include(src="CommonHeader" lastCell="E5")

Cell A3 comment (for logo):
  jx:image(src="logoBytes" lastCell="B5" imageType="png")

Cell A9 comment:
  jx:each(items="items" var="item" lastCell="E9")

Cell E10 comment (subtotal):
  Formula: =SUM(E9:E9)
  (XLFill expands this to =SUM(E9:E{lastRow}) automatically)

Cell E11 comment (tax):
  Formula: =E10*${taxRate}

Cell E12:
  Formula: =E10+E11
```

### Go code

```go
data := map[string]any{
    "invoiceNo":   "INV-2025-0042",
    "invoiceDate": time.Now().Format("January 2, 2006"),
    "dueDate":     time.Now().AddDate(0, 0, 30).Format("January 2, 2006"),
    "customer": map[string]any{
        "Name":    "Acme Corporation",
        "Address": "123 Business Ave, Suite 400",
        "City":    "San Francisco, CA 94102",
        "Email":   "billing@acme.com",
    },
    "items": []map[string]any{
        {"Name": "Consulting", "Description": "Senior developer, 40 hours", "Qty": 40, "UnitPrice": 150.00},
        {"Name": "Code Review", "Description": "Architecture review", "Qty": 8, "UnitPrice": 200.00},
        {"Name": "Training", "Description": "Go workshop, full day", "Qty": 1, "UnitPrice": 2500.00},
    },
    "taxRate":  0.08,
    "logoBytes": loadCompanyLogo(),
}

xlfill.Fill("invoice_template.xlsx", "invoice_INV-2025-0042.xlsx", data)
```

That's it. The invoice has a company logo, customer details, line items, subtotal, 8% tax, and a grand total. All formatting comes from the template.

## Company logo with jx:image

```
Cell A1 comment:
  jx:image(src="logoBytes" lastCell="B3" imageType="png")
```

The `logoBytes` field in your data map should be a `[]byte` containing the image data:

```go
func loadCompanyLogo() []byte {
    data, err := os.ReadFile("assets/logo.png")
    if err != nil {
        return nil // no logo — cell stays empty
    }
    return data
}
```

> **Pro tip:** Load the logo once at startup and reuse it across all invoices. For batch generation, include it in the compiled template's data:
> ```go
> logo := loadCompanyLogo()
> for _, invoice := range invoices {
>     invoice["logoBytes"] = logo
> }
> ```

## Shared header with jx:include

If multiple templates share the same company header (invoices, statements, receipts), use `jx:include`:

1. Create a sheet called "CommonHeader" in your template with the logo, company name, and address
2. In your invoice sheet, reference it:

```
Cell A1 comment:
  jx:include(src="CommonHeader" lastCell="E5")
```

The header is pulled from the CommonHeader sheet. Update it once, and every template that includes it gets the change.

## Line items with jx:each

```
Cell A9 comment:
  jx:each(items="items" var="item" lastCell="E9")

A9: ${item.Name}
B9: ${item.Description}
C9: ${item.Qty}
D9: ${formatNumber(item.UnitPrice, "#,##0.00")}
E9: =C9*D9
```

The formula `=C9*D9` is in the template. XLFill copies it for every row and adjusts the row references automatically. Three items? You get `=C9*D9`, `=C10*D10`, `=C11*D11`.

## Tax with parameterized formulas

```
Cell E11: =E10*${taxRate}
```

XLFill evaluates `${taxRate}` to `0.08`, then the cell formula becomes `=E10*0.08`. The formula is a real Excel formula — recalculates if users change values.

> **Did you know?** You can use any expression in formulas. `=${discountRate}*E10` for discounts, `=E10*(1+${taxRate})` for tax-inclusive totals — any mix of expressions and Excel formula syntax works.

## Conditional formatting for overdue invoices

Highlight the due date cell in red if the invoice is overdue:

```
Cell C6 comment:
  jx:conditionalFormat(lastCell="C6" type="cellIs" operator="lessThan"
                        formula="TODAY()" format="font:red,bold;fill:FFC7CE")
```

If the due date is in the past, it shows up in bold red with a pink background. Your users immediately see which invoices need attention.

## Sheet protection

Lock the invoice so customers can't accidentally modify it:

```
Cell A1 comment:
  jx:protect(password="inv2025" lastCell="E12")
```

All cells are locked. The invoice is read-only. Professional and tamper-resistant.

## Blank line padding with jx:repeat

If your invoice template expects a minimum number of line items (for consistent layout or printing), use `jx:repeat` to pad with blank rows:

```
Cell A13 comment:
  jx:repeat(count="${max(0, 10 - len(items))}" lastCell="E13")
```

If there are 3 items, 7 blank rows are added. If there are 10 or more items, no blank rows. The invoice always has at least 10 data rows for consistent page layout.

## Currency formatting with formatNumber

```
${formatNumber(item.UnitPrice, "#,##0.00")}
```

Or set the Excel number format on the cell in the template (Format Cells → Number → Currency). XLFill preserves it. Either approach works.

> **Pro tip:** For multi-currency invoices, use the number format in the template for the display format, and pass the currency symbol in the data:
> ```
> ${item.CurrencySymbol}${formatNumber(item.UnitPrice, "#,##0.00")}
> ```

## Page breaks for batch invoices

When generating multiple invoices into a single workbook (e.g., for printing), use `jx:pageBreak`:

```
Cell A1 comment:
  jx:area(lastCell="E12")
  jx:each(items="invoices" var="inv" lastCell="E12")
  jx:pageBreak(lastCell="E12")
```

Each invoice starts on a new page when printed.

## Complete batch generation

Generate hundreds of invoices from one template:

```go
compiled, _ := xlfill.Compile("invoice_template.xlsx",
    xlfill.WithRecalculateOnOpen(true),
)

logo := loadCompanyLogo()

datasets := make([]map[string]any, len(customers))
for i, cust := range customers {
    datasets[i] = map[string]any{
        "invoiceNo":   cust.NextInvoiceNumber(),
        "invoiceDate": time.Now().Format("January 2, 2006"),
        "dueDate":     time.Now().AddDate(0, 0, 30).Format("January 2, 2006"),
        "customer":    cust,
        "items":       cust.PendingLineItems(),
        "taxRate":     cust.TaxRate,
        "logoBytes":   logo,
    }
}

files, err := compiled.FillBatch(datasets, "./invoices/",
    func(i int, data map[string]any) string {
        return fmt.Sprintf("%s.xlsx", data["invoiceNo"])
    },
)

fmt.Printf("Generated %d invoices\n", len(files))
```

## Tips and tricks

- **Use `jx:include` for shared headers.** Company logo and address in one place, used by invoices, statements, receipts.

- **Put formulas in the template, not in Go.** `=SUM(E9:E9)` in the template auto-expands. Building formula strings in Go is fragile.

- **Use `WithRecalculateOnOpen(true)` for formula-heavy invoices.** This tells Excel to recalculate all formulas when the file is opened, ensuring totals are correct even if the formula engine has rounding differences.

- **Add `jx:freezePanes` for long invoices.** Keep the header visible while scrolling:
  ```
  jx:freezePanes(lastCell="E8")
  ```

- **Test with zero items, one item, and many items.** Edge cases in invoices are embarrassing — an invoice with no line items should still look reasonable.

## What's next?

- Build financial statements with grouping and charts: **[Financial Reports &rarr;](/xlfill/use-cases/financial-reports/)**
- Generate thousands of invoices at once: **[Batch Generation &rarr;](/xlfill/guides/batch-generation/)**
- Add data validation for editable invoice fields: **[Data Validation and Forms &rarr;](/xlfill/guides/data-validation-forms/)**
