---
title: "Financial Report Generation in Go"
description: "Build financial statements, P&L reports, and budget summaries from Excel templates — with charts, grouping, and conditional formatting."
---

Financial reports are the ultimate test of an Excel generation library. They need everything: categorized sections, subtotals, grand totals, trend charts, variance highlighting, collapsible groups, frozen headers, and print-ready layouts. With XLFill, all of this lives in the template. Your Go code just feeds the numbers.

## P&L statement with categorized rows

A profit and loss statement has revenue categories, expense categories, subtotals, and a net income line. Here's how to build it with XLFill.

### Template layout

```
Row 1: Company Name — P&L Statement — ${period}
Row 2: (blank)
Row 3: REVENUE
Row 4: ${rev.Category} | ${rev.Actual} | ${rev.Budget} | =B4-C4
Row 5: Total Revenue | ${sumBy(revenueItems, "Actual")} | ${sumBy(revenueItems, "Budget")}
Row 6: (blank)
Row 7: EXPENSES
Row 8: ${exp.Category} | ${exp.Actual} | ${exp.Budget} | =B8-C8
Row 9: Total Expenses | ${sumBy(expenseItems, "Actual")} | ${sumBy(expenseItems, "Budget")}
Row 10: (blank)
Row 11: NET INCOME | =B5-B9 | =C5-C9 | =D5-D9
```

### Cell comments

```
Cell A1 comment:
  jx:area(lastCell="D11")
  jx:freezePanes(lastCell="D2")

Cell A4 comment:
  jx:each(items="revenueItems" var="rev" lastCell="D4")

Cell A8 comment:
  jx:each(items="expenseItems" var="exp" lastCell="D8")

Cell D4 comment:
  jx:conditionalFormat(lastCell="D4" type="cellIs" operator="lessThan"
                        formula="0" format="font:red")

Cell D8 comment:
  jx:conditionalFormat(lastCell="D8" type="cellIs" operator="greaterThan"
                        formula="0" format="font:red")
```

### Go code

```go
data := map[string]any{
    "period": "Q4 2025",
    "revenueItems": []map[string]any{
        {"Category": "Product Sales", "Actual": 1250000, "Budget": 1100000},
        {"Category": "Service Revenue", "Actual": 480000, "Budget": 500000},
        {"Category": "Licensing Fees", "Actual": 320000, "Budget": 300000},
    },
    "expenseItems": []map[string]any{
        {"Category": "Cost of Goods Sold", "Actual": 625000, "Budget": 550000},
        {"Category": "Salaries & Benefits", "Actual": 450000, "Budget": 460000},
        {"Category": "Marketing", "Actual": 180000, "Budget": 200000},
        {"Category": "R&D", "Actual": 220000, "Budget": 210000},
        {"Category": "General & Admin", "Actual": 95000, "Budget": 100000},
    },
}

xlfill.Fill("pnl_template.xlsx", "pnl_q4_2025.xlsx", data)
```

Revenue items expand. Expense items expand. The variance column (`=B4-C4`) auto-adjusts row references. Negative variances in revenue show red; over-budget expenses show red. Frozen headers stay visible while scrolling.

## Department grouping with groupBy and jx:group

For reports with many departments, use `groupBy` to automatically group and `jx:group` for collapsible sections:

```
Cell A2 comment:
  jx:each(items="expenses" var="exp" groupBy="Department"
          orderBy="Department asc" lastCell="D3")
  jx:group(lastCell="D3" collapsed="false")

Row 2: ${exp.Department} (section header — merged across columns)
Row 3: ${exp.Category} | ${exp.Amount} | ${exp.Budget} | =B3-C3
```

```go
data := map[string]any{
    "expenses": []map[string]any{
        {"Department": "Engineering", "Category": "Cloud Infra", "Amount": 45000, "Budget": 40000},
        {"Department": "Engineering", "Category": "Tools & Licenses", "Amount": 12000, "Budget": 15000},
        {"Department": "Marketing", "Category": "Ad Spend", "Amount": 80000, "Budget": 75000},
        {"Department": "Marketing", "Category": "Events", "Amount": 25000, "Budget": 30000},
        {"Department": "Sales", "Category": "Travel", "Amount": 18000, "Budget": 20000},
        {"Department": "Sales", "Category": "CRM License", "Amount": 8000, "Budget": 8000},
    },
}
```

Each department becomes a collapsible group. Users can expand Engineering, collapse Marketing, and focus on the data they care about.

> **Pro tip:** Use `groupBy` instead of nested `jx:each` when your data comes from a flat SQL query. No need to restructure it into a hierarchy — XLFill handles the grouping for you.

## Charts for trends

Add a column chart showing actual vs. budget alongside the numbers:

```
Cell F1 comment:
  jx:chart(lastCell="K15" type="col" title="Revenue: Actual vs Budget"
           catRange="A4:A4" valRange="B4:B4,C4:C4"
           seriesNames="Actual,Budget" legendPosition="bottom" style="26")
```

The chart categories come from the revenue category names, and the two series show actual and budget side by side. The data ranges auto-expand with the revenue items.

### Monthly trend line chart

If your data includes monthly figures:

```
Row 2: ${m.Month} | ${m.Revenue} | ${m.Expenses} | ${m.NetIncome}

Cell F1 comment:
  jx:chart(lastCell="K15" type="line" title="Monthly P&L Trend"
           catRange="A2:A2" valRange="B2:B2,C2:C2,D2:D2"
           seriesNames="Revenue,Expenses,Net Income" legendPosition="bottom")
```

## Sparklines for monthly data

Add inline trend indicators next to each category:

```
A1: Category  B1: Jan  C1: Feb  D1: Mar  E1: Apr  ...  M1: Trend
A2: ${cat.Name}  B2: ${cat.Jan}  C2: ${cat.Feb}  ...

Cell M2 comment:
  jx:sparkline(lastCell="M2" type="line" dataRange="B2:L2")
```

Each category row shows a tiny line chart of its 12-month trend. Finance teams love this — instant visual insight without scrolling to a chart.

## Conditional formatting for variances

### Color scale on variance percentages

```
Cell E2 comment:
  jx:conditionalFormat(lastCell="E2" type="colorScale"
                        minColor="FF0000" midColor="FFFF00" maxColor="00FF00")

Cell E2: ${formatNumber((B2-C2)/C2*100, 1) + "%"}
```

Red for negative variances, yellow for near-zero, green for positive. The entire variance column becomes a heat map.

### Data bars for budget utilization

```
Cell F2 comment:
  jx:conditionalFormat(lastCell="F2" type="dataBar"
                        minColor="E6F2FF" maxColor="2F5496")

Cell F2: =B2/C2
```

Each cell shows a proportional bar — visually compare budget utilization across categories at a glance.

### Icon sets for status

```
Cell G2 comment:
  jx:conditionalFormat(lastCell="G2" type="iconSet" iconStyle="3TrafficLights")
```

Green light for under budget, yellow for within 5%, red for over budget.

## Aggregation with sumBy

XLFill's built-in aggregation functions work directly in expressions — no Excel formulas needed:

```
Total Revenue:    ${sumBy(revenueItems, "Actual")}
Average Expense:  ${avgBy(expenseItems, "Amount")}
Category Count:   ${countBy(expenseItems, "Category")}
Highest Expense:  ${maxBy(expenseItems, "Amount")}
Lowest Expense:   ${minBy(expenseItems, "Amount")}
```

These evaluate during template filling. The output cells contain the computed values, not formulas. This is useful for summary rows that should show static values.

> **Did you know?** You can mix `sumBy` expressions with Excel formulas in the same report. Use `sumBy` for summary sections and Excel `=SUM()` for ranges that users might want to verify.

## Named ranges for downstream formulas

If your report feeds into pivot tables or other workbooks, use `jx:definedName` to create named ranges:

```
Cell A4 comment:
  jx:each(items="revenueItems" var="rev" lastCell="D4")
  jx:definedName(name="RevenueData" lastCell="D4")
```

The output workbook has a named range "RevenueData" covering all revenue rows. Other sheets or workbooks can reference it by name.

## Layout polish

### Freeze header rows

```
Cell A1 comment:
  jx:freezePanes(lastCell="D2")
```

The first two rows (title and column headers) stay visible while scrolling.

### Auto-fit column widths

```
Cell A1 comment:
  jx:autoColWidth(lastCell="D1")
```

Columns resize to fit the widest value — no manual width adjustment.

## Complete financial report template

Putting it all together:

```
Cell A1 comment:
  jx:area(lastCell="M20")
  jx:freezePanes(lastCell="M2")
  jx:autoColWidth(lastCell="M1")

Row 1: Company — Financial Report — ${period}
Row 2: Category | Jan | Feb | Mar | ... | Dec | Total | Trend

Cell A3 comment:
  jx:each(items="categories" var="cat" groupBy="Section"
          orderBy="Section asc, Category asc" lastCell="M3")
  jx:group(lastCell="M3")

Row 3: ${cat.Category} | ${cat.Jan} | ... | ${cat.Dec} | =SUM(B3:M3)

Cell N3 comment:
  jx:sparkline(lastCell="N3" type="line" dataRange="B3:M3")

Cell B3 comment:
  jx:conditionalFormat(lastCell="M3" type="colorScale"
                        minColor="FFC7CE" maxColor="C6EFCE")

Cell P1 comment:
  jx:chart(lastCell="V12" type="line" title="Monthly Trends"
           catRange="A3:A3" valRange="B3:B3" legendPosition="right")
```

```go
data := map[string]any{
    "period":     "FY 2025",
    "categories": buildCategoryData(), // flat list with Section, Category, Jan-Dec fields
}

xlfill.Fill("financial_template.xlsx", "financial_report_2025.xlsx", data)
```

## Tips and tricks

- **Use `jx:freezePanes` for every financial report.** Headers should always be visible.

- **Use `jx:autoColWidth` for readability.** Number columns need enough width for formatted values like `$1,250,000.00`.

- **Use `jx:definedName` if the report feeds into pivot tables.** Named ranges are more robust than cell references for cross-sheet formulas.

- **Combine `groupBy` with aggregation.** Use `sumBy` in the group header row and `jx:group` for collapsible sections — users get subtotals and drill-down in one report.

- **Test with real financial data.** Numbers like `$0.00`, negative values, and very large amounts can break formatting assumptions. Run edge case tests.

- **Use `WithRecalculateOnOpen(true)` for formula-heavy reports.** Ensures Excel recalculates everything when the file is opened.

## What's next?

- Export raw data with auto-filter tables: **[Data Exports &rarr;](/xlfill/use-cases/data-exports/)**
- Generate hundreds of department reports at once: **[Batch Generation &rarr;](/xlfill/guides/batch-generation/)**
- See all chart and conditional formatting options: **[Charts and Dashboards &rarr;](/xlfill/guides/charts-and-dashboards/)**
