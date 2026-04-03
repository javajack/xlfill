---
title: "Excel Charts and Dashboards with XLFill"
description: "Add bar charts, line charts, pie charts, sparklines, and conditional formatting to your Excel reports — all from the template, zero chart code."
---

Charts, sparklines, and conditional formatting turn a wall of numbers into a story. With XLFill, you add all three from the template — no Go code for chart objects, no series definitions, no color hex codes. Design the layout in Excel, drop in a `jx:chart` command, and the data does the rest.

## jx:chart — 8 chart types, one command

The `jx:chart` command embeds a native Excel chart in your output. XLFill automatically sizes the data ranges based on actual output rows.

### Supported chart types

| Type | Keyword | Best for |
|------|---------|----------|
| Bar (horizontal) | `bar` | Comparing categories |
| Column (vertical) | `col` | Comparing categories (when labels are short) |
| Line | `line` | Time series, trends |
| Pie | `pie` | Composition, proportions |
| Doughnut | `doughnut` | Composition with a center label |
| Area | `area` | Cumulative totals over time |
| Scatter | `scatter` | Correlation between two variables |
| Radar | `radar` | Multi-axis comparison (skill maps, balanced scorecards) |

### Bar chart example

Template layout:
```
A1: Region          B1: Revenue
A2: ${r.Name}       B2: ${r.Revenue}

Cell A1 comment:
  jx:area(lastCell="H20")
  jx:each(items="regions" var="r" lastCell="B2")

Cell D1 comment:
  jx:chart(lastCell="H15" type="bar" title="Revenue by Region"
           catRange="A2:A2" valRange="B2:B2")
```

Go code:
```go
data := map[string]any{
    "regions": []map[string]any{
        {"Name": "North America", "Revenue": 1200000},
        {"Name": "Europe", "Revenue": 980000},
        {"Name": "Asia Pacific", "Revenue": 750000},
        {"Name": "Latin America", "Revenue": 420000},
    },
}
xlfill.Fill("template.xlsx", "dashboard.xlsx", data)
```

XLFill expands the data rows, then adjusts the chart's `catRange` and `valRange` to cover all output rows. If you have 4 regions, the chart references `A2:A5` and `B2:B5` automatically.

> **Pro tip:** The `catRange` and `valRange` use template-relative coordinates (just one row: `A2:A2`). XLFill expands them to match the actual output. You never hard-code row counts.

### Multi-series line chart

Show revenue and costs over time in one chart:

```
A1: Month    B1: Revenue    C1: Costs

Cell D1 comment:
  jx:chart(lastCell="H18" type="line" title="Revenue vs Costs"
           catRange="A2:A2" valRange="B2:B2,C2:C2"
           seriesNames="Revenue,Costs" legendPosition="bottom")
```

Multiple value ranges are comma-separated. Each gets its own series with the name from `seriesNames`.

### Pie chart

```
Cell D1 comment:
  jx:chart(lastCell="H12" type="pie" title="Market Share"
           catRange="A2:A2" valRange="B2:B2")
```

Pie charts use categories as slice labels and values as slice sizes. Simple and effective for composition data.

### Chart styling options

| Attribute | Values | Effect |
|-----------|--------|--------|
| `style` | 1-48 | Matches Excel's built-in chart styles |
| `legendPosition` | `bottom`, `left`, `right`, `top`, `none` | Where the legend appears |
| `width` | EMUs or cell count | Chart width |
| `height` | EMUs or cell count | Chart height |

```
jx:chart(lastCell="H15" type="col" title="Sales by Product"
         catRange="A2:A2" valRange="B2:B2"
         style="26" legendPosition="none")
```

> **Did you know?** Excel's chart style 26 gives you a clean, modern look with subtle gradients. Try styles in the 20-30 range for professional reports.

## jx:sparkline — inline trends

Sparklines are tiny charts that fit inside a single cell. They're perfect for showing trends without taking up chart real estate.

```
Cell D1 comment:
  jx:sparkline(lastCell="D1" type="line" dataRange="B2:G2")
```

### Sparkline types

| Type | Keyword | Best for |
|------|---------|----------|
| Line | `line` | Trends over time |
| Column | `column` | Comparing discrete values |
| Win/Loss | `winLoss` | Positive/negative outcomes |

### Sparkline in a loop

Add a trend sparkline to each row in a report:

```
A1: Product    B1: Jan   C1: Feb   D1: Mar   E1: Trend
A2: ${p.Name}  B2: ${p.Jan}  C2: ${p.Feb}  D2: ${p.Mar}

Cell E2 comment:
  jx:sparkline(lastCell="E2" type="line" dataRange="B2:D2")

Cell A1 comment:
  jx:area(lastCell="E2")
  jx:each(items="products" var="p" lastCell="E2")
```

Each product row gets its own sparkline showing the 3-month trend. The sparkline data range is relative to each output row.

### Sparklines vs. charts

| | Sparkline | Chart |
|---|-----------|-------|
| Size | One cell | Multi-cell area |
| Interactivity | None | Hover, click, resize |
| Best for | Row-level trends | Report-level summaries |
| Types | 3 (line, column, win/loss) | 8 types |
| Multiple series | No | Yes |

Use sparklines for per-row context, charts for overall picture. They complement each other beautifully.

## jx:conditionalFormat — data bars, color scales, icon sets

Conditional formatting makes patterns jump off the page without a single chart.

### Data bars

Turn cells into inline bar charts:

```
Cell B2 comment:
  jx:conditionalFormat(lastCell="B2" type="dataBar"
                       minColor="E6F2FF" maxColor="2F5496")
```

Every cell in column B gets a proportional bar fill. Higher values = longer bars. The min/max colors define the gradient.

### Color scales

Color-code a range from low to high:

```
Cell C2 comment:
  jx:conditionalFormat(lastCell="C2" type="colorScale"
                       minColor="FF0000" midColor="FFFF00" maxColor="00FF00")
```

Red → yellow → green. Classic heat map. Great for performance metrics, test scores, or financial variances.

### Icon sets

Add visual indicators (arrows, traffic lights, stars):

```
Cell D2 comment:
  jx:conditionalFormat(lastCell="D2" type="iconSet"
                       iconStyle="3Arrows")
```

Available icon styles: `3Arrows`, `3TrafficLights`, `3Stars`, `4Arrows`, `5Arrows`, `3Symbols`, and more.

### Conditional formatting inside loops

When used inside a `jx:each` area, conditional formatting applies to every output row:

```
Cell A1 comment:
  jx:area(lastCell="D2")
  jx:each(items="employees" var="e" lastCell="D2")

Cell C2 comment:
  jx:conditionalFormat(lastCell="C2" type="colorScale"
                       minColor="FFC7CE" maxColor="C6EFCE")
```

Every employee's salary cell gets the color scale. The formatting range expands with the data.

## Building a complete dashboard

Combine all three — chart, sparklines, and conditional formatting — for a data-rich dashboard:

### Template layout

```
A1: Department    B1: Q1     C1: Q2     D1: Q3     E1: Q4     F1: Total    G1: Trend
A2: ${d.Name}     B2: ${d.Q1} C2: ${d.Q2} D2: ${d.Q3} E2: ${d.Q4} F2: =SUM(B2:E2)

Cell A1 comment:
  jx:area(lastCell="K25")
  jx:each(items="departments" var="d" lastCell="G2")

Cell G2 comment:
  jx:sparkline(lastCell="G2" type="line" dataRange="B2:E2")

Cell F2 comment:
  jx:conditionalFormat(lastCell="F2" type="dataBar"
                       minColor="E6F2FF" maxColor="2F5496")

Cell I1 comment:
  jx:chart(lastCell="K15" type="col" title="Annual Revenue by Department"
           catRange="A2:A2" valRange="F2:F2" style="26")

Cell I16 comment:
  jx:chart(lastCell="K25" type="line" title="Quarterly Trends"
           catRange="A2:A2" valRange="B2:B2,C2:C2,D2:D2,E2:E2"
           seriesNames="Q1,Q2,Q3,Q4" legendPosition="bottom")
```

### Go code

```go
data := map[string]any{
    "departments": []map[string]any{
        {"Name": "Engineering", "Q1": 420000, "Q2": 485000, "Q3": 510000, "Q4": 530000},
        {"Name": "Sales",       "Q1": 380000, "Q2": 410000, "Q3": 450000, "Q4": 520000},
        {"Name": "Marketing",   "Q1": 180000, "Q2": 195000, "Q3": 210000, "Q4": 230000},
        {"Name": "Operations",  "Q1": 290000, "Q2": 300000, "Q3": 310000, "Q4": 325000},
        {"Name": "HR",          "Q1": 120000, "Q2": 125000, "Q3": 130000, "Q4": 135000},
    },
}

xlfill.Fill("dashboard_template.xlsx", "dashboard.xlsx", data)
```

The output has:
- Data table with quarterly figures and formula-driven totals
- Sparklines in column G showing each department's quarterly trend
- Data bars in column F showing relative total revenue
- A column chart summarizing annual revenue
- A line chart showing quarterly trends across departments

All from a single template. Zero chart code in Go.

## Tips and tricks

- **Chart sizing:** Place the chart's `lastCell` far enough from the data to avoid overlap. A good rule of thumb: leave 3-4 empty columns between data and chart.

- **Multi-series charts:** For charts with 4+ series, use `legendPosition="bottom"` — it keeps the chart area taller and more readable.

- **Choosing the right chart type:**
  - **Comparing items?** Use `bar` (horizontal) or `col` (vertical)
  - **Showing trends?** Use `line` or `area`
  - **Showing composition?** Use `pie` or `doughnut`
  - **Showing correlation?** Use `scatter`

- **Sparklines in summary rows:** You can use sparklines outside of loops too — for example, in a totals row that summarizes data from a different range.

- **Conditional formatting thresholds:** For icon sets, the thresholds are based on percentiles by default. The top third gets the green arrow, middle gets yellow, bottom gets red.

- **Combine formatting types:** Nothing stops you from putting both a color scale and data bars on the same range. Excel will layer them, though it can get visually busy.

> **Pro tip:** Design your dashboard template in Excel first with sample static data. Get the chart sizes and positions right visually, then replace the static data with `${...}` expressions. This is much faster than guessing at EMU sizes.

## What's next?

- Create data entry forms with dropdowns: **[Data Validation and Forms &rarr;](/xlfill/guides/data-validation-forms/)**
- Build financial reports with grouping and aggregation: **[Financial Reports &rarr;](/xlfill/use-cases/financial-reports/)**
- See all chart command attributes: **[jx:chart Reference &rarr;](/xlfill/commands/chart/)**
