---
title: "jx:chart"
description: Embed bar, line, pie, and other Excel charts — data ranges automatically sized to your output.
---

`jx:chart` embeds an Excel chart in the output workbook. Define the chart type, data series, and placement in the template — XLFill handles the data range sizing based on actual output rows.

## Syntax

```
jx:chart(lastCell="H15" type="bar" title="Revenue by Region" catRange="A2:A2" valRange="B2:B2")
jx:chart(lastCell="H15" type="line" title="Monthly Trend" catRange="A2:A2" valRange="B2:B2,C2:C2" seriesNames="Revenue,Costs")
jx:chart(lastCell="F12" type="pie" title="Market Share" catRange="A2:A2" valRange="B2:B2")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the chart placement area |
| `type` | Yes | Chart type: `bar`, `col`, `line`, `pie`, `doughnut`, `area`, `scatter`, `radar` |
| `title` | No | Chart title text |
| `catRange` | Yes | Category axis data range (template-relative, auto-expanded) |
| `valRange` | Yes | Value data range(s), comma-separated for multiple series |
| `seriesNames` | No | Comma-separated names for each data series |
| `style` | No | Chart style index (1-48, matching Excel's chart styles) |
| `width` | No | Chart width in EMUs or cells |
| `height` | No | Chart height in EMUs or cells |
| `legendPosition` | No | Legend position: `bottom`, `left`, `right`, `top`, `none` |

## When to use jx:chart

Use it when your report data needs **visual representation** alongside the numbers:

- Bar charts for comparing categories (revenue by region, sales by product)
- Line charts for time series (monthly trends, growth rates)
- Pie charts for composition (market share, budget allocation)
- Multi-series charts for correlation (revenue vs. costs over time)

## Example: Bar chart

```
Cell A1 comment:
  jx:area(lastCell="H20")
  jx:each(items="regions" var="r" lastCell="B1")

Cell E1 comment:
  jx:chart(lastCell="H15" type="bar" title="Revenue by Region" catRange="A2:A2" valRange="B2:B2")
```

Cell A1: `${r.Name}`, Cell B1: `${r.Revenue}`

XLFill expands the data rows, then sizes the chart's data ranges to cover all output rows. If you have 10 regions, the chart references A2:A11 and B2:B11 automatically.

## Example: Multi-series line chart

```
jx:chart(lastCell="H18" type="line" title="Revenue vs Costs" catRange="A2:A2" valRange="B2:B2,C2:C2" seriesNames="Revenue,Costs" legendPosition="bottom")
```

Two series plotted on the same chart. The `seriesNames` attribute labels them in the legend.

## Example: Pie chart

```
jx:chart(lastCell="F12" type="pie" title="Market Share" catRange="A2:A2" valRange="B2:B2")
```

Clean pie chart showing the proportion of each category. Excel handles the percentage labels.

## Chart placement

The chart is anchored to the cell containing the command and extends to `lastCell`. Choose an area to the right of or below your data table so the chart doesn't overlap with data rows. The chart floats as an overlay — it doesn't shift when rows expand.

## Deferred execution

`jx:chart` uses deferred execution. Charts are defined during processing and created after all data rows are written. This is essential because the chart's data ranges must reference the final output rows, not the template rows.

## Common pitfalls

- **`catRange` and `valRange` are template-relative.** Write them for the *one-iteration* template (e.g. `A2:A2`); XLFill expands them after the loop runs. Hand-writing the expanded range like `A2:A100` confuses the engine.
- **Pie charts allow only one series.** Multiple `valRange` values on a pie chart cause Excel to ignore the extras.
- **Chart placement doesn't shift with row expansion.** The chart is anchored where you put the command. If your data table grows past where the chart starts, they'll overlap. Place charts to the right or far below.
- **`type` typos fail silently in older Excel.** Stick to the documented type names: `bar`, `col`, `line`, `pie`, `doughnut`, `area`, `scatter`, `radar`.

## What's next?

Add compact in-cell visualizations with sparklines:

**[jx:sparkline &rarr;](/xlfill/commands/sparkline/)**
