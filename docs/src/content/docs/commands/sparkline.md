---
title: "jx:sparkline"
description: Add mini in-cell charts — line, bar, or win/loss sparklines alongside your data.
---

`jx:sparkline` inserts Excel sparklines — tiny charts that fit inside a single cell. Use them to show trends, distributions, or win/loss patterns right next to the data they represent.

## Syntax

```
jx:sparkline(lastCell="E1" type="line" dataRange="B1:D1")
jx:sparkline(lastCell="E1" type="column" dataRange="B1:D1")
jx:sparkline(lastCell="E1" type="winLoss" dataRange="B1:D1")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the sparkline area (typically the target cell) |
| `type` | Yes | Sparkline type: `line`, `column`, `winLoss` |
| `dataRange` | Yes | The range of cells containing the sparkline data (template-relative, auto-expanded) |
| `colorSeries` | No | Series color (hex, e.g., `#4472C4`) |
| `colorNegative` | No | Negative value color |
| `colorHigh` | No | Highlight color for highest point |
| `colorLow` | No | Highlight color for lowest point |
| `colorFirst` | No | First point color |
| `colorLast` | No | Last point color |

## When to use jx:sparkline

Sparklines are perfect when you need **trend context without taking up chart space**:

- Monthly revenue trend next to each product row
- Quarterly performance history beside each employee
- Win/loss indicators for sports data or A/B test results
- Stock price mini-charts in a portfolio report

## Example: Line sparklines for monthly trends

```
Cell A1 comment:
  jx:area(lastCell="E10")
  jx:each(items="products" var="p" lastCell="E1")

Cell E1 comment:
  jx:sparkline(lastCell="E1" type="line" dataRange="B1:D1")
```

Cell A1: `${p.Name}`, B1: `${p.Jan}`, C1: `${p.Feb}`, D1: `${p.Mar}`, E1 gets a sparkline.

Each product row shows a mini line chart of its 3-month trend in column E. Users scan down the column to spot which products are trending up or down — no separate chart needed.

## Example: Column sparklines

```
jx:sparkline(lastCell="E1" type="column" dataRange="B1:D1" colorSeries="#4472C4")
```

Bar-style sparklines. Good for comparing discrete values (quarterly revenue, monthly counts).

## Example: Win/loss sparklines

```
jx:sparkline(lastCell="F1" type="winLoss" dataRange="B1:E1")
```

Binary up/down indicators. Positive values show as up bars, negative as down bars. Perfect for game results, pass/fail sequences, or month-over-month change direction.

## Deferred execution

Like charts and tables, sparklines use deferred execution. The sparkline definitions are collected during processing and applied after all rows are written, ensuring data ranges reference the correct output cells.

## Common pitfalls

- **One sparkline per cell, max.** You can't stack a line and a column sparkline in the same cell. Need both? Use adjacent cells.
- **`dataRange` must be a single row or single column.** A 2D range breaks Excel's sparkline rendering.
- **`type` is case-sensitive.** It's `line`, `column`, `winLoss` — not `Line` or `bar` (no bar sparkline; use `column`).
- **Some viewers don't render sparklines.** Excel and Google Sheets do; older LibreOffice versions may show blank cells. Test where it matters.

## What's next?

Control print layouts with page breaks:

**[jx:pageBreak &rarr;](/xlfill/commands/pagebreak/)**
