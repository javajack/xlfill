---
title: "jx:conditionalFormat"
description: Apply data bars, color scales, icon sets, and cell highlighting rules — dynamically sized to your output.
---

`jx:conditionalFormat` applies Excel conditional formatting rules to the output range. Data bars, color scales, icon sets, and value-based highlighting — all configured in the template, automatically sized to however many rows the output produces.

## Syntax

```
jx:conditionalFormat(lastCell="B1" type="dataBar" minColor="#638EC6" maxColor="#638EC6")
jx:conditionalFormat(lastCell="C1" type="colorScale" minColor="#F8696B" midColor="#FFEB84" maxColor="#63BE7B")
jx:conditionalFormat(lastCell="D1" type="iconSet" iconStyle="3Arrows")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the format area |
| `type` | Yes | Format type: `dataBar`, `colorScale`, `iconSet`, `cellIs`, `expression` |
| `minColor` | No | Color for minimum values (hex, e.g., `#F8696B`) |
| `midColor` | No | Color for midpoint values (3-color scales) |
| `maxColor` | No | Color for maximum values |
| `iconStyle` | No | Icon set style: `3Arrows`, `3TrafficLights`, `3Symbols`, `4Arrows`, `5Arrows`, `3Stars`, etc. |
| `operator` | No | For `cellIs`: `greaterThan`, `lessThan`, `between`, `equal`, etc. |
| `formula` | No | For `expression` type: a formula expression |
| `value` | No | For `cellIs`: the comparison value |
| `format` | No | Style to apply when condition matches (for `cellIs`/`expression` types) |

## When to use jx:conditionalFormat

Use it when your data needs **visual context** — when users should be able to scan a column and immediately understand the distribution:

- Data bars on revenue/sales columns for instant magnitude comparison
- Color scales on performance scores (red-yellow-green)
- Icon sets on status indicators (arrows for trend, traffic lights for health)
- Cell highlighting for threshold alerts (red background when value < target)

## Example: Data bars on a revenue column

```
Cell C1 comment:
  jx:area(lastCell="C10")
  jx:each(items="regions" var="r" lastCell="C1")

Cell C1 also has:
  jx:conditionalFormat(lastCell="C1" type="dataBar" minColor="#638EC6" maxColor="#638EC6")
```

Cell C1 value: `${r.Revenue}`

Every revenue cell gets a proportional data bar. The highest value fills 100% of the cell; others are proportional. Users can compare regions at a glance without reading numbers.

## Example: 3-color scale for scores

```
Cell B1 comment:
  jx:conditionalFormat(lastCell="B1" type="colorScale" minColor="#F8696B" midColor="#FFEB84" maxColor="#63BE7B")
```

Red for low scores, yellow for median, green for high. The classic heat map.

## Example: Icon sets for trend indicators

```
Cell D1 comment:
  jx:conditionalFormat(lastCell="D1" type="iconSet" iconStyle="3Arrows")
```

Up arrows for positive values, sideways for neutral, down for negative. Ideal for month-over-month change columns.

## Inside loops

When placed inside a `jx:each`, the conditional format rule is applied to the entire output range of the expanded column — not just one cell. XLFill tracks how many rows the loop produces and sizes the formatting rule accordingly.

## Deferred execution

Like `jx:table` and `jx:chart`, conditional formatting uses deferred execution. Rules are collected during processing and applied after all rows are written. This ensures the format range covers the correct number of output rows, even with nested loops and conditional areas.

## Common pitfalls

- **Rule order matters.** Multiple rules on the same range are evaluated top-to-bottom. The first matching `cellIs` rule wins.
- **Data bars and color scales need numeric data.** Apply them to columns that hold numbers. Mixing strings in a numeric column makes Excel show empty bars for the string rows.
- **Icon sets quantize values into buckets** (3 icons → 3 buckets, 5 icons → 5 buckets). If your data is heavily skewed, most rows land in the same bucket. Use `colorScale` for smoother gradients.
- **Some viewers render conditional formats differently.** LibreOffice's data bars are subtly different from Excel's; Google Sheets supports a smaller subset. Test in the viewer your users use.

## What's next?

Group rows into collapsible outline sections:

**[jx:group &rarr;](/xlfill/commands/group/)**
