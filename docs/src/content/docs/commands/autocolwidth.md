---
title: "jx:autoColWidth"
description: Auto-fit column widths to content — no manual resizing needed.
---

`jx:autoColWidth` adjusts column widths after content is written so that all text fits without truncation. The column counterpart to `jx:autoRowHeight`.

## Syntax

```
jx:autoColWidth(lastCell="D1")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the area whose columns should be auto-sized |

## When to use jx:autoColWidth

Use it when your data has **variable-length content** and you don't want users to manually resize columns:

- Name columns where some entries are short ("Al") and others long ("Christopher Alexander")
- Description or notes columns
- Any column where the template's fixed width might truncate output data

## Example: Auto-size all columns

```
Cell A1 comment:
  jx:area(lastCell="D10")
  jx:each(items="employees" var="e" lastCell="D1")
  jx:autoColWidth(lastCell="D1")
```

After all rows are written, columns A through D are resized to fit the widest content in each column. The header row and all data rows are considered.

## Template width as minimum

The auto-sizing uses the template column width as a minimum. If you set column A to 15 characters wide in your template, `jx:autoColWidth` will never shrink it below that — it only widens columns when content requires it.

## Combining with jx:autoRowHeight

For fully adaptive layouts, use both:

```
Cell A1 comment:
  jx:autoColWidth(lastCell="D1")
  jx:autoRowHeight(lastCell="D1")
```

Columns expand to fit the widest value; rows expand to fit wrapped text. The output is always readable regardless of data length.

## What's next?

Freeze header rows and columns for easier scrolling:

**[jx:freezePanes &rarr;](/xlfill/commands/freezepanes/)**
