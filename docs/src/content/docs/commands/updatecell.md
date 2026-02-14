---
title: "jx:updateCell"
description: Update a single cell's value using an expression.
---

`jx:updateCell` replaces a single cell's value with the result of an expression. Use it for summary cells, totals, timestamps, or any standalone computed value.

## Syntax

```
jx:updateCell(lastCell="A1" updater="totalAmount")
```

## Attributes

| Attribute | Description | Required |
|-----------|-------------|----------|
| `lastCell` | The cell to update (typically same as the command cell) | Yes |
| `updater` | Expression whose result becomes the cell value | Yes |

## Example

```go
data := map[string]any{
    "employees":   employees,
    "totalAmount": 15750.50,
    "reportDate":  "2024-01-15",
}
```

Template cell A10 comment:
```
jx:updateCell(lastCell="A10" updater="totalAmount")
```

Cell A10 in the template can contain any placeholder text — it gets replaced with `15750.50`.

## When to use this vs. expressions

Regular `${...}` expressions work for most cases. `jx:updateCell` is useful when:
- The cell is outside a loop and you want to set it from a specific context variable
- You need the update to happen after other commands have run (it respects command ordering)

## Try it

Browse all 19 runnable examples with input templates and filled outputs on the [Examples](/xlfill/reference/examples/) page.

## Next command

Auto-fit row heights for cells with wrapped text:

**[jx:autoRowHeight &rarr;](/xlfill/commands/autorowheight/)**
