---
title: "jx:repeat"
description: Repeat a template area a fixed number of times without needing a collection.
---

`jx:repeat` is the simplest looping command. It repeats a template area N times — no collection needed. Perfect for blank rows, fixed layouts, and padding.

## Syntax

```
jx:repeat(count="5" lastCell="C1")
jx:repeat(count="numRows" var="i" lastCell="C1")
jx:repeat(count="3" var="i" direction="RIGHT" lastCell="A1")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `count` | Yes | Number of repetitions. Can be a literal (`"5"`) or an expression (`"numRows"`) |
| `lastCell` | Yes | Bottom-right corner of the area to repeat |
| `var` | No | Variable name for the 0-based iteration index. Available as `${i}` in the area |
| `direction` | No | `"DOWN"` (default) or `"RIGHT"` |

## When to use jx:repeat vs jx:each

| Scenario | Use |
|----------|-----|
| You have data to loop over | `jx:each` |
| You need N blank rows (e.g., invoice line padding) | `jx:repeat` |
| You need numbered rows without data | `jx:repeat` with `var` |
| You need N identical columns | `jx:repeat` with `direction="RIGHT"` |

The key difference: **jx:each iterates over a collection** (employees, orders, etc.), while **jx:repeat just repeats N times**. No dummy slices needed.

## Example: Blank invoice lines

Invoices often need a fixed number of line rows regardless of how many items exist. Instead of creating a dummy slice in Go:

```
jx:repeat(count="extraLines" var="i" lastCell="D1")
```

With `${i+1}` in a cell, the output shows numbered blank rows. Your Go code simply provides:

```go
data := map[string]any{
    "extraLines": 10,
}
```

## Example: Fixed column headers

Generate columns dynamically:

```
jx:repeat(count="4" var="i" direction="RIGHT" lastCell="A1")
```

With `${i}` in cell A1, produces columns `0 | 1 | 2 | 3` across A1:D1.

## Zero count

If `count` evaluates to 0 or negative, the command produces no output — no error, no empty rows. This is useful for conditional padding:

```go
data := map[string]any{
    "extraLines": max(0, 10 - len(lineItems)),
}
```

## Common pitfalls

- **`count` is capped at 1,000,000.** Passing a larger value returns an error to prevent runaway memory use.
- **Zero count = no rows, no error.** Useful for padding (`max(0, 10 - len(items))`); confusing if you expected an error for empty input.
- **`var` is optional.** If you don't reference the index inside the area, you don't need to declare it.
- **Use `jx:each` if you have data.** `jx:repeat` is for *N identical things*; iterating over a real list is what `jx:each` is for.

## What's next?

For data-driven loops, see the workhorse command:

**[jx:each &rarr;](/xlfill/commands/each/)**
