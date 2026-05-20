---
title: "jx:freezePanes"
description: Freeze rows and columns so headers stay visible while scrolling through large reports.
---

`jx:freezePanes` freezes rows and/or columns in the output sheet so they stay visible while the user scrolls. Essential for any report with more rows than fit on screen.

## Syntax

```
jx:freezePanes(lastCell="A1" row="1" col="0")
jx:freezePanes(lastCell="A1" row="1" col="1")
jx:freezePanes(lastCell="A1" row="2" col="0")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the command area |
| `row` | No | Number of rows to freeze from the top (default: `0`) |
| `col` | No | Number of columns to freeze from the left (default: `0`) |

## When to use jx:freezePanes

Use it on **any report with headers** that users will scroll through:

- Freeze row 1 so column headers are always visible
- Freeze column A so row labels stay visible when scrolling right
- Freeze both for large matrix reports (row 1 + column A)

## Example: Freeze header row

```
Cell A1 comment:
  jx:area(lastCell="D100")
  jx:each(items="employees" var="e" lastCell="D1")
  jx:freezePanes(lastCell="A1" row="1")
```

Row 1 (the header) stays pinned while the user scrolls through hundreds of employee rows. This is the single most impactful UX improvement for large reports.

## Example: Freeze header row and first column

```
jx:freezePanes(lastCell="A1" row="1" col="1")
```

Both the top row and left column are frozen. Users can scroll right to see more columns while the row labels in column A stay visible, and scroll down while headers stay visible.

## Why not just freeze panes in the template?

You can freeze panes directly in Excel when designing your template. But if your template sheet gets copied (via `jx:each` with `multisheet`), or if the output has a different structure than the template, the freeze position might be wrong. `jx:freezePanes` applies the freeze to the output, ensuring it's correct regardless of template transformations.

## Common pitfalls

- **`row` and `col` count *frozen* rows/columns, not the freeze line position.** `row="1"` freezes row 1; the user can scroll rows 2+. `row="2"` freezes rows 1 *and* 2.
- **Both default to 0.** Setting only `row="1"` freezes the top row only (no column freeze) — that's usually what you want.
- **Multisheet templates need it per sheet.** Apply on each generated sheet (typically the template, since it gets copied).
- **Existing freeze panes on the template are overwritten.** If your template already has a freeze, `jx:freezePanes` replaces it with the engine's settings.

## What's next?

Protect sheets from unintended edits:

**[jx:protect &rarr;](/xlfill/commands/protect/)**
