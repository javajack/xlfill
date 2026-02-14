---
title: "jx:autoRowHeight"
description: Auto-fit row height after content is written.
---

`jx:autoRowHeight` adjusts the row height after content is written to fit wrapped text. Without this, rows with long text content may appear truncated.

## Syntax

```
jx:autoRowHeight(lastCell="C1")
```

## Attributes

| Attribute | Description | Required |
|-----------|-------------|----------|
| `lastCell` | Bottom-right cell of the command area | Yes |

## When to use this

Use it when your template cells have **word wrap enabled** and the data may contain text of varying lengths. Without auto-fitting, Excel uses the row height from the template, which may be too short for longer content.

## Example

Template cell A1 comment:
```
jx:autoRowHeight(lastCell="C1")
```

After XLFill writes the cell content, it recalculates the row height so all wrapped text is visible.

## Inside loops

Combine with `jx:each` to auto-fit every generated row:

```
Cell A1 comment:
  jx:area(lastCell="C1")
  jx:each(items="items" var="e" lastCell="C1")

Cell A1 also has:
  jx:autoRowHeight(lastCell="C1")
```

Each row gets its height adjusted based on its content.

---

That's every command in XLFill. You now know the full template language. For most reports, you'll use `jx:area` + `jx:each` and occasionally `jx:if`. The rest are there when you need them.

## What's next?

Learn how formulas in your templates are automatically expanded when rows are inserted:

**[Formulas &rarr;](/xlfill/guides/formulas/)**
