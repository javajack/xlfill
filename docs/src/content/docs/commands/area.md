---
title: "jx:area"
description: Define the working region of an XLFill template.
---

The `jx:area` command marks the rectangular region of your spreadsheet that XLFill will process. It's the first command you'll put on every template.

## Syntax

```
jx:area(lastCell="D10")
```

Place this in a **cell comment** on the top-left cell of your template region. The `lastCell` attribute is the bottom-right corner.

## Attributes

| Attribute | Description | Required |
|-----------|-------------|----------|
| `lastCell` | Bottom-right cell of the template area | Yes |

## Why it exists

XLFill needs to know which part of your spreadsheet is the template and which part is static content. The `jx:area` command draws that boundary.

- Everything **inside** the area is processed — expressions are evaluated, commands are executed
- Everything **outside** the area is left untouched in the output

This means you can have headers, footers, or instructions outside the area that won't be affected.

## Typical usage

You almost always combine `jx:area` with another command in the same cell comment:

```
jx:area(lastCell="D5")
jx:each(items="employees" var="e" lastCell="D1")
```

This says: *"The template region is A1:D5. Within that region, loop over `employees`."*

## Tips

- Only one `jx:area` per template sheet
- The area defines what gets **processed**, not what gets output — a `jx:each` inside may expand beyond the original area boundaries
- If your template has a header row above the repeating row, include the header in the area so it appears in the output

## Next command

The command you'll pair with `jx:area` on almost every template:

**[jx:each &rarr;](/commands/each/)**
