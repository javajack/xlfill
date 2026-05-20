---
title: "jx:protect"
description: Protect sheets from editing — lock formulas, structure, or specific cell ranges.
---

`jx:protect` applies Excel sheet protection to the output. Lock formulas, prevent structure changes, or allow editing only in specific unlocked cells. Essential for reports that are shared with end users who shouldn't modify certain data.

## Syntax

```
jx:protect(lastCell="A1" password="secret")
jx:protect(lastCell="A1" password="secret" sheet="true" formatCells="false" insertRows="false")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the command area |
| `password` | No | Protection password (empty = protection without password) |
| `sheet` | No | Protect the entire sheet (default: `true`) |
| `formatCells` | No | Allow formatting cells (default: `false`) |
| `formatColumns` | No | Allow formatting columns (default: `false`) |
| `formatRows` | No | Allow formatting rows (default: `false`) |
| `insertColumns` | No | Allow inserting columns (default: `false`) |
| `insertRows` | No | Allow inserting rows (default: `false`) |
| `insertHyperlinks` | No | Allow inserting hyperlinks (default: `false`) |
| `deleteColumns` | No | Allow deleting columns (default: `false`) |
| `deleteRows` | No | Allow deleting rows (default: `false`) |
| `selectLockedCells` | No | Allow selecting locked cells (default: `true`) |
| `sort` | No | Allow sorting (default: `false`) |
| `autoFilter` | No | Allow auto-filter (default: `false`) |
| `pivotTables` | No | Allow pivot table operations (default: `false`) |

## When to use jx:protect

Use it when your generated reports are **shared with non-technical users** and you need to prevent accidental or intentional modification:

- Lock formula cells in financial reports (users see results but can't edit formulas)
- Protect structure in audit reports (no row/column insertion or deletion)
- Allow only specific cells for input (combine with unlocked cell formatting)
- Prevent tampering with generated compliance or regulatory documents

## Example: Basic sheet protection

```
Cell A1 comment:
  jx:area(lastCell="D10")
  jx:each(items="items" var="item" lastCell="D1")
  jx:protect(lastCell="A1" password="report2024")
```

The entire output sheet is protected. Users can view and select cells but cannot edit anything. The password `report2024` is required to unprotect.

## Example: Allow sorting and filtering

```
jx:protect(lastCell="A1" password="secret" sort="true" autoFilter="true")
```

The sheet is protected but users can sort and filter data. Perfect for read-only reports that still need to be interactive.

## Example: Protection without password

```
jx:protect(lastCell="A1")
```

Light protection — prevents accidental edits but anyone can unprotect via the Review tab. Good for "don't touch this" signals without enforcing it.

## Unlocked cells for input

To allow editing specific cells while protecting everything else:

1. In your template, select the input cells and set their format to "Unlocked" (Format Cells > Protection > uncheck Locked)
2. Add `jx:protect` to the template

Protected cells stay locked; unlocked cells remain editable. This is the standard Excel pattern for form-style spreadsheets.

## Common pitfalls

- **It's not encryption.** Excel sheet protection is convenience-grade — a determined user with [free tools](https://www.google.com/search?q=excel+unprotect+sheet) can remove it in seconds. Don't use this to hide sensitive data.
- **Forgetting the password locks YOU out too.** XLFill doesn't store the password anywhere accessible — keep it in your secrets manager.
- **Lock formula cells via the template, not via this command.** Set "Locked" on individual cells in Excel before saving the template. `jx:protect` activates the protection; cell locking is a per-cell template attribute.
- **Default options are restrictive.** Most defaults are `false`. If you want users to be able to sort/filter your protected sheet, set those explicitly to `true`.

## What's next?

Create named ranges for formula references:

**[jx:definedName &rarr;](/xlfill/commands/definedname/)**
