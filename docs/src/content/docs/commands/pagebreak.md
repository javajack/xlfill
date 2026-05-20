---
title: "jx:pageBreak"
description: Insert page breaks for print-ready reports — one section per page.
---

`jx:pageBreak` inserts a horizontal page break after the command's area. Use it inside loops to ensure each section starts on a new page when printed.

## Syntax

```
jx:pageBreak(lastCell="D1")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the page break area |

## When to use jx:pageBreak

Use it when your report will be **printed or exported to PDF** and each logical section should start on a fresh page:

- One department per page in an org report
- One invoice per page in a batch print
- One student per page in a grade report
- Section breaks in a multi-part financial statement

## Example: Department per page

```
Cell A1 comment:
  jx:area(lastCell="D20")
  jx:each(items="departments" var="dept" lastCell="D5")

Cell A1 also has:
  jx:pageBreak(lastCell="D5")
```

Each department's section ends with a page break. When the user prints or uses Print Preview, each department starts on its own page.

## Example: Invoice batch

```
Cell A1 comment:
  jx:area(lastCell="F30")
  jx:each(items="invoices" var="inv" lastCell="F15")

Cell A1 also has:
  jx:pageBreak(lastCell="F15")
```

Generate 100 invoices in one workbook, each on its own page. The recipient prints the file and gets clean, separated invoices.

## Without jx:pageBreak

Without this command, Excel uses its default page breaking algorithm based on paper size and margins. For data-driven reports where section length varies, this almost always breaks pages in the wrong place — mid-table, mid-section, or splitting a header from its data.

## Common pitfalls

- **Excel may override your breaks.** If page setup (margins, scale-to-fit, orientation) doesn't allow your break to fall where you put it, Excel ignores it. Set "Fit to: 1 page wide by Many tall" in Page Setup to keep your breaks intact.
- **It's horizontal-only.** XLFill inserts a horizontal break (start a new page below). Vertical breaks (start a new page to the right) aren't supported.
- **Test in Print Preview.** Page breaks aren't visible on the normal sheet view — switch to Page Break Preview or run Print Preview to confirm placement.
- **Skipped iterations skip their breaks.** If `jx:if` excludes an iteration, the break for that iteration doesn't fire — which is what you want.

## What's next?

Auto-size columns to fit content:

**[jx:autoColWidth &rarr;](/xlfill/commands/autocolwidth/)**
