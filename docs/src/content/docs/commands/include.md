---
title: "jx:include"
description: Compose templates from reusable fragments — include one template area inside another.
---

`jx:include` inserts content from another sheet or area into the current template position. Use it to build modular templates with shared headers, footers, disclaimers, or reusable report sections.

## Syntax

```
jx:include(lastCell="D3" area="CommonHeader!A1:D3")
jx:include(lastCell="D1" area="Disclaimer!A1:D5")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the include area in the host template |
| `area` | Yes | Source area reference in `Sheet!TopLeft:BottomRight` format |

## When to use jx:include

Use it when multiple templates share **common sections**:

- Company header with logo, address, and report date
- Standard disclaimer or legal footer
- Signature blocks
- Common formatting/styling sections
- Reusable sub-reports (e.g., a KPI summary widget used across dashboards)

## Example: Shared header

Your workbook has two sheets: `Report` (the main template) and `Header` (the reusable header).

**Header sheet** (A1:D3): Company logo, name, and report date — beautifully formatted.

**Report sheet:**
```
Cell A1 comment:
  jx:area(lastCell="D20")
  jx:include(lastCell="D3" area="Header!A1:D3")

Cell A4 comment:
  jx:each(items="employees" var="e" lastCell="D4")
```

The header from the `Header` sheet is inserted at A1:D3 of the output, followed by the employee loop starting at A4. All formatting from the header sheet is preserved.

## Example: Footer disclaimer

```
Cell A15 comment:
  jx:include(lastCell="D17" area="Legal!A1:D3")
```

A three-row legal disclaimer from the `Legal` sheet is inserted at the bottom of the report. Update the disclaimer in one place, and every report that includes it gets the update.

## Template composition pattern

For large organizations with many reports, the include pattern enables a **template library** approach:

1. Create a workbook with shared sheets: `Header`, `Footer`, `Disclaimer`, `KPISummary`
2. Each report template includes the shared sheets via `jx:include`
3. Update a shared section once — all reports pick it up

This eliminates copy-paste drift across dozens of report templates.

## Expressions in included areas

Expressions (`${...}`) in the included area are evaluated using the host template's data context. The included content is not a static copy — it participates fully in template processing.

## What's next?

That covers all 20 commands in XLFill. For a complete overview:

**[Commands Overview &rarr;](/xlfill/guides/commands-overview/)**

Or learn about the built-in functions available in expressions:

**[Built-in Functions &rarr;](/xlfill/guides/built-in-functions/)**
