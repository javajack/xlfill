---
title: "jx:group"
description: Create collapsible outline groups for rows — perfect for hierarchical reports with expandable sections.
---

`jx:group` creates Excel outline grouping on the output rows, giving users the ability to expand and collapse sections. Ideal for hierarchical reports, department breakdowns, and any template where detail rows should be hideable.

## Syntax

```
jx:group(lastCell="D1" collapsed="false")
jx:group(lastCell="D1" collapsed="true")
```

## Attributes

| Attribute | Required | Description |
|-----------|----------|-------------|
| `lastCell` | Yes | Bottom-right cell of the group area |
| `collapsed` | No | Whether the group starts collapsed (default: `false`) |

## When to use jx:group

Use it when your report has **hierarchical structure** that users should be able to drill into:

- Department summary with expandable employee details
- Quarterly totals with collapsible monthly breakdowns
- Category headers with hidden line items
- Executive summary that expands into full detail

## Example: Departments with expandable employees

```
Cell A1 comment:
  jx:area(lastCell="D20")
  jx:each(items="departments" var="dept" lastCell="D5")

Cell A2 comment:
  jx:each(items="dept.Employees" var="e" lastCell="D2")
  jx:group(lastCell="D2" collapsed="false")
```

Row 1 contains `${dept.Name}` (the department header). Rows 2+ contain `${e.Name}`, `${e.Title}`, etc. The `jx:group` wraps the employee rows in an outline group. In the output, users see the [+]/[-] controls in the row gutter to expand/collapse each department's employees.

## Example: Collapsed by default

For executive summaries where detail should be hidden initially:

```
jx:group(lastCell="D1" collapsed="true")
```

The group starts collapsed. Users see only the summary rows and can click to expand when they need details.

## Nested groups

Excel supports up to 8 levels of outline nesting. Nest `jx:group` commands to create multi-level hierarchies:

```
Cell A1 comment:
  jx:each(items="regions" var="region" lastCell="D10")

Cell A2 comment:
  jx:each(items="region.Departments" var="dept" lastCell="D5")
  jx:group(lastCell="D5")

Cell A3 comment:
  jx:each(items="dept.Employees" var="e" lastCell="D3")
  jx:group(lastCell="D3")
```

Region > Department > Employee — each level independently collapsible.

## Deferred execution

`jx:group` uses deferred execution. The group ranges are recorded during template processing and applied after all rows are written, ensuring correct row references even when loops expand the output.

## What's next?

Add charts to visualize your data:

**[jx:chart &rarr;](/xlfill/commands/chart/)**
