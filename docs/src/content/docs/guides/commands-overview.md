---
title: Commands Overview
description: How commands work in XLFill templates and what each one does.
---

Commands are the structural directives in your template. They go in **cell comments** and control how XLFill processes regions of your spreadsheet — looping, branching, inserting images, merging cells, and more.

## How commands work

Every command follows this pattern:

```
jx:commandName(attr1="value1" attr2="value2" lastCell="ref")
```

- The command name follows `jx:`
- Attributes are `key="value"` pairs
- Most commands require `lastCell` — the bottom-right corner of the command's area
- The cell containing the comment is the top-left corner

### Multiple commands per cell

Put multiple commands in the same cell comment, separated by newlines:

```
jx:area(lastCell="D10")
jx:each(items="employees" var="e" lastCell="D1")
```

### Automatic nesting

Commands whose areas are strictly inside another command's area are automatically treated as children. You don't need to declare nesting explicitly — XLFill figures it out from the geometry.

## The commands

Here's every command available, in order of how often you'll use them:

### Core (you'll use these on every template)

| Command | What it does | Page |
|---------|-------------|------|
| **[jx:area](/xlfill/commands/area/)** | Defines the template working region. Required on every template. | [Details &rarr;](/xlfill/commands/area/) |
| **[jx:each](/xlfill/commands/each/)** | Loops over a collection, repeating rows or columns for each item. The workhorse command. | [Details &rarr;](/xlfill/commands/each/) |
| **[jx:if](/xlfill/commands/if/)** | Conditionally shows or hides a template area. | [Details &rarr;](/xlfill/commands/if/) |

### Data iteration and layout

| Command | What it does | Page |
|---------|-------------|------|
| **[jx:repeat](/xlfill/commands/repeat/)** | Repeats an area N times without needing a collection. Great for blank rows, padding, and numbered rows. | [Details &rarr;](/xlfill/commands/repeat/) |
| **[jx:grid](/xlfill/commands/grid/)** | Fills a dynamic grid with headers and data rows. Great for pivot-style reports. | [Details &rarr;](/xlfill/commands/grid/) |
| **[jx:image](/xlfill/commands/image/)** | Inserts an image from byte data. Photos, logos, charts. | [Details &rarr;](/xlfill/commands/image/) |
| **[jx:mergeCells](/xlfill/commands/mergecells/)** | Merges cells in a range. Useful for section headers in loops. | [Details &rarr;](/xlfill/commands/mergecells/) |
| **[jx:updateCell](/xlfill/commands/updatecell/)** | Sets a single cell's value from an expression. For totals and summaries. | [Details &rarr;](/xlfill/commands/updatecell/) |
| **[jx:include](/xlfill/commands/include/)** | Inserts content from another sheet or area. Template composition for shared headers, footers, and sections. | [Details &rarr;](/xlfill/commands/include/) |

### Data presentation

| Command | What it does | Page |
|---------|-------------|------|
| **[jx:table](/xlfill/commands/table/)** | Creates a structured Excel table with auto-filter, banded rows, and total rows. | [Details &rarr;](/xlfill/commands/table/) |
| **[jx:chart](/xlfill/commands/chart/)** | Embeds bar, line, pie, and other chart types with auto-sized data ranges. | [Details &rarr;](/xlfill/commands/chart/) |
| **[jx:sparkline](/xlfill/commands/sparkline/)** | Adds mini in-cell charts (line, column, win/loss) alongside data. | [Details &rarr;](/xlfill/commands/sparkline/) |
| **[jx:conditionalFormat](/xlfill/commands/conditionalformat/)** | Applies data bars, color scales, icon sets, and cell highlighting rules. | [Details &rarr;](/xlfill/commands/conditionalformat/) |
| **[jx:dataValidation](/xlfill/commands/datavalidation/)** | Adds dropdown lists, integer/decimal constraints, and input messages. | [Details &rarr;](/xlfill/commands/datavalidation/) |
| **[jx:definedName](/xlfill/commands/definedname/)** | Creates named ranges for formulas, pivot tables, and cross-sheet references. | [Details &rarr;](/xlfill/commands/definedname/) |

### Layout and formatting

| Command | What it does | Page |
|---------|-------------|------|
| **[jx:autoRowHeight](/xlfill/commands/autorowheight/)** | Auto-fits row height after content is written. For cells with wrapped text. | [Details &rarr;](/xlfill/commands/autorowheight/) |
| **[jx:autoColWidth](/xlfill/commands/autocolwidth/)** | Auto-fits column widths to content. No manual resizing needed. | [Details &rarr;](/xlfill/commands/autocolwidth/) |
| **[jx:freezePanes](/xlfill/commands/freezepanes/)** | Freezes rows and columns so headers stay visible while scrolling. | [Details &rarr;](/xlfill/commands/freezepanes/) |
| **[jx:pageBreak](/xlfill/commands/pagebreak/)** | Inserts page breaks for print-ready reports. One section per page. | [Details &rarr;](/xlfill/commands/pagebreak/) |
| **[jx:group](/xlfill/commands/group/)** | Creates collapsible outline groups for hierarchical reports. | [Details &rarr;](/xlfill/commands/group/) |

### Protection

| Command | What it does | Page |
|---------|-------------|------|
| **[jx:protect](/xlfill/commands/protect/)** | Protects sheets from editing — lock formulas, structure, or specific cells. | [Details &rarr;](/xlfill/commands/protect/) |

## A typical template uses 2-3 commands

Don't be overwhelmed by the list of 20 commands. Most real-world templates use just `jx:area` + `jx:each`, and occasionally `jx:if`. The rest are there when you need them — charts for dashboards, tables for interactive reports, data validation for input forms, protection for compliance documents.

## What's next?

Start with the most important command — the one you'll use on every template:

**[jx:area &rarr;](/xlfill/commands/area/)**

Or jump straight to the loop command that does most of the heavy lifting:

**[jx:each &rarr;](/xlfill/commands/each/)**
