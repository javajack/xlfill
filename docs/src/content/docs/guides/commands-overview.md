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
| **[jx:area](/commands/area/)** | Defines the template working region. Required on every template. | [Details &rarr;](/commands/area/) |
| **[jx:each](/commands/each/)** | Loops over a collection, repeating rows or columns for each item. The workhorse command. | [Details &rarr;](/commands/each/) |
| **[jx:if](/commands/if/)** | Conditionally shows or hides a template area. | [Details &rarr;](/commands/if/) |

### Specialized (use when you need them)

| Command | What it does | Page |
|---------|-------------|------|
| **[jx:grid](/commands/grid/)** | Fills a dynamic grid with headers and data rows. Great for pivot-style reports. | [Details &rarr;](/commands/grid/) |
| **[jx:image](/commands/image/)** | Inserts an image from byte data. Photos, logos, charts. | [Details &rarr;](/commands/image/) |
| **[jx:mergeCells](/commands/mergecells/)** | Merges cells in a range. Useful for section headers in loops. | [Details &rarr;](/commands/mergecells/) |
| **[jx:updateCell](/commands/updatecell/)** | Sets a single cell's value from an expression. For totals and summaries. | [Details &rarr;](/commands/updatecell/) |
| **[jx:autoRowHeight](/commands/autorowheight/)** | Auto-fits row height after content is written. For cells with wrapped text. | [Details &rarr;](/commands/autorowheight/) |

## A typical template uses 2-3 commands

Don't be overwhelmed by the list. Most real-world templates use just `jx:area` + `jx:each`, and occasionally `jx:if`. The specialized commands are there when you need them, but you can build powerful reports with just the basics.

## What's next?

Start with the most important command — the one you'll use on every template:

**[jx:area &rarr;](/commands/area/)**

Or jump straight to the loop command that does most of the heavy lifting:

**[jx:each &rarr;](/commands/each/)**
