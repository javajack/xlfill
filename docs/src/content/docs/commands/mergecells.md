---
title: "jx:mergeCells"
description: Merge cells in a specified range.
---

`jx:mergeCells` merges cells in a rectangular range during template processing. This is essential when you need section headers that span multiple columns inside a loop.

## Syntax

```
jx:mergeCells(lastCell="C1" cols="3" rows="1")
```

## Attributes

| Attribute | Description | Required |
|-----------|-------------|----------|
| `lastCell` | Bottom-right cell of the command area | Yes |
| `cols` | Number of columns to merge | Yes |
| `rows` | Number of rows to merge | Yes |

## Why you need this

You can merge cells in your template file directly — and for static headers, you should. But when merges happen **inside a loop** (like a department header that spans 3 columns, repeated for each department), you need `jx:mergeCells` because the merge positions change with each iteration.

## Example

A department report with a merged header per department:

```
Cell A1 comment:
  jx:area(lastCell="C5")
  jx:each(items="departments" var="dept" lastCell="C5")

Cell A1 also has:
  jx:mergeCells(lastCell="C1" cols="3" rows="1")
```

Cell A1 value: `${dept.Name}`

For each department, the department name cell spans columns A through C, creating a clean section header.

## Common pitfalls

- **Don't use this for static merges.** If your header row is always 3 columns wide and never iterates, merge it directly in Excel — it's simpler and renders identically.
- **`cols` and `rows` count from the command cell.** `cols="3"` means "merge 3 columns starting at this cell" (this cell + the next 2). Not "merge through column 3".
- **The merged range must stay inside the template area.** A merge that extends past `jx:area`'s `lastCell` produces a malformed output file.
- **Values in cells that get merged-over are lost.** Excel keeps only the top-left cell's value when cells are merged. Put `${e.Name}` in the top-left of the merge, leave the others blank.

## Try it

Download the runnable example: **template** [t11.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/input/t11.xlsx) | **output** [11_mergecells.xlsx](https://github.com/javajack/xlfill/raw/main/examples/xlfill-test/output/11_mergecells.xlsx) | [code snippet](/xlfill/reference/examples/#11-merge-cells)

## Next command

Need to set a summary or total cell?

**[jx:updateCell &rarr;](/xlfill/commands/updatecell/)**
